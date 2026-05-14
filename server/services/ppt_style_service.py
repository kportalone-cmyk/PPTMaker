"""PPT 스타일(파워포인트 스타일) 비즈니스 로직

- 샘플 이미지 Vision 분석 → design_tokens 빈 슬롯 자동 채움
- 빌더가 즉시 사용 가능한 정규화된 스킬 JSON 빌드 (font_refs → 실제 폰트 메타 join 등)

Vision 분석은 llm_service.analyze_image_content() 를 재사용한다.
프롬프트는 PPT 디자인 분석 전용으로 본 서비스에서 직접 주입한다.
"""

import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Optional

from bson import ObjectId

from config import settings
from services.mongo_service import get_db
from services import llm_service


# ============ Vision 분석 전용 프롬프트 ============

_VISION_PPT_PROMPT = (
    "당신은 기업용 프레젠테이션 디자인 분석가입니다.\n"
    "주어진 슬라이드/표지 이미지를 분석하여 디자인 스타일을 추출하세요.\n\n"
    "다음 JSON 스키마로만 응답하세요 (코드 블록/설명 없이 순수 JSON):\n"
    "{\n"
    '  "primary_color": "#RRGGBB",            // 주조색 (가장 비중 큰 브랜드 컬러)\n'
    '  "secondary_colors": ["#RRGGBB", ...],  // 보조색 2~5개 (light/ink/grey/line/darker 후보)\n'
    '  "font_style": "sans-serif|serif|display|geometric|humanist",\n'
    '  "layout": "한 줄 요약 - 그리드/여백/시각적 특징",\n'
    '  "spacing_density": "tight|comfortable|spacious",   // 여백/요소 간격 인상\n'
    '  "h1_weight": "regular|medium|bold|black",          // 제목 굵기\n'
    '  "body_weight": "regular|medium|bold",              // 본문 굵기\n'
    '  "stat_weight": "regular|medium|bold|black",        // 큰 숫자(스탯) 굵기\n'
    '  "title_letter_spacing": "tight|normal|wide",       // 제목 자간\n'
    '  "title_line_height": "tight|normal|loose",         // 제목 행간\n'
    '  "body_line_height": "tight|normal|loose",          // 본문 행간\n'
    '  "corner_style": "sharp|rounded|pill",              // 카드/도형 모서리\n'
    '  "image_aspect": "16:9|4:3|1:1|free",               // 주된 이미지 비율\n'
    '  "grid_columns": 6,                                 // 6 | 12 | 24 중 가장 가까운 값\n'
    '  "compositions": ["..."]                            // 슬라이드 합성 패턴 태그\n'
    "}\n\n"
    "compositions 후보 (해당되는 것만 선택, 최대 4개):\n"
    "  split_60_40 · centered_hero · grid_2col_cards · grid_3col_cards · grid_2x2 ·\n"
    "  full_bleed_photo · sidebar_left · sidebar_right · big_stat · timeline · quote · comparison\n\n"
    "주의:\n"
    "- 색상은 반드시 #RRGGBB 6자리 hex (대문자) 형식으로 출력\n"
    "- 차트/이미지 영역의 색은 제외하고, 텍스트·배경·강조 막대·키 컬러만 추출\n"
    "- 한국어 텍스트가 있어도 모든 키워드는 영문 enum 그대로 사용\n"
    "- 확신이 없는 필드는 가장 가까운 보수적 기본값으로 답해도 됨 (null 금지)\n"
)


# ============ 패턴 추출 전용 프롬프트 (M8) ============

_VISION_PATTERN_EXTRACT_PROMPT = (
    "이 이미지는 한 PPT 디자인 시리즈의 여러 슬라이드 레이아웃이 한 장에 담겨 있는 레퍼런스입니다.\n"
    "이미지 안에서 식별 가능한 각 슬라이드를 구조화된 JSON 패턴으로 추출하세요.\n"
    "추출된 패턴은 이후 새 outline 텍스트(제목/카드/아이콘/이미지)를 채워 슬라이드를 자동 생성하는 데 사용됩니다.\n\n"
    "## 각 패턴 스키마 (모든 필드 필수, 다만 해당없으면 빈 값/0/[] 허용)\n"
    "{\n"
    '  "label": "왼쪽 인물 + 오른쪽 본문 (히어로)",\n'
    '  "regions": [\n'
    "    {\n"
    '      "role": "background|image|title|subtitle|body|card_title|card_desc|stat|unit|icon|decoration|bullet|number_badge|divider|page_indicator|chip|logo|quote",\n'
    '      "x_pct": 0, "y_pct": 0, "w_pct": 50, "h_pct": 100,\n'
    '      "text_align": "left|center|right",\n'
    '      "font_scale": "xs|sm|md|lg|xl|hero",\n'
    '      "font_weight": "regular|medium|bold",\n'
    '      "text_color_role": "ink|grey|primary|white|darker",\n'
    '      "group_index": 0,\n'
    '      "icon_category": "stat|bullet|feature|decoration",\n'
    '      "image_role": "hero|thumbnail|photo|icon-illustration|background|none",\n'
    '      "shape": "rectangle|rounded_rect|ellipse|circle|line|none",\n'
    '      "fill_color_role": "primary|light|white|ink|grey|line|darker|none",\n'
    '      "opacity": 1.0,\n'
    '      "approx_color": "#hex"\n'
    "    }\n"
    "  ],\n"
    '  "suitable_for_outline_types": ["title|toc|section|content|closing"],\n'
    '  "card_count": 0,\n'
    '  "visual_keywords": ["hero","split","grid"]\n'
    "}\n\n"
    "## 핵심 규칙\n"
    "- 한 이미지에서 보이는 슬라이드를 최대 12개까지 추출 (작은 썸네일도 식별되면 포함).\n"
    "- 좌표는 슬라이드 한 장 기준 백분율 (0-100, x+w<=100, y+h<=100).\n"
    "- 본문 카드가 N개 보이면 card_title region 을 N개 추출하고 group_index 로 묶으세요 (card_desc/icon 도 같은 group_index).\n"
    "- card_count = card_title region 의 개수.\n"
    "- 텍스트는 font_scale 로 상대 크기 (body=md 기준, 제목 lg~hero, 라벨 xs~sm).\n"
    "- 텍스트 색은 design_tokens 의 토큰 이름(text_color_role)으로 매핑.\n"
    "- 도형/장식은 shape + fill_color_role + opacity.\n"
    "- background role 의 region 1개에는 approx_color 또는 fill_color_role 을 채우세요.\n"
    "- 불명확하거나 해당없는 선택 필드는 생략 가능. 필수는 role/x_pct/y_pct/w_pct/h_pct.\n\n"
    "## 응답 형식\n"
    "반드시 ```json 코드 블록 1개 안에 다음 형태로 출력 (그 외 설명/주석 금지):\n"
    '{"patterns": [패턴1, 패턴2, ...]}\n'
    "JSON 이외 텍스트는 절대 포함하지 마세요.\n"
)


_HEX_RE = re.compile(r"#?([0-9A-Fa-f]{6})\b")


# ============ 의미값 → 수치 매핑 (spacing / typography / archetypes) ============

# spacing_density → 구체 pt 값 (16:9 슬라이드 960x540pt 기준).
# 빌더가 직접 참고할 수 있도록 density 의미값도 함께 보존한다.
_DENSITY_TO_SPACING: dict = {
    "tight": {
        "base_unit":      4,
        "slide_margin_x": 28,
        "slide_margin_y": 20,
        "section_gap":    16,
        "element_gap":    8,
        "card_padding":   10,
    },
    "comfortable": {
        "base_unit":      8,
        "slide_margin_x": 40,
        "slide_margin_y": 32,
        "section_gap":    24,
        "element_gap":    12,
        "card_padding":   16,
    },
    "spacious": {
        "base_unit":      8,
        "slide_margin_x": 56,
        "slide_margin_y": 44,
        "section_gap":    36,
        "element_gap":    18,
        "card_padding":   22,
    },
}

# letter-spacing 의미값 → 퍼센트(em 대비) 정수
_LETTER_SPACING_PCT: dict = {"tight": -2, "normal": 0, "wide": 5}

# line-height 의미값 → 실제 배율 (제목/본문 별로 차이를 둠)
_LINE_HEIGHT_H1: dict = {"tight": 1.05, "normal": 1.15, "loose": 1.3}
_LINE_HEIGHT_BODY: dict = {"tight": 1.25, "normal": 1.45, "loose": 1.65}

# corner_style → corner_radius 토큰
_CORNER_TO_RADIUS: dict = {"sharp": "sharp", "rounded": "md", "pill": "pill"}


# 폰트 분류 휴리스틱 — family/name 문자열로 typography.font_style 추정.
# 등록된 폰트(섹션 5)의 family/name 만 보고 키워드 매칭. 명확하지 않으면 None.
# 순서가 중요: serif 판별 전에 "sans serif" 가 먼저 매칭되도록 sans-serif 그룹을 위에 둔다.
def _classify_font_style(*labels: str) -> Optional[str]:
    """폰트 family/name 후보 문자열들 중 첫 번째로 분류 가능한 결과를 반환."""
    for raw in labels:
        if not raw or not isinstance(raw, str):
            continue
        s = raw.lower()

        # 1) sans-serif 계열 — sans 키워드 먼저 (serif 보다 우선)
        sans_keywords = (
            "sans", "gothic", "고딕", "돋움", "굴림", "맑은", "맑은고딕",
            "noto sans", "노토산스", "noto-sans",
            "pretendard", "프리텐다드", "suite", "spoqa",
            "nanum gothic", "나눔고딕", "ibm plex sans", "inter", "roboto",
            "helvetica", "arial",
        )
        if any(k in s for k in sans_keywords):
            return "sans-serif"

        # 2) serif 계열
        serif_keywords = (
            "serif", "myungjo", "명조", "본명조", "batang", "바탕",
            "noto serif", "노토세리프", "nanum myeongjo", "나눔명조",
            "times", "georgia", "garamond",
        )
        if any(k in s for k in serif_keywords):
            return "serif"

        # 3) display / geometric / humanist — 폰트 카테고리 단어가 이름에 명시된 경우
        if "display" in s:
            return "display"
        if "geometric" in s or "futura" in s or "avenir" in s:
            return "geometric"
        if "humanist" in s or "calibri" in s or "verdana" in s:
            return "humanist"

    return None


async def _derive_font_style_from_design_tokens(design_tokens: dict) -> Optional[str]:
    """design_tokens.fonts.title_font_id / body_font_id 로부터 typography.font_style 도출.

    제목 폰트를 우선 참조, 실패 시 본문 폰트로 폴백. 둘 다 없거나 분류 실패면 None.
    """
    fonts_meta = (design_tokens or {}).get("fonts") or {}
    candidate_ids = [
        fonts_meta.get("title_font_id"),
        fonts_meta.get("body_font_id"),
    ]
    if not any(candidate_ids):
        return None

    db = get_db()
    for fid in candidate_ids:
        if not fid:
            continue
        try:
            oid = ObjectId(fid)
        except Exception:
            continue
        doc = await db.fonts.find_one({"_id": oid})
        if not doc:
            continue
        result = _classify_font_style(doc.get("family") or "", doc.get("name") or "")
        if result:
            return result
    return None


def _mode(values: list, allowed: list | None = None):
    """리스트에서 최빈값 반환. 동률 시 첫 등장 우선. 빈 리스트면 None."""
    counts: dict = {}
    order: list = []
    for v in values:
        if v is None or v == "":
            continue
        if allowed is not None and v not in allowed:
            continue
        if v not in counts:
            order.append(v)
            counts[v] = 0
        counts[v] += 1
    if not counts:
        return None
    best = max(order, key=lambda k: (counts[k], -order.index(k)))
    return best


def _aggregate_design_hints(parsed_responses: list[dict]) -> dict:
    """샘플별 Vision 응답에서 spacing / typography / archetypes 의미값을 집계.

    Returns:
        {
          "spacing":    { density, base_unit, slide_margin_x, ... },
          "typography": { font_style, h1_weight, ..., h1_line_height: float, ... },
          "archetypes": { grid_columns, column_gutter, corner_radius, image_aspect, compositions: [...] },
        }
    각 필드는 집계가 불가능하면 생략된다 (자동 채움 시 빈 슬롯만 덮어쓰는 정책).
    """
    if not parsed_responses:
        return {"spacing": {}, "typography": {}, "archetypes": {}}

    density_list   = [p.get("spacing_density") for p in parsed_responses]
    h1_weight      = [p.get("h1_weight")       for p in parsed_responses]
    body_weight    = [p.get("body_weight")     for p in parsed_responses]
    stat_weight    = [p.get("stat_weight")     for p in parsed_responses]
    font_style     = [p.get("font_style")      for p in parsed_responses]
    title_letter   = [p.get("title_letter_spacing") for p in parsed_responses]
    title_line     = [p.get("title_line_height")    for p in parsed_responses]
    body_line      = [p.get("body_line_height")     for p in parsed_responses]
    corner_style   = [p.get("corner_style")    for p in parsed_responses]
    image_aspect   = [p.get("image_aspect")    for p in parsed_responses]
    grid_cols      = [p.get("grid_columns")    for p in parsed_responses]

    spacing: dict = {}
    density = _mode(density_list, ["tight", "comfortable", "spacious"])
    if density:
        spacing["density"] = density
        spacing.update(_DENSITY_TO_SPACING[density])

    typography: dict = {}
    fs = _mode(font_style, ["sans-serif", "serif", "display", "geometric", "humanist"])
    if fs:
        typography["font_style"] = fs
    h1w = _mode(h1_weight, ["regular", "medium", "bold", "black"])
    if h1w:
        typography["h1_weight"] = h1w
    bw = _mode(body_weight, ["regular", "medium", "bold"])
    if bw:
        typography["body_weight"] = bw
    sw = _mode(stat_weight, ["regular", "medium", "bold", "black"])
    if sw:
        typography["stat_weight"] = sw
    # h2_weight 는 별도 추출하지 않고 h1/body 의 중간값으로 추정
    if h1w and bw:
        order = ["regular", "medium", "bold", "black"]
        try:
            mid_idx = (order.index(h1w) + order.index(bw)) // 2
            typography["h2_weight"] = order[mid_idx]
        except Exception:
            pass
    tl = _mode(title_letter, ["tight", "normal", "wide"])
    if tl:
        typography["h1_letter_spacing"] = _LETTER_SPACING_PCT[tl]
    # body letter-spacing 은 별도 추출 안 함 — 0 으로 둠
    tlh = _mode(title_line, ["tight", "normal", "loose"])
    if tlh:
        typography["h1_line_height"] = _LINE_HEIGHT_H1[tlh]
    blh = _mode(body_line, ["tight", "normal", "loose"])
    if blh:
        typography["body_line_height"] = _LINE_HEIGHT_BODY[blh]

    archetypes: dict = {}
    # grid_columns: 6/12/24 후보 중 가장 가까운 값
    nums: list[int] = []
    for v in grid_cols:
        try:
            n = int(v)
            if n > 0:
                nums.append(min([6, 12, 24], key=lambda c: abs(c - n)))
        except Exception:
            continue
    gc = _mode(nums, [6, 12, 24])
    if gc:
        archetypes["grid_columns"] = gc
        # gutter 는 컬럼 수에 따라 보수적 기본값
        archetypes["column_gutter"] = 24 if gc == 6 else (16 if gc == 12 else 12)
    cs = _mode(corner_style, ["sharp", "rounded", "pill"])
    if cs:
        archetypes["corner_radius"] = _CORNER_TO_RADIUS[cs]
    ia = _mode(image_aspect, ["16:9", "4:3", "1:1", "free"])
    if ia:
        archetypes["image_aspect"] = ia

    # compositions: 모든 응답의 합집합 (등장 횟수 내림차순)
    allowed_comp = {
        "split_60_40", "centered_hero",
        "grid_2col_cards", "grid_3col_cards", "grid_2x2",
        "full_bleed_photo", "sidebar_left", "sidebar_right",
        "big_stat", "timeline", "quote", "comparison",
    }
    comp_counts: dict = {}
    for p in parsed_responses:
        items = p.get("compositions") or []
        if isinstance(items, str):
            items = [items]
        for it in items:
            if isinstance(it, str) and it in allowed_comp:
                comp_counts[it] = comp_counts.get(it, 0) + 1
    if comp_counts:
        archetypes["compositions"] = sorted(
            comp_counts.keys(), key=lambda k: -comp_counts[k]
        )

    return {"spacing": spacing, "typography": typography, "archetypes": archetypes}


def _normalize_hex(value: str) -> Optional[str]:
    """hex 컬러 정규화 - '#RRGGBB' 대문자 반환, 실패 시 None"""
    if not isinstance(value, str):
        return None
    m = _HEX_RE.search(value.strip())
    if not m:
        return None
    return "#" + m.group(1).upper()


def _extract_json_block(text: str) -> Optional[dict]:
    """LLM 응답에서 JSON 객체 추출 (코드펜스 / 잡음 포함 케이스 대응)"""
    if not text:
        return None
    # 1) 통째로 시도
    try:
        return json.loads(text)
    except Exception:
        pass
    # 2) ```json ... ``` 코드 펜스 (greedy 매칭으로 중첩 구조 허용)
    fenced = re.search(r"```(?:json)?\s*(\{[\s\S]*\})\s*```", text, flags=re.DOTALL)
    if fenced:
        try:
            return json.loads(fenced.group(1))
        except Exception:
            pass
    # 3) ```json ... ``` non-greedy 도 폴백
    fenced2 = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, flags=re.DOTALL)
    if fenced2:
        try:
            return json.loads(fenced2.group(1))
        except Exception:
            pass
    # 4) 본문 안 첫 번째 { ... } 매칭 (가장 바깥 중괄호 균형)
    start = text.find("{")
    if start >= 0:
        depth = 0
        in_str = False
        esc = False
        for i in range(start, len(text)):
            ch = text[i]
            if esc:
                esc = False
                continue
            if ch == "\\":
                esc = True
                continue
            if ch == '"':
                in_str = not in_str
                continue
            if in_str:
                continue
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(text[start:i + 1])
                    except Exception:
                        break
    return None


# ============ 파일 경로 ============

def get_sample_dir(style_id: str) -> Path:
    """샘플 이미지 저장 디렉토리 (settings.UPLOAD_DIR/ppt_style_samples/<style_id>)"""
    return Path(settings.UPLOAD_DIR) / "ppt_style_samples" / style_id


def sample_url(style_id: str, filename: str) -> str:
    """샘플 이미지 정적 URL"""
    return f"/uploads/ppt_style_samples/{style_id}/{filename}"


# ============ URL → 로컬 경로 변환 ============

def _url_to_local_path(url: str) -> Optional[Path]:
    """샘플 URL(`/uploads/...`) 을 실제 디스크 경로로 변환.

    `analyze_samples_with_vision` 와 동일한 규칙을 별도 함수로 추출하여
    `extract_patterns_from_samples` 와 공유한다.
    """
    if not isinstance(url, str) or not url:
        return None
    rel = url.lstrip("/")
    if rel.startswith("uploads/"):
        rel = rel[len("uploads/"):]
    uploads_root = Path(settings.UPLOAD_DIR)
    fp = uploads_root / rel
    return fp if fp.exists() else None


# ============ Vision 분석 ============

async def analyze_samples_with_vision(style_id: str) -> dict:
    """등록된 샘플 이미지들을 Vision 분석하여 vision_analysis 필드 갱신.

    1. 도큐먼트 로드 → sample_image_refs 읽기
    2. 각 파일에 대해 llm_service.analyze_image_content() (PPT 전용 프롬프트)
    3. 결과 파싱 → primary/secondary/font/layout 합치기
    4. extracted_colors 로 design_tokens.colors 빈 슬롯 자동 채움 (사용자 수동값 보존)
    5. vision_analysis 및 design_tokens 업데이트 후 최종 도큐먼트 반환
    """
    db = get_db()
    try:
        oid = ObjectId(style_id)
    except Exception:
        return {}

    doc = await db.ppt_styles.find_one({"_id": oid})
    if not doc:
        return {}

    sample_refs = doc.get("sample_image_refs", []) or []

    extracted_colors: list[str] = []
    secondary_colors: list[str] = []
    detected_fonts: list[str] = []
    layout_hints: list[str] = []
    raw_responses: list[dict] = []
    parsed_for_hints: list[dict] = []   # spacing/typography/archetypes 집계용 원시 응답 모음

    uploads_root = Path(settings.UPLOAD_DIR)

    for ref in sample_refs:
        url = ref.get("url") or ""
        original_filename = ref.get("original_filename") or ""
        # URL → 실제 파일 경로
        rel = url.lstrip("/")
        if rel.startswith("uploads/"):
            rel = rel[len("uploads/"):]
        file_path = uploads_root / rel
        if not file_path.exists():
            continue

        text = await llm_service.analyze_image_content(
            str(file_path),
            original_filename=original_filename,
            prompt=_VISION_PPT_PROMPT,
        )
        if not text:
            continue

        parsed = _extract_json_block(text)
        raw_responses.append({
            "file_id": ref.get("file_id"),
            "original_filename": original_filename,
            "text": text,
            "parsed": parsed,
        })
        if not parsed:
            continue
        parsed_for_hints.append(parsed)

        p = _normalize_hex(parsed.get("primary_color", ""))
        if p and p not in extracted_colors:
            extracted_colors.append(p)
        for s in parsed.get("secondary_colors") or []:
            ns = _normalize_hex(s)
            if ns and ns not in secondary_colors and ns not in extracted_colors:
                secondary_colors.append(ns)
        fs = parsed.get("font_style")
        if isinstance(fs, str) and fs and fs not in detected_fonts:
            detected_fonts.append(fs.strip())
        lay = parsed.get("layout")
        if isinstance(lay, str) and lay:
            layout_hints.append(lay.strip())

    # 합쳐서 색상 풀 구성 (primary > secondary)
    all_colors = list(extracted_colors)
    for c in secondary_colors:
        if c not in all_colors:
            all_colors.append(c)

    vision_analysis = {
        "extracted_colors": all_colors,
        "primary_colors": extracted_colors,
        "secondary_colors": secondary_colors,
        "detected_fonts": detected_fonts,
        "layout_hints": layout_hints,
        "raw_responses": raw_responses,
        "analyzed_at": datetime.utcnow(),
    }

    # design_tokens.colors / sizes 자동 채움.
    #
    # 정책 (디자인 토큰은 기본값이 없어야 한다는 사용자 요구):
    #   - 비어 있는 슬롯은 추출된 색상으로 채운다.
    #   - 사용자가 수동으로 입력한 값(어떤 값이든)은 보존한다 — 단, 분석 결과와
    #     "기본 팔레트" 값이 우연히 일치할 때는 보존하지 않고 분석값으로 갱신한다.
    #   - 기본 팔레트(DEFAULT_COLORS) 와 동일한 값도 "손대지 않음" 으로 간주해 덮어쓴다.
    from models.ppt_style import DEFAULT_COLORS, DEFAULT_FONT_SIZES

    design_tokens = doc.get("design_tokens") or {}
    colors = (design_tokens.get("colors") or {}).copy()
    slot_order = ["primary", "light", "ink", "grey", "line", "darker", "white"]
    pool = list(all_colors)
    for slot in slot_order:
        current = colors.get(slot)
        # 비었거나, 기본 팔레트와 동일하면 분석값으로 채움.
        if (not current) or (current == DEFAULT_COLORS.get(slot)):
            if pool:
                colors[slot] = pool.pop(0)
    # white 슬롯이 비어 있으면 기본값(#FFFFFF) 으로 보완 (사용자가 흰색을 의도하지 않을 일은 거의 없음)
    if not colors.get("white"):
        colors["white"] = DEFAULT_COLORS["white"]
    design_tokens["colors"] = colors

    # fonts.sizes 자동 채움 — Vision 으로는 pt 추정이 어려우므로, 비어 있을 때만
    # DEFAULT_FONT_SIZES 로 채운다 (전부 비어 있으면 한 번에 채워 들어가게 됨).
    fonts = design_tokens.get("fonts") or {}
    sizes = (fonts.get("sizes") or {}).copy()
    for k, v in DEFAULT_FONT_SIZES.items():
        if not sizes.get(k):
            sizes[k] = v
    fonts["sizes"] = sizes
    # title_font_id / body_font_id 는 그대로 유지 (사용자 선택)
    design_tokens["fonts"] = fonts

    # spacing / typography / archetypes — 빈 슬롯만 자동 채움 (사용자 수동값 보존).
    # 분석 결과가 비어 있으면 기본값(DEFAULT_*) 도 함께 보완해 슬롯이 완성된 상태로 둔다.
    from models.ppt_style import (
        DEFAULT_SPACING,
        DEFAULT_TYPOGRAPHY,
        DEFAULT_ARCHETYPES,
    )
    hints = _aggregate_design_hints(parsed_for_hints)

    def _merge_slot(existing: dict | None, hint: dict, defaults: dict) -> dict:
        out = dict(existing or {})
        for k, v in defaults.items():
            cur = out.get(k)
            if cur is None or cur == "" or cur == []:
                # hint 값이 있으면 우선, 없으면 default 로 보완
                out[k] = hint[k] if k in hint else v
        # compositions 같은 list 슬롯에 hint 가 새 값을 들고 왔으면 우선 적용
        for k, v in hint.items():
            if k not in defaults:
                continue
            cur = out.get(k)
            if cur is None or cur == "" or cur == [] or cur == defaults[k]:
                out[k] = v
        return out

    design_tokens["spacing"]    = _merge_slot(design_tokens.get("spacing"),    hints["spacing"],    DEFAULT_SPACING)
    design_tokens["typography"] = _merge_slot(design_tokens.get("typography"), hints["typography"], DEFAULT_TYPOGRAPHY)
    design_tokens["archetypes"] = _merge_slot(design_tokens.get("archetypes"), hints["archetypes"], DEFAULT_ARCHETYPES)

    # typography.font_style 은 섹션 5 (등록된 제목/본문 폰트) 가 진실의 소스다.
    # 폰트가 선택돼 있으면 그 family/name 으로 분류한 값으로 비전 추정값을 덮어쓴다.
    derived_fs = await _derive_font_style_from_design_tokens(design_tokens)
    if derived_fs:
        design_tokens["typography"]["font_style"] = derived_fs

    now = datetime.utcnow()
    await db.ppt_styles.update_one(
        {"_id": oid},
        {"$set": {
            "vision_analysis": vision_analysis,
            "design_tokens": design_tokens,
            "updated_at": now,
        }},
    )

    updated = await db.ppt_styles.find_one({"_id": oid})
    if updated and "_id" in updated:
        updated["_id"] = str(updated["_id"])
    return updated or {}


# ============ 패턴 추출 (M8) ============

async def extract_patterns_from_samples(style_id: str) -> dict:
    """샘플 이미지에서 슬라이드 레이아웃 패턴들을 자동 추출.

    각 `sample_image_refs` 의 이미지를 Vision LLM 으로 분석해
    이미지에 보이는 개별 슬라이드 레이아웃 N개를 구조화 패턴으로 추출한다.

    ppt_styles 도큐먼트의 `extracted_patterns` 필드에 저장하고 결과를 반환.

    Returns:
        {
          "extracted_patterns": [...],
          "by_sample": [{"sample_index": int, "url": str, "patterns": [...]}],
          "total_patterns": int
        }
    """
    db = get_db()
    try:
        oid = ObjectId(style_id)
    except Exception:
        return {
            "extracted_patterns": [],
            "by_sample": [],
            "total_patterns": 0,
        }

    doc = await db.ppt_styles.find_one({"_id": oid})
    if not doc:
        return {
            "extracted_patterns": [],
            "by_sample": [],
            "total_patterns": 0,
        }

    sample_refs = doc.get("sample_image_refs", []) or []

    extracted_patterns: list[dict] = []
    by_sample: list[dict] = []
    extraction_debug: list[dict] = []  # 원본 응답 + 파싱 상태 (운영 진단용)

    print(f"[PatternExtract] style_id={style_id} samples={len(sample_refs)}")

    for sample_index, ref in enumerate(sample_refs):
        url = ref.get("url") or ""
        original_filename = ref.get("original_filename") or ""

        file_path = _url_to_local_path(url)
        sample_entry = {
            "sample_index": sample_index,
            "url": url,
            "patterns": [],
        }
        debug_entry: dict = {
            "sample_index": sample_index,
            "url": url,
            "original_filename": original_filename,
            "status": "ok",
            "raw_text_len": 0,
            "raw_text_preview": "",
            "raw_text_tail": "",
            "patterns_found": 0,
        }
        if file_path is None:
            debug_entry["status"] = "file_not_found"
            print(f"[PatternExtract] sample[{sample_index}] file not found: url={url}")
            by_sample.append(sample_entry)
            extraction_debug.append(debug_entry)
            continue

        # 패턴 추출 LLM 호출 (max_tokens 충분히)
        text = await llm_service.analyze_image_for_patterns(
            str(file_path),
            original_filename=original_filename,
            prompt=_VISION_PATTERN_EXTRACT_PROMPT,
            max_tokens=12000,
        )
        if not text:
            debug_entry["status"] = "llm_empty_response"
            print(f"[PatternExtract] sample[{sample_index}] LLM returned empty text: {original_filename}")
            by_sample.append(sample_entry)
            extraction_debug.append(debug_entry)
            continue

        debug_entry["raw_text_len"] = len(text)
        debug_entry["raw_text_preview"] = text[:500]
        debug_entry["raw_text_tail"] = text[-300:] if len(text) > 300 else ""

        parsed = _extract_json_block(text)

        # 1차 파싱 실패 → 잘린 JSON 의 patterns 배열 부분 복구 시도
        if not parsed or not isinstance(parsed, dict) or "patterns" not in parsed:
            recovered = _recover_patterns_from_truncated(text)
            if recovered is not None:
                parsed = {"patterns": recovered}
                debug_entry["status"] = "parsed_via_recovery"
                print(f"[PatternExtract] sample[{sample_index}] recovered {len(recovered)} patterns from truncated JSON")
            else:
                debug_entry["status"] = "json_parse_failed"
                print(f"[PatternExtract] sample[{sample_index}] JSON parse FAILED for {original_filename}, "
                      f"raw_len={len(text)}, head: {text[:200]!r}")
                by_sample.append(sample_entry)
                extraction_debug.append(debug_entry)
                continue

        raw_patterns = parsed.get("patterns") or []
        if not isinstance(raw_patterns, list):
            debug_entry["status"] = "patterns_not_list"
            print(f"[PatternExtract] sample[{sample_index}] 'patterns' is not a list (got {type(raw_patterns).__name__})")
            by_sample.append(sample_entry)
            extraction_debug.append(debug_entry)
            continue

        added = 0
        for n, pat in enumerate(raw_patterns):
            if not isinstance(pat, dict):
                continue
            # 좌표가 없으면 패턴으로 사용 불가 — 스킵
            regions = pat.get("regions") or []
            if not isinstance(regions, list) or len(regions) == 0:
                continue
            pat_copy = dict(pat)
            pat_copy["id"] = f"ext_{sample_index}_{n}"
            pat_copy["source_sample_index"] = sample_index
            pat_copy["source_sample_url"] = url
            extracted_patterns.append(pat_copy)
            sample_entry["patterns"].append(pat_copy)
            added += 1

        debug_entry["patterns_found"] = added
        if added == 0 and len(raw_patterns) > 0:
            debug_entry["status"] = "patterns_invalid_skipped_all"
            print(f"[PatternExtract] sample[{sample_index}] all {len(raw_patterns)} patterns lacked regions; skipped")
        else:
            print(f"[PatternExtract] sample[{sample_index}] ok — extracted {added} patterns from {original_filename}")

        by_sample.append(sample_entry)
        extraction_debug.append(debug_entry)

    # 도큐먼트 업데이트 (덮어쓰기) — extraction_debug 도 함께 저장해 관리자/개발자가 확인 가능
    now = datetime.utcnow()
    await db.ppt_styles.update_one(
        {"_id": oid},
        {"$set": {
            "extracted_patterns": extracted_patterns,
            "pattern_extraction_debug": extraction_debug,
            "pattern_extracted_at": now,
            "updated_at": now,
        }},
    )

    print(f"[PatternExtract] style_id={style_id} 완료 — 총 {len(extracted_patterns)} 패턴 추출")

    return {
        "extracted_patterns": extracted_patterns,
        "by_sample": by_sample,
        "total_patterns": len(extracted_patterns),
        "extraction_debug": extraction_debug,
    }


def _recover_patterns_from_truncated(text: str) -> Optional[list]:
    """잘린/오염된 LLM 응답에서 `"patterns": [...]` 배열의 완성된 객체들만 복구.

    LLM 이 12000 토큰을 다 채워 잘렸을 때, JSON 파서는 전체를 못 읽지만 앞쪽
    완성된 패턴 객체는 살릴 수 있다.
    """
    if not text:
        return None
    # 코드펜스 안쪽만 추출 시도
    fence = re.search(r"```(?:json)?\s*(\{[\s\S]*)", text)
    body = fence.group(1) if fence else text

    # "patterns": [ 시작 위치
    m = re.search(r'"patterns"\s*:\s*\[', body)
    if not m:
        return None
    start = m.end()

    # 배열 안에서 완성된 { ... } 객체만 하나씩 읽어 들임 (중괄호 균형)
    patterns: list = []
    i = start
    n = len(body)
    while i < n:
        # 공백/콤마 스킵
        while i < n and body[i] in " \t\r\n,":
            i += 1
        if i >= n or body[i] == "]":
            break
        if body[i] != "{":
            # 다음 { 까지 점프 — 손상 데이터 방어
            nxt = body.find("{", i)
            if nxt == -1:
                break
            i = nxt
        # 한 객체 균형 잡기
        depth = 0
        in_str = False
        esc = False
        obj_start = i
        end_idx = -1
        while i < n:
            ch = body[i]
            if esc:
                esc = False
                i += 1
                continue
            if ch == "\\":
                esc = True
                i += 1
                continue
            if ch == '"':
                in_str = not in_str
            elif not in_str:
                if ch == "{":
                    depth += 1
                elif ch == "}":
                    depth -= 1
                    if depth == 0:
                        end_idx = i
                        i += 1
                        break
            i += 1
        if end_idx < 0:
            break  # 잘림 — 이 객체는 폐기
        chunk = body[obj_start:end_idx + 1]
        try:
            obj = json.loads(chunk)
            if isinstance(obj, dict):
                patterns.append(obj)
        except Exception:
            # 손상된 객체 1개 → 건너뛰고 계속
            pass
    return patterns if patterns else None


# ============ 빌더 친화 JSON 빌드 ============

async def generate_skill_file(style_id: str) -> dict:
    """빌더 즉시 사용 가능한 정규화 JSON 반환.

    - font_refs → 실제 fonts 컬렉션의 {name, family, url} 메타로 join
    - sample_image_refs → URL/파일명 리스트
    - design_tokens, pattern_library, structurer_prompt, builder_hints 그대로 포함
    """
    db = get_db()
    try:
        oid = ObjectId(style_id)
    except Exception:
        return {}

    doc = await db.ppt_styles.find_one({"_id": oid})
    if not doc:
        return {}

    # 폰트 메타 join
    font_refs: list[str] = doc.get("font_refs", []) or []
    fonts: list[dict] = []
    for fid in font_refs:
        try:
            f = await db.fonts.find_one({"_id": ObjectId(fid)})
        except Exception:
            f = None
        if f:
            fonts.append({
                "_id": str(f["_id"]),
                "name": f.get("name", ""),
                "family": f.get("family", ""),
                "url": f.get("url", ""),
                # 글리프 PNG 미리보기 — 디자이너 멀티모달 첨부 (M6 자동 생성)
                "preview_url": f.get("preview_url", ""),
                "preview_local_path": f.get("preview_local_path", ""),
            })

    sample_refs = doc.get("sample_image_refs", []) or []
    samples = [
        {
            "file_id": s.get("file_id"),
            "url": s.get("url"),
            "original_filename": s.get("original_filename"),
        }
        for s in sample_refs
    ]

    return {
        "_id": str(doc["_id"]),
        "title": doc.get("title", ""),
        "description": doc.get("description", ""),
        "lang": doc.get("lang", "ko"),
        "is_published": bool(doc.get("is_published", False)),
        "design_tokens": doc.get("design_tokens", {}),
        "fonts": fonts,
        "samples": samples,
        "pattern_library": doc.get("pattern_library", []),
        "extracted_patterns": doc.get("extracted_patterns", []),
        "structurer_prompt": doc.get("structurer_prompt", ""),
        "builder_hints": doc.get("builder_hints", {}),
        "vision_analysis": doc.get("vision_analysis", {}),
    }


# ============ 샘플 파일 삭제 헬퍼 ============

def delete_sample_file(style_id: str, filename: str) -> bool:
    """샘플 이미지 파일 삭제 (실패해도 예외 던지지 않음)"""
    try:
        target = get_sample_dir(style_id) / filename
        if target.exists():
            target.unlink()
            return True
    except Exception as e:
        print(f"[PPTStyle] 샘플 파일 삭제 실패 {filename}: {e}")
    return False


def delete_all_samples(style_id: str) -> int:
    """스타일에 속한 모든 샘플 디렉토리 제거. 삭제된 파일 수 반환"""
    import shutil
    sample_dir = get_sample_dir(style_id)
    count = 0
    if sample_dir.exists():
        try:
            for p in sample_dir.iterdir():
                if p.is_file():
                    count += 1
            shutil.rmtree(sample_dir, ignore_errors=True)
        except Exception as e:
            print(f"[PPTStyle] 샘플 디렉토리 삭제 실패 {style_id}: {e}")
    return count


# ============================================================================
# M11 — 결정론적 패턴 매칭 + design_spec 생성 (LLM 호출 없음)
# ============================================================================

# 캔버스 상수 (ppt_builder_service 와 동일)
_CANVAS_W_IN = 10.0
_CANVAS_H_IN = 5.625


# 12개 기본 패턴 ID → outline type 매핑 (extracted_patterns 가 없을 때 fallback 으로 사용)
_DEFAULT_PATTERN_TYPE_MAP: dict = {
    "cover":                    {"outline_types": ["title", "cover"],         "card_count": 0},
    "toc":                      {"outline_types": ["toc"],                    "card_count": 0},
    "chapter":                  {"outline_types": ["section", "chapter"],     "card_count": 0},
    "content_3col":             {"outline_types": ["content"],                "card_count": 3},
    "content_2col_hero":        {"outline_types": ["content"],                "card_count": 2},
    "content_2x2":              {"outline_types": ["content"],                "card_count": 4},
    "big_stat":                 {"outline_types": ["content"],                "card_count": 1},
    "content_3col_icon_block":  {"outline_types": ["content"],                "card_count": 3},
    "content_2_numbered":       {"outline_types": ["content"],                "card_count": 2},
    "content_3col_sidebar":     {"outline_types": ["content"],                "card_count": 3},
    "content_2x2_top_line":     {"outline_types": ["content"],                "card_count": 4},
    "closing":                  {"outline_types": ["closing"],                "card_count": 0},
}


def _builtin_pattern_geometry(pid: str) -> dict | None:
    """12 기본 패턴 ID 에 대해 region 좌표/역할이 포함된 정규화 패턴 dict 를 반환.

    extracted_patterns 가 비어있을 때만 사용. percent 좌표(0-100) 기반.
    """
    meta = _DEFAULT_PATTERN_TYPE_MAP.get(pid)
    if not meta:
        return None

    outline_types = meta["outline_types"]
    card_count = meta["card_count"]

    if pid == "cover":
        regions = [
            {"role": "background"},
            {"role": "title",    "x_pct": 8,  "y_pct": 38, "w_pct": 84, "h_pct": 16, "size_hint": "large"},
            {"role": "subtitle", "x_pct": 8,  "y_pct": 56, "w_pct": 84, "h_pct": 8,  "size_hint": "medium"},
            {"role": "body",     "x_pct": 8,  "y_pct": 66, "w_pct": 84, "h_pct": 6,  "size_hint": "small"},
        ]
    elif pid == "toc":
        regions = [
            {"role": "background"},
            {"role": "title",      "x_pct": 6,  "y_pct": 8,  "w_pct": 88, "h_pct": 14, "size_hint": "large"},
            {"role": "card_title", "x_pct": 6,  "y_pct": 30, "w_pct": 42, "h_pct": 12, "size_hint": "medium"},
            {"role": "card_title", "x_pct": 52, "y_pct": 30, "w_pct": 42, "h_pct": 12, "size_hint": "medium"},
            {"role": "card_title", "x_pct": 6,  "y_pct": 50, "w_pct": 42, "h_pct": 12, "size_hint": "medium"},
            {"role": "card_title", "x_pct": 52, "y_pct": 50, "w_pct": 42, "h_pct": 12, "size_hint": "medium"},
            {"role": "card_title", "x_pct": 6,  "y_pct": 70, "w_pct": 42, "h_pct": 12, "size_hint": "medium"},
            {"role": "card_title", "x_pct": 52, "y_pct": 70, "w_pct": 42, "h_pct": 12, "size_hint": "medium"},
        ]
    elif pid == "chapter":
        regions = [
            {"role": "background"},
            {"role": "decoration", "x_pct": 6, "y_pct": 30, "w_pct": 4,  "h_pct": 40, "size_hint": "small"},
            {"role": "title",      "x_pct": 12, "y_pct": 30, "w_pct": 80, "h_pct": 22, "size_hint": "large"},
            {"role": "subtitle",   "x_pct": 12, "y_pct": 56, "w_pct": 80, "h_pct": 12, "size_hint": "medium"},
        ]
    elif pid == "content_3col":
        regions = [
            {"role": "background"},
            {"role": "title",     "x_pct": 6,  "y_pct": 6,  "w_pct": 88, "h_pct": 12, "size_hint": "large"},
            {"role": "subtitle",  "x_pct": 6,  "y_pct": 19, "w_pct": 88, "h_pct": 6,  "size_hint": "small"},
            # 3 cards
            {"role": "card_title", "x_pct": 6,  "y_pct": 32, "w_pct": 28, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 6,  "y_pct": 44, "w_pct": 28, "h_pct": 40, "size_hint": "small"},
            {"role": "card_title", "x_pct": 36, "y_pct": 32, "w_pct": 28, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 36, "y_pct": 44, "w_pct": 28, "h_pct": 40, "size_hint": "small"},
            {"role": "card_title", "x_pct": 66, "y_pct": 32, "w_pct": 28, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 66, "y_pct": 44, "w_pct": 28, "h_pct": 40, "size_hint": "small"},
        ]
    elif pid == "content_2col_hero":
        regions = [
            {"role": "background"},
            {"role": "title",     "x_pct": 6,  "y_pct": 6,  "w_pct": 88, "h_pct": 12, "size_hint": "large"},
            {"role": "image",     "x_pct": 6,  "y_pct": 22, "w_pct": 42, "h_pct": 70, "size_hint": "large"},
            {"role": "card_title","x_pct": 52, "y_pct": 26, "w_pct": 42, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc", "x_pct": 52, "y_pct": 38, "w_pct": 42, "h_pct": 24, "size_hint": "small"},
            {"role": "card_title","x_pct": 52, "y_pct": 66, "w_pct": 42, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc", "x_pct": 52, "y_pct": 78, "w_pct": 42, "h_pct": 14, "size_hint": "small"},
        ]
    elif pid == "content_2x2":
        regions = [
            {"role": "background"},
            {"role": "title",      "x_pct": 6,  "y_pct": 6,  "w_pct": 88, "h_pct": 12, "size_hint": "large"},
            # 2x2 grid
            {"role": "card_title", "x_pct": 6,  "y_pct": 24, "w_pct": 42, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 6,  "y_pct": 35, "w_pct": 42, "h_pct": 20, "size_hint": "small"},
            {"role": "card_title", "x_pct": 52, "y_pct": 24, "w_pct": 42, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 52, "y_pct": 35, "w_pct": 42, "h_pct": 20, "size_hint": "small"},
            {"role": "card_title", "x_pct": 6,  "y_pct": 60, "w_pct": 42, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 6,  "y_pct": 71, "w_pct": 42, "h_pct": 20, "size_hint": "small"},
            {"role": "card_title", "x_pct": 52, "y_pct": 60, "w_pct": 42, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 52, "y_pct": 71, "w_pct": 42, "h_pct": 20, "size_hint": "small"},
        ]
    elif pid == "big_stat":
        regions = [
            {"role": "background"},
            {"role": "title",    "x_pct": 6,  "y_pct": 10, "w_pct": 88, "h_pct": 12, "size_hint": "large"},
            {"role": "stat",     "x_pct": 6,  "y_pct": 28, "w_pct": 88, "h_pct": 40, "size_hint": "large"},
            {"role": "subtitle", "x_pct": 6,  "y_pct": 72, "w_pct": 88, "h_pct": 10, "size_hint": "medium"},
            {"role": "body",     "x_pct": 6,  "y_pct": 84, "w_pct": 88, "h_pct": 10, "size_hint": "small"},
        ]
    elif pid == "content_3col_icon_block":
        regions = [
            {"role": "background"},
            {"role": "title",      "x_pct": 6,  "y_pct": 6,  "w_pct": 88, "h_pct": 12, "size_hint": "large"},
            # 3 cards with icons
            {"role": "icon",       "x_pct": 14, "y_pct": 28, "w_pct": 8,  "h_pct": 12, "size_hint": "medium"},
            {"role": "card_title", "x_pct": 6,  "y_pct": 44, "w_pct": 28, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 6,  "y_pct": 55, "w_pct": 28, "h_pct": 35, "size_hint": "small"},
            {"role": "icon",       "x_pct": 44, "y_pct": 28, "w_pct": 8,  "h_pct": 12, "size_hint": "medium"},
            {"role": "card_title", "x_pct": 36, "y_pct": 44, "w_pct": 28, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 36, "y_pct": 55, "w_pct": 28, "h_pct": 35, "size_hint": "small"},
            {"role": "icon",       "x_pct": 74, "y_pct": 28, "w_pct": 8,  "h_pct": 12, "size_hint": "medium"},
            {"role": "card_title", "x_pct": 66, "y_pct": 44, "w_pct": 28, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 66, "y_pct": 55, "w_pct": 28, "h_pct": 35, "size_hint": "small"},
        ]
    elif pid == "content_2_numbered":
        regions = [
            {"role": "background"},
            {"role": "title",      "x_pct": 6,  "y_pct": 6,  "w_pct": 88, "h_pct": 12, "size_hint": "large"},
            {"role": "decoration", "x_pct": 6,  "y_pct": 28, "w_pct": 8,  "h_pct": 16, "size_hint": "medium"},
            {"role": "card_title", "x_pct": 16, "y_pct": 28, "w_pct": 78, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 16, "y_pct": 40, "w_pct": 78, "h_pct": 18, "size_hint": "small"},
            {"role": "decoration", "x_pct": 6,  "y_pct": 62, "w_pct": 8,  "h_pct": 16, "size_hint": "medium"},
            {"role": "card_title", "x_pct": 16, "y_pct": 62, "w_pct": 78, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 16, "y_pct": 74, "w_pct": 78, "h_pct": 18, "size_hint": "small"},
        ]
    elif pid == "content_3col_sidebar":
        regions = [
            {"role": "background"},
            {"role": "title",     "x_pct": 32, "y_pct": 6,  "w_pct": 62, "h_pct": 12, "size_hint": "large"},
            {"role": "subtitle",  "x_pct": 6,  "y_pct": 6,  "w_pct": 24, "h_pct": 88, "size_hint": "medium"},
            # 3 stacked rows
            {"role": "card_title", "x_pct": 32, "y_pct": 22, "w_pct": 62, "h_pct": 8,  "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 32, "y_pct": 31, "w_pct": 62, "h_pct": 14, "size_hint": "small"},
            {"role": "card_title", "x_pct": 32, "y_pct": 48, "w_pct": 62, "h_pct": 8,  "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 32, "y_pct": 57, "w_pct": 62, "h_pct": 14, "size_hint": "small"},
            {"role": "card_title", "x_pct": 32, "y_pct": 74, "w_pct": 62, "h_pct": 8,  "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 32, "y_pct": 83, "w_pct": 62, "h_pct": 14, "size_hint": "small"},
        ]
    elif pid == "content_2x2_top_line":
        regions = [
            {"role": "background"},
            {"role": "title",      "x_pct": 6,  "y_pct": 6,  "w_pct": 88, "h_pct": 12, "size_hint": "large"},
            {"role": "decoration", "x_pct": 6,  "y_pct": 22, "w_pct": 42, "h_pct": 1,  "size_hint": "small"},
            {"role": "card_title", "x_pct": 6,  "y_pct": 25, "w_pct": 42, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 6,  "y_pct": 36, "w_pct": 42, "h_pct": 18, "size_hint": "small"},
            {"role": "decoration", "x_pct": 52, "y_pct": 22, "w_pct": 42, "h_pct": 1,  "size_hint": "small"},
            {"role": "card_title", "x_pct": 52, "y_pct": 25, "w_pct": 42, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 52, "y_pct": 36, "w_pct": 42, "h_pct": 18, "size_hint": "small"},
            {"role": "decoration", "x_pct": 6,  "y_pct": 60, "w_pct": 42, "h_pct": 1,  "size_hint": "small"},
            {"role": "card_title", "x_pct": 6,  "y_pct": 63, "w_pct": 42, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 6,  "y_pct": 74, "w_pct": 42, "h_pct": 18, "size_hint": "small"},
            {"role": "decoration", "x_pct": 52, "y_pct": 60, "w_pct": 42, "h_pct": 1,  "size_hint": "small"},
            {"role": "card_title", "x_pct": 52, "y_pct": 63, "w_pct": 42, "h_pct": 10, "size_hint": "medium"},
            {"role": "card_desc",  "x_pct": 52, "y_pct": 74, "w_pct": 42, "h_pct": 18, "size_hint": "small"},
        ]
    elif pid == "closing":
        regions = [
            {"role": "background"},
            {"role": "title",    "x_pct": 8,  "y_pct": 30, "w_pct": 84, "h_pct": 18, "size_hint": "large"},
            {"role": "subtitle", "x_pct": 8,  "y_pct": 50, "w_pct": 84, "h_pct": 12, "size_hint": "medium"},
            {"role": "body",     "x_pct": 8,  "y_pct": 66, "w_pct": 84, "h_pct": 10, "size_hint": "small"},
        ]
    else:
        return None

    return {
        "id": pid,
        "label": pid,
        "regions": regions,
        "suitable_for_outline_types": outline_types,
        "card_count": card_count,
        "visual_keywords": [],
        "enabled": True,
        "source": "builtin",
    }


def _normalize_outline_type(outline_type: str) -> str:
    """outline.type 정규화 (소문자, 공백제거)"""
    if not isinstance(outline_type, str):
        return "content"
    t = outline_type.strip().lower()
    return t or "content"


def _outline_type_matches_pattern(outline_type: str, pattern: dict) -> bool:
    """outline.type 이 pattern.suitable_for_outline_types 와 매칭되는지 확인.

    매칭 규칙:
      title   → "title" 또는 "cover"
      toc     → "toc"
      section → "section" 또는 "chapter"
      content → "content"
      closing → "closing"
    """
    t = _normalize_outline_type(outline_type)
    suitable = pattern.get("suitable_for_outline_types") or []
    if not isinstance(suitable, list):
        return False
    suitable_lower = [str(s).strip().lower() for s in suitable]

    if t == "title":
        return any(s in ("title", "cover") for s in suitable_lower)
    if t == "toc":
        return "toc" in suitable_lower
    if t == "section":
        return any(s in ("section", "chapter") for s in suitable_lower)
    if t == "content":
        return "content" in suitable_lower
    if t == "closing":
        return "closing" in suitable_lower
    # fallback: 직접 비교
    return t in suitable_lower


def _count_card_regions(pattern: dict) -> int:
    """패턴이 가진 카드 슬롯 수 추정 — card_title 가 가장 신뢰성 있는 척도."""
    regions = pattern.get("regions") or []
    if not isinstance(regions, list):
        return 0
    card_titles = sum(
        1 for r in regions
        if isinstance(r, dict) and (r.get("role") or "").lower() == "card_title"
    )
    if card_titles > 0:
        return card_titles
    # fallback: card_desc 수
    return sum(
        1 for r in regions
        if isinstance(r, dict) and (r.get("role") or "").lower() == "card_desc"
    )


def match_pattern_for_outline_slide(
    outline_slide: dict,
    skill_file: dict,
    used_patterns: dict | None = None,
) -> dict | None:
    """outline 슬라이드 1개에 가장 적합한 패턴을 결정론적으로 선택.

    선택 우선순위:
      1. extracted_patterns 중 enabled != False 이고 outline.type 호환되는 것
      2. extracted_patterns 비어있으면 pattern_library(12 기본) 중 enabled 인 것 + builtin 좌표
      3. content type 의 경우 card_count == len(items) 우선 (±1 차이 허용)
      4. 같은 type 의 슬라이드가 여러 개면 라운드로빈 (used_patterns 사용)

    Args:
      outline_slide: {"type": "title|toc|section|content|closing", ...}
      skill_file: ppt_style_service.generate_skill_file() 결과
      used_patterns: {type: [pattern_id, ...]} 라운드로빈 추적용 (mutable)

    Returns:
      선택된 패턴 dict 또는 None (적합 패턴 없음)
    """
    if not isinstance(outline_slide, dict) or not isinstance(skill_file, dict):
        return None

    out_type = _normalize_outline_type(outline_slide.get("type", "content"))

    # 후보 패턴 수집
    candidates: list[dict] = []
    extracted = skill_file.get("extracted_patterns") or []
    if isinstance(extracted, list):
        for p in extracted:
            if not isinstance(p, dict):
                continue
            if p.get("enabled") is False:
                continue
            if not p.get("regions"):
                continue
            if _outline_type_matches_pattern(out_type, p):
                candidates.append(p)

    # extracted 비어있으면 12 기본 패턴 사용
    if not candidates:
        pat_lib = skill_file.get("pattern_library") or []
        if not isinstance(pat_lib, list):
            pat_lib = []
        for entry in pat_lib:
            if not isinstance(entry, dict):
                continue
            if entry.get("enabled") is False:
                continue
            pid = entry.get("id") or ""
            built = _builtin_pattern_geometry(pid)
            if not built:
                continue
            if _outline_type_matches_pattern(out_type, built):
                candidates.append(built)
        # pattern_library 가 비어있거나 전부 비활성이면 12개 전체 fallback
        if not candidates:
            for pid in _DEFAULT_PATTERN_TYPE_MAP.keys():
                built = _builtin_pattern_geometry(pid)
                if built and _outline_type_matches_pattern(out_type, built):
                    candidates.append(built)

    if not candidates:
        return None

    # content 타입: card_count 가까운 것 우선
    if out_type == "content":
        items = outline_slide.get("items") or []
        n_items = len(items) if isinstance(items, list) else 0
        if n_items > 0:
            def _card_distance(p: dict) -> int:
                cc = _count_card_regions(p)
                if cc <= 0:
                    cc = int(p.get("card_count") or 0)
                return abs(cc - n_items) if cc > 0 else 99

            # 가장 가까운 거리 그룹만 선별
            distances = [(p, _card_distance(p)) for p in candidates]
            min_dist = min(d for _, d in distances)
            # ±1 허용 (최소가 0이면 0만, 1이면 0/1)
            close = [p for p, d in distances if d <= max(1, min_dist)]
            if close:
                candidates = close

    # 라운드로빈: 같은 type 에서 사용 횟수가 가장 적은 패턴 선택
    used_by_type = (used_patterns or {}).get(out_type) or []
    # 패턴 id 별 사용 횟수
    use_counts: list[tuple[int, int, dict]] = []
    for idx, p in enumerate(candidates):
        pid = p.get("id") or f"_anon_{idx}"
        used = sum(1 for x in used_by_type if x == pid)
        use_counts.append((used, idx, p))
    # 사용 횟수 오름차순, 동률이면 원래 순서
    use_counts.sort(key=lambda t: (t[0], t[1]))
    return use_counts[0][2]


def _pct_to_inch_x(pct) -> float:
    try:
        v = float(pct)
    except (TypeError, ValueError):
        v = 0.0
    return max(0.0, min(v, 100.0)) / 100.0 * _CANVAS_W_IN


def _pct_to_inch_y(pct) -> float:
    try:
        v = float(pct)
    except (TypeError, ValueError):
        v = 0.0
    return max(0.0, min(v, 100.0)) / 100.0 * _CANVAS_H_IN


def _pct_to_inch_w(pct) -> float:
    try:
        v = float(pct)
    except (TypeError, ValueError):
        v = 0.0
    return max(0.01, min(v, 100.0)) / 100.0 * _CANVAS_W_IN


def _pct_to_inch_h(pct) -> float:
    try:
        v = float(pct)
    except (TypeError, ValueError):
        v = 0.0
    return max(0.01, min(v, 100.0)) / 100.0 * _CANVAS_H_IN


def _get_color(skill_file: dict, key: str, fallback: str) -> str:
    """skill_file.design_tokens.colors[key] 반환 (없으면 fallback)"""
    if not isinstance(skill_file, dict):
        return fallback
    dt = skill_file.get("design_tokens") or {}
    colors = dt.get("colors") or {}
    val = colors.get(key)
    if isinstance(val, str) and val.strip():
        return val if val.startswith("#") else "#" + val.strip()
    return fallback


def _get_font_family(skill_file: dict, role: str) -> str:
    """skill_file.design_tokens.fonts.{role}_font_family 또는 fonts 목록의 family"""
    if not isinstance(skill_file, dict):
        return ""
    dt = skill_file.get("design_tokens") or {}
    fonts_meta = dt.get("fonts") or {}
    # 직접 family 키
    key = "title_font_family" if role == "title" else "body_font_family"
    val = fonts_meta.get(key)
    if isinstance(val, str) and val.strip():
        return val.strip()
    # fonts 목록 fallback
    fonts_list = skill_file.get("fonts") or []
    target_id_key = "title_font_id" if role == "title" else "body_font_id"
    target_id = fonts_meta.get(target_id_key)
    if target_id and isinstance(fonts_list, list):
        for f in fonts_list:
            if isinstance(f, dict) and str(f.get("_id")) == str(target_id):
                fam = f.get("family") or f.get("name")
                if isinstance(fam, str) and fam.strip():
                    return fam.strip()
    if isinstance(fonts_list, list) and fonts_list:
        first = fonts_list[0]
        if isinstance(first, dict):
            fam = first.get("family") or first.get("name")
            if isinstance(fam, str) and fam.strip():
                return fam.strip()
    return ""


def _get_font_size(skill_file: dict, role: str) -> int:
    """skill_file.design_tokens.fonts.sizes 에서 role 별 사이즈"""
    dt = (skill_file or {}).get("design_tokens") or {}
    sizes = (dt.get("fonts") or {}).get("sizes") or {}
    # role → size key
    role_to_size = {
        "title":      "h1",
        "subtitle":   "h2",
        "body":       "body",
        "card_title": "h2",
        "card_desc":  "body",
        "stat":       "stat",
    }
    key = role_to_size.get(role, "body")
    default = {"h1": 26, "h2": 22, "body": 11, "label": 10, "stat": 88}.get(key, 14)
    val = sizes.get(key, default)
    try:
        return int(val)
    except (TypeError, ValueError):
        return default


def _is_dark_color(hex_str: str) -> bool:
    """hex 컬러가 어두운 색인지 (배경 위 텍스트 색 결정용)"""
    if not isinstance(hex_str, str):
        return False
    s = hex_str.strip().lstrip("#")
    if len(s) != 6:
        return False
    try:
        r = int(s[0:2], 16)
        g = int(s[2:4], 16)
        b = int(s[4:6], 16)
    except Exception:
        return False
    # YIQ 휘도
    yiq = (r * 299 + g * 587 + b * 114) / 1000
    return yiq < 128


def _resolve_outline_field(
    outline_slide: dict,
    role: str,
    item_index: int = 0,
) -> str | None:
    """outline 필드를 role 에 따라 추출.

    role → 우선순위 필드 매핑.
    """
    if not isinstance(outline_slide, dict):
        return None
    out_type = _normalize_outline_type(outline_slide.get("type"))
    items = outline_slide.get("items") or []
    if not isinstance(items, list):
        items = []

    def _txt(*keys: str) -> str | None:
        for k in keys:
            v = outline_slide.get(k)
            if isinstance(v, str) and v.strip():
                return v.strip()
        return None

    def _item_txt(idx: int, *keys: str) -> str | None:
        if not (0 <= idx < len(items)):
            return None
        it = items[idx]
        if not isinstance(it, dict):
            return None
        for k in keys:
            v = it.get(k)
            if isinstance(v, str) and v.strip():
                return v.strip()
        return None

    if role == "title":
        return _txt("title", "section_title")
    if role == "subtitle":
        return _txt("subtitle", "section_subtitle", "summary")
    if role == "body":
        v = _txt("message", "closing_message", "description")
        if v:
            return v
        return _item_txt(0, "detail", "description")
    if role == "card_title":
        # toc: section_no + title, content: heading
        if out_type == "toc":
            it = items[item_index] if 0 <= item_index < len(items) else None
            if isinstance(it, dict):
                section_no = it.get("section_no")
                title = it.get("title")
                parts: list[str] = []
                if isinstance(section_no, str) and section_no.strip():
                    parts.append(section_no.strip())
                if isinstance(title, str) and title.strip():
                    parts.append(title.strip())
                if parts:
                    return "  ".join(parts)
            return None
        return _item_txt(item_index, "heading", "title")
    if role == "card_desc":
        return _item_txt(item_index, "detail", "description", "body")
    if role == "stat":
        v = _txt("stat", "metric", "value")
        if v:
            return v
        return _item_txt(0, "stat", "metric", "value")
    if role == "icon":
        return _item_txt(item_index, "icon")
    return None


# Valid icon keys (ppt_builder_service 와 동기화)
_VALID_ICON_KEYS = {
    "users", "atom", "ship", "bomb", "shield", "fire", "chart", "star",
    "lightning", "target", "brain", "gear", "lock", "leaf", "globe", "rocket",
    "eye", "check", "warning", "idea", "clock", "money", "building", "network",
    "book", "cloud", "code", "database", "mobile", "palette",
}


def build_design_spec_from_pattern(
    outline_slide: dict,
    pattern: dict | None,
    skill_file: dict,
    slide_index: int,
    sample_image_pool: list | None = None,
    sample_pointer: list | None = None,
) -> dict:
    """outline + pattern 으로부터 빌더가 바로 쓸 수 있는 design_spec 생성.

    캔버스: 10x5.625 인치. 패턴의 (x_pct, y_pct, w_pct, h_pct) 를 인치로 변환.

    Returns:
      design_spec dict (빌더 호환)
    """
    skill_file = skill_file if isinstance(skill_file, dict) else {}
    outline_slide = outline_slide if isinstance(outline_slide, dict) else {}

    # 토큰
    c_primary = _get_color(skill_file, "primary", "#1C60EF")
    c_light = _get_color(skill_file, "light", "#DBE8FE")
    c_white = _get_color(skill_file, "white", "#FFFFFF")
    c_ink = _get_color(skill_file, "ink", "#0B1E3F")
    c_grey = _get_color(skill_file, "grey", "#6B7B9A")
    c_line = _get_color(skill_file, "line", "#D5E2FB")

    title_font = _get_font_family(skill_file, "title")
    body_font = _get_font_family(skill_file, "body")

    out_type = _normalize_outline_type(outline_slide.get("type"))

    # --- pattern 이 None 이면 fallback (제목만 가운데) ---
    if not pattern or not isinstance(pattern, dict):
        title_text = _resolve_outline_field(outline_slide, "title") or "발표 자료"
        return {
            "slide_index": slide_index,
            "layout_hint": "fallback_blank",
            "background": {
                "type": "gradient",
                "from_color": c_light,
                "to_color": c_white,
                "style": "radial",
                "origin": [0.3, 0.5],
            },
            "regions": [
                {
                    "type": "text",
                    "x": 0.5, "y": 2.5, "w": 9.0, "h": 1.0,
                    "text": title_text,
                    "font_family": title_font,
                    "font_size": _get_font_size(skill_file, "title"),
                    "bold": True,
                    "color": c_ink,
                    "align": "center",
                    "valign": "middle",
                }
            ],
        }

    pattern_regions = pattern.get("regions") or []
    if not isinstance(pattern_regions, list):
        pattern_regions = []

    # background 영역 분리 & 배경 spec 결정
    bg_region: dict | None = None
    body_regions: list[dict] = []
    for r in pattern_regions:
        if not isinstance(r, dict):
            continue
        role = (r.get("role") or "").lower()
        if role == "background" and bg_region is None:
            bg_region = r
        else:
            body_regions.append(r)

    # 배경 결정
    background: dict
    if bg_region:
        approx = bg_region.get("approx_color")
        if isinstance(approx, str) and re.match(r"^#?[0-9A-Fa-f]{6}$", approx.strip()):
            color_hex = approx if approx.startswith("#") else "#" + approx
            background = {"type": "solid", "from_color": color_hex.upper()}
        else:
            background = {
                "type": "gradient",
                "from_color": c_light,
                "to_color": c_white,
                "style": "radial",
                "origin": [0.3, 0.5],
            }
    else:
        background = {
            "type": "gradient",
            "from_color": c_light,
            "to_color": c_white,
            "style": "radial",
            "origin": [0.3, 0.5],
        }

    bg_is_dark = False
    if background.get("type") == "solid":
        bg_is_dark = _is_dark_color(background.get("from_color", ""))

    # outline.items 카운트
    items = outline_slide.get("items") or []
    if not isinstance(items, list):
        items = []
    n_items = len(items)

    # 카드 슬롯 수 카운트 (mismatch 처리용)
    card_title_count = sum(
        1 for r in body_regions if (r.get("role") or "").lower() == "card_title"
    )
    # K (pattern card slots) vs N (outline items) 처리
    max_cards = card_title_count
    if max_cards > 0 and n_items > 0:
        used_card_count = min(max_cards, n_items)
    else:
        used_card_count = max(max_cards, n_items)

    # role 별 인덱스 카운터
    role_counters: dict[str, int] = {}

    result_regions: list[dict] = []

    sample_pool = sample_image_pool or []
    sample_ptr = sample_pointer if isinstance(sample_pointer, list) else None

    for r in body_regions:
        role = (r.get("role") or "").lower()
        idx_for_role = role_counters.get(role, 0)

        x_in = _pct_to_inch_x(r.get("x_pct", 0))
        y_in = _pct_to_inch_y(r.get("y_pct", 0))
        w_in = _pct_to_inch_w(r.get("w_pct", 0))
        h_in = _pct_to_inch_h(r.get("h_pct", 0))

        # 카드 mismatch 처리
        if role in ("card_title", "card_desc"):
            if idx_for_role >= n_items:
                # K > N: 남는 슬롯 생략
                role_counters[role] = idx_for_role + 1
                continue
            if idx_for_role >= max_cards and max_cards > 0:
                # K < N 인데 슬롯이 max_cards 까지인 경우 → 이미 위 분기로 처리됨
                role_counters[role] = idx_for_role + 1
                continue

        # ─── role → outline 필드 매핑 ───
        if role == "title":
            text = _resolve_outline_field(outline_slide, "title")
            if not text:
                role_counters[role] = idx_for_role + 1
                continue
            color = c_ink
            if bg_is_dark:
                color = c_white
            result_regions.append({
                "type": "text",
                "x": x_in, "y": y_in, "w": w_in, "h": h_in,
                "text": text,
                "font_family": title_font,
                "font_size": _get_font_size(skill_file, "title"),
                "bold": True,
                "color": color,
                "align": "left",
                "valign": "middle",
            })

        elif role == "subtitle":
            text = _resolve_outline_field(outline_slide, "subtitle")
            if not text:
                role_counters[role] = idx_for_role + 1
                continue
            color = c_grey if not bg_is_dark else c_light
            result_regions.append({
                "type": "text",
                "x": x_in, "y": y_in, "w": w_in, "h": h_in,
                "text": text,
                "font_family": body_font,
                "font_size": _get_font_size(skill_file, "subtitle"),
                "bold": False,
                "color": color,
                "align": "left",
                "valign": "top",
            })

        elif role == "body":
            text = _resolve_outline_field(outline_slide, "body")
            if not text:
                role_counters[role] = idx_for_role + 1
                continue
            color = c_grey if not bg_is_dark else c_light
            result_regions.append({
                "type": "text",
                "x": x_in, "y": y_in, "w": w_in, "h": h_in,
                "text": text,
                "font_family": body_font,
                "font_size": _get_font_size(skill_file, "body"),
                "bold": False,
                "color": color,
                "align": "left",
                "valign": "top",
            })

        elif role == "card_title":
            text = _resolve_outline_field(outline_slide, "card_title", idx_for_role)
            if not text:
                role_counters[role] = idx_for_role + 1
                continue
            color = c_ink if not bg_is_dark else c_white
            result_regions.append({
                "type": "text",
                "x": x_in, "y": y_in, "w": w_in, "h": h_in,
                "text": text,
                "font_family": title_font,
                "font_size": _get_font_size(skill_file, "card_title"),
                "bold": True,
                "color": color,
                "align": "left",
                "valign": "middle",
            })

        elif role == "card_desc":
            text = _resolve_outline_field(outline_slide, "card_desc", idx_for_role)
            if not text:
                role_counters[role] = idx_for_role + 1
                continue
            color = c_grey if not bg_is_dark else c_light
            result_regions.append({
                "type": "text",
                "x": x_in, "y": y_in, "w": w_in, "h": h_in,
                "text": text,
                "font_family": body_font,
                "font_size": _get_font_size(skill_file, "card_desc"),
                "bold": False,
                "color": color,
                "align": "left",
                "valign": "top",
            })

        elif role == "stat":
            text = _resolve_outline_field(outline_slide, "stat")
            if not text:
                role_counters[role] = idx_for_role + 1
                continue
            result_regions.append({
                "type": "text",
                "x": x_in, "y": y_in, "w": w_in, "h": h_in,
                "text": text,
                "font_family": title_font,
                "font_size": _get_font_size(skill_file, "stat"),
                "bold": True,
                "color": c_primary,
                "align": "center",
                "valign": "middle",
            })

        elif role == "icon":
            icon_text = _resolve_outline_field(outline_slide, "icon", idx_for_role)
            icon_key = (icon_text or "").strip().lower()
            if not icon_key or icon_key not in _VALID_ICON_KEYS:
                role_counters[role] = idx_for_role + 1
                continue
            result_regions.append({
                "type": "icon",
                "x": x_in, "y": y_in, "w": w_in, "h": h_in,
                "icon_key": icon_key,
                "color": c_primary,
            })

        elif role == "image":
            if sample_pool and sample_ptr is not None:
                ptr = sample_ptr[0] if sample_ptr else 0
                if 0 <= ptr < len(sample_pool):
                    image_url = sample_pool[ptr]
                else:
                    image_url = sample_pool[ptr % len(sample_pool)] if sample_pool else None
                if sample_ptr:
                    sample_ptr[0] = (sample_ptr[0] + 1) % max(1, len(sample_pool))
                if image_url:
                    result_regions.append({
                        "type": "image",
                        "x": x_in, "y": y_in, "w": w_in, "h": h_in,
                        "image_url": image_url,
                    })
                else:
                    # 이미지 풀이 비었으면 placeholder shape
                    result_regions.append({
                        "type": "shape",
                        "shape": "rectangle",
                        "x": x_in, "y": y_in, "w": w_in, "h": h_in,
                        "fill": c_light,
                        "stroke": "none",
                    })
            else:
                result_regions.append({
                    "type": "shape",
                    "shape": "rectangle",
                    "x": x_in, "y": y_in, "w": w_in, "h": h_in,
                    "fill": c_light,
                    "stroke": "none",
                })

        elif role == "decoration":
            # 도형/색만 적용 (텍스트 없음). 기본은 primary 색의 얇은 직사각형/포인트
            result_regions.append({
                "type": "shape",
                "shape": "rectangle",
                "x": x_in, "y": y_in, "w": w_in, "h": h_in,
                "fill": c_primary,
                "stroke": "none",
            })

        else:
            # 알 수 없는 role 은 무시
            pass

        role_counters[role] = idx_for_role + 1

    return {
        "slide_index": slide_index,
        "layout_hint": pattern.get("id") or pattern.get("label") or out_type,
        "background": background,
        "regions": result_regions,
    }


def generate_design_specs_from_outline(
    outline: dict,
    skill_file: dict,
) -> dict:
    """outline 의 모든 슬라이드에 대해 패턴 매칭 + design_spec 생성.

    Args:
      outline: {"meta":..., "slides":[raw_slides], "sources":[...]}
      skill_file: ppt_style_service.generate_skill_file() 결과

    Returns:
      {
        "design_specs": [...],
        "total_slides": N,
        "pattern_usage": {<type>: [<pattern_id>, ...]},
        "fallback_count": int,
      }
    """
    outline = outline if isinstance(outline, dict) else {}
    skill_file = skill_file if isinstance(skill_file, dict) else {}

    slides = outline.get("slides") or []
    if not isinstance(slides, list):
        slides = []

    # 샘플 이미지 풀 (skill_file.samples 우선, 없으면 sample_image_refs)
    sample_image_pool: list[str] = []
    samples = skill_file.get("samples") or skill_file.get("sample_image_refs") or []
    if isinstance(samples, list):
        for s in samples:
            if isinstance(s, dict):
                url = s.get("url")
                if isinstance(url, str) and url.strip():
                    sample_image_pool.append(url.strip())
            elif isinstance(s, str) and s.strip():
                sample_image_pool.append(s.strip())

    sample_pointer = [0]

    used_patterns: dict[str, list[str]] = {}
    pattern_usage: dict[str, list[str]] = {}
    design_specs: list[dict] = []
    fallback_count = 0

    for i, sl in enumerate(slides, start=1):
        if not isinstance(sl, dict):
            sl = {}
        out_type = _normalize_outline_type(sl.get("type"))

        pattern = match_pattern_for_outline_slide(sl, skill_file, used_patterns)

        spec = build_design_spec_from_pattern(
            sl,
            pattern,
            skill_file,
            slide_index=i,
            sample_image_pool=sample_image_pool,
            sample_pointer=sample_pointer,
        )

        if pattern is None:
            fallback_count += 1
            used_id = "_fallback_"
        else:
            used_id = str(pattern.get("id") or pattern.get("label") or f"anon_{i}")

        used_patterns.setdefault(out_type, []).append(used_id)
        pattern_usage.setdefault(out_type, []).append(used_id)

        design_specs.append(spec)

    return {
        "design_specs": design_specs,
        "total_slides": len(design_specs),
        "pattern_usage": pattern_usage,
        "fallback_count": fallback_count,
    }
