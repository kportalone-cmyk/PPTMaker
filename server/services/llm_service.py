"""
Claude Opus LLM 서비스 - 리치 스키마 기반 슬라이드 콘텐츠 생성

리소스 내용을 분석하여 구조화된 프레젠테이션 스키마를 생성하고,
관리자가 설정한 템플릿 오브젝트(제목/거버넌스/부제목/내용 등)에 매핑합니다.
"""

import json
import re
import random
import base64
import os
import io
import httpx
from PIL import Image
from config import settings
from routers.prompt import get_prompt_content, get_prompt_model


# ============ 슬라이드 타입 ↔ content_type 매핑 ============
SCHEMA_TYPE_TO_CONTENT_TYPE = {
    "title": "title_slide",
    "toc": "toc",
    "section": "section_divider",
    "content": "body",
    "closing": "closing",
}

CONTENT_TYPE_TO_SCHEMA_TYPE = {v: k for k, v in SCHEMA_TYPE_TO_CONTENT_TYPE.items()}


async def generate_slide_content(
    resources_text: str,
    instructions: str,
    slides_meta: list[dict],
    lang: str = "",
    slide_count: str = "auto",
) -> dict:
    """리소스 텍스트를 분석하여 리치 스키마 기반 프레젠테이션을 설계하고 콘텐츠 생성

    Returns:
        {
            "slides": [{"template_index": int, "contents": {placeholder: text}}],
            "meta": {"title", "subtitle", "author", "date"},
            "sources": [{"ref", "title"}]
        }
    """
    num_templates = len(slides_meta)

    if not settings.ANTHROPIC_API_KEY or settings.ANTHROPIC_API_KEY == "your_anthropic_api_key_here":
        return {
            "slides": _fallback_content(resources_text, slides_meta),
            "meta": {},
            "sources": [],
        }

    # DB에서 프롬프트 로드 및 변수 바인딩
    system_prompt, user_prompt = await _build_generation_prompts(
        resources_text, instructions, slides_meta, lang, slide_count=slide_count,
    )

    # 프롬프트별 모델 설정 조회 (DB 우선, 없으면 config 기본값)
    prompt_model = await get_prompt_model("slide_generation_system")
    effective_model = prompt_model or settings.ANTHROPIC_OUTLINE_MODEL

    try:
        result = await _call_claude_api(system_prompt, user_prompt, model=effective_model)
        print(f"[LLM] API 응답 길이: {len(result)} chars (model: {effective_model})")
        parsed = _parse_rich_schema(result, slides_meta)
        if parsed:
            print(f"[LLM] 파싱 성공 - slides: {len(parsed.get('slides', []))}개")
            return parsed
        print(f"[LLM] 파싱 실패 → 폴백 콘텐츠 사용")
        fallback = _fallback_content(resources_text, slides_meta)
        return {"slides": fallback, "meta": {}, "sources": []}
    except Exception as e:
        import traceback
        print(f"[LLM] Claude API 호출 실패: {e}")
        traceback.print_exc()
        fallback = _fallback_content(resources_text, slides_meta)
        return {"slides": fallback, "meta": {}, "sources": []}


async def _call_claude_api(system_prompt: str, user_prompt: str, model: str = "") -> str:
    """Claude API 호출 (httpx 비동기)"""
    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": settings.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    payload = {
        "model": model or settings.ANTHROPIC_MODEL,
        "system": system_prompt,
        "messages": [
            {"role": "user", "content": user_prompt}
        ],
    }
    # 모델별 max_tokens 적용 (Anthropic API 필수 파라미터)
    effective_model = model or settings.ANTHROPIC_MODEL
    if effective_model == settings.ANTHROPIC_OUTLINE_MODEL and settings.ANTHROPIC_OUTLINE_MAX_TOKENS > 0:
        payload["max_tokens"] = settings.ANTHROPIC_OUTLINE_MAX_TOKENS
    elif settings.ANTHROPIC_MAX_TOKENS > 0:
        payload["max_tokens"] = settings.ANTHROPIC_MAX_TOKENS
    else:
        payload["max_tokens"] = 4096  # 기본값

    async with httpx.AsyncClient(timeout=120.0) as client:
        response = await client.post(url, json=payload, headers=headers)
        response.raise_for_status()
        data = response.json()

        content_blocks = data.get("content", [])
        text_parts = []
        for block in content_blocks:
            if block.get("type") == "text":
                text_parts.append(block["text"])

        return "\n".join(text_parts)


async def analyze_image_content(file_path: str, original_filename: str = "") -> str:
    """Claude Vision API로 이미지 내용 분석

    지원 형식: jpg, jpeg, png, gif, webp
    비지원 형식(svg, bmp)은 빈 문자열 반환
    """
    if not settings.ANTHROPIC_API_KEY or settings.ANTHROPIC_API_KEY == "your_anthropic_api_key_here":
        return ""

    ext = os.path.splitext(file_path)[1].lower()
    SUPPORTED_EXT = {".jpg", ".jpeg", ".png", ".gif", ".webp"}
    if ext not in SUPPORTED_EXT:
        return ""

    MEDIA_TYPE_MAP = {
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".gif": "image/gif",
        ".webp": "image/webp",
    }

    try:
        # 이미지 읽기 + 리사이즈 (1568px 이하)
        img = Image.open(file_path)
        w, h = img.size
        MAX_DIM = 1568
        if max(w, h) > MAX_DIM:
            ratio = MAX_DIM / max(w, h)
            img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)

        # RGBA → RGB 변환 (JPEG 저장 시 필요)
        if img.mode in ("RGBA", "LA", "P"):
            img = img.convert("RGB")

        # 바이트 변환 + base64 인코딩
        buf = io.BytesIO()
        out_fmt = "JPEG" if ext in (".jpg", ".jpeg") else "PNG" if ext == ".png" else "WEBP" if ext == ".webp" else "PNG"
        img.save(buf, format=out_fmt)
        image_data = base64.standard_b64encode(buf.getvalue()).decode("utf-8")
        media_type = MEDIA_TYPE_MAP.get(ext, "image/png")

        prompt = (
            f"이 이미지(파일명: {original_filename})의 내용을 자세히 분석하여 설명해주세요.\n"
            "다음 항목을 포함하세요:\n"
            "1. 이미지에 보이는 주요 내용과 요소\n"
            "2. 텍스트가 있다면 텍스트 내용 전문\n"
            "3. 차트/그래프/도표가 있다면 데이터 해석\n"
            "4. 이미지의 전체적인 주제와 맥락\n\n"
            "프레젠테이션 자료 작성에 활용할 수 있도록 구체적이고 정확하게 설명하세요."
        )

        headers = {
            "x-api-key": settings.ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        }
        payload = {
            "model": settings.ANTHROPIC_OUTLINE_MODEL,
            "max_tokens": 1024,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": image_data,
                            },
                        },
                        {"type": "text", "text": prompt},
                    ],
                }
            ],
        }

        async with httpx.AsyncClient(timeout=60.0) as client:
            response = await client.post(
                "https://api.anthropic.com/v1/messages",
                json=payload,
                headers=headers,
            )
            response.raise_for_status()
            data = response.json()

        content_blocks = data.get("content", [])
        text_parts = [block["text"] for block in content_blocks if block.get("type") == "text"]
        result = "\n".join(text_parts)
        print(f"[Vision] 이미지 분석 완료: {original_filename} ({len(result)} chars)")
        return result

    except Exception as e:
        print(f"[Vision] 이미지 분석 실패 ({original_filename}): {e}")
        return ""


# ============ 템플릿 카탈로그 빌드 ============

def _build_slides_description(slides_meta: list[dict]) -> str:
    """템플릿 슬라이드 카탈로그를 LLM이 이해할 수 있는 텍스트로 변환"""

    # 템플릿에 등록된 타입 확인
    type_check = {
        "title_slide": False,
        "toc": False,
        "section_divider": False,
        "body": False,
        "closing": False,
    }
    for sm in slides_meta:
        ct = sm.get("slide_meta", {}).get("content_type", "body")
        if ct in type_check:
            type_check[ct] = True

    type_labels = {
        "title_slide": ("타이틀 슬라이드", "title"),
        "toc": ("목차 슬라이드", "toc"),
        "section_divider": ("섹션 간지", "section"),
        "body": ("본문 슬라이드", "content"),
        "closing": ("마지막 슬라이드", "closing"),
    }

    lines = [
        "아래는 사용 가능한 템플릿 슬라이드 카탈로그입니다.",
        "각 템플릿을 여러 번 사용하거나, 사용하지 않을 수 있습니다.",
        "",
        "## 템플릿 타입 현황",
    ]

    missing_types = []
    for ct, (label, schema_type) in type_labels.items():
        if type_check[ct]:
            lines.append(f"- ✓ {label} (type: {schema_type}) → 사용 가능")
        else:
            lines.append(f"- ✗ {label} (type: {schema_type}) → 템플릿 미등록")
            missing_types.append(schema_type)

    if missing_types:
        lines.append("")
        lines.append(f"**중요: 다음 타입은 템플릿이 없으므로 절대 생성하지 마세요: {', '.join(missing_types)}**")
        lines.append("해당 타입의 슬라이드는 slides 배열에 포함하지 마세요.")

    lines.append("")
    lines.append("## 상세 카탈로그")
    lines.append("")

    for sm in slides_meta:
        idx = sm["slide_index"]
        meta = sm.get("slide_meta", {})
        content_type = meta.get("content_type", "body")
        layout = meta.get("layout", "")

        # content_type → schema type 변환
        schema_type = CONTENT_TYPE_TO_SCHEMA_TYPE.get(content_type, "content")

        type_label = {
            "title_slide": "타이틀 슬라이드 (type: title)",
            "toc": "목차 슬라이드 (type: toc)",
            "section_divider": "섹션 간지 슬라이드 (type: section)",
            "body": "본문 슬라이드 (type: content)",
            "closing": "마무리 슬라이드 (type: closing)",
        }.get(content_type, "본문 슬라이드 (type: content)")

        has_title = meta.get("has_title", False)
        has_governance = meta.get("has_governance", False)
        subtitle_count = meta.get("subtitle_count", 0)
        desc_count = meta.get("description_count", 0)

        placeholders = sm.get("placeholders", [])
        ph_desc = []
        for ph in placeholders:
            role = ph.get("role", "")
            name = ph.get("placeholder", "")
            role_label = {
                "title": "제목 → title 필드 매핑",
                "subtitle": "부제목 → items[].heading 매핑 (순서대로)",
                "governance": "거버넌스 → governance/section_num 매핑",
                "body": "본문 → message/body_text 매핑",
                "description": "설명 → items[].detail 매핑 (순서대로)",
                "number": "번호 → items 순서 번호 자동 매핑 (LLM이 생성하지 않음)",
                "table": "표 → table_data 필드로 매핑 (LLM이 표 데이터를 생성해야 함)",
                "chart": "차트 → chart_data 필드로 매핑 (LLM이 차트 데이터를 생성해야 함)",
            }.get(role, role or "텍스트")
            ph_desc.append(f'    - placeholder: "{name}" (역할: {role_label})')

        lines.append(f"[템플릿 {idx}] {type_label}")
        if layout:
            lines.append(f"  레이아웃: {layout}")
        # 실제 placeholder 수 기반으로 effective_slots 계산
        ph_sub_actual = sum(1 for ph in placeholders if ph.get("role") == "subtitle")
        ph_desc_actual = sum(1 for ph in placeholders if ph.get("role") == "description")
        if ph_sub_actual > 0 and ph_desc_actual > 0:
            effective_items = min(ph_sub_actual, ph_desc_actual)
        else:
            effective_items = max(ph_sub_actual, ph_desc_actual, subtitle_count, desc_count)
        effective_items = min(effective_items, 4)  # 슬라이드당 최대 4개 항목 제한

        lines.append(f"  특성: has_title={has_title}, has_governance={has_governance}, subtitle_count={ph_sub_actual}, description_count={ph_desc_actual}")
        if content_type == "toc":
            # 목차 슬라이드: 슬롯 수에 맞게 항목 수 제한
            toc_slots = max(ph_sub_actual, ph_desc_actual)
            if toc_slots > 0:
                lines.append(f"  ※ 목차 슬라이드: 이 템플릿은 최대 {toc_slots}개 항목을 표시할 수 있습니다. 목차 items를 반드시 {toc_slots}개 이하로 생성하세요. 섹션이 {toc_slots}개보다 많으면 관련 섹션끼리 그룹핑하여 {toc_slots}개로 맞추세요.")
        elif content_type == "body":
            # 슬라이드에 표시 가능한 항목 수 안내 (다양한 items 수 허용)
            lines.append(f"  ※ 본문 슬라이드: 이 템플릿은 {effective_items}개 항목을 표시할 수 있습니다. items를 {effective_items}개 생성하세요.")
        elif content_type == "section_divider":
            # 간지 슬라이드: 부제목 필드 유무 안내
            has_subtitle_ph = ph_sub_actual > 0
            if has_subtitle_ph:
                lines.append(f"  ※ 간지 슬라이드: section_title(제목)은 필수입니다. 부제목 있음 → section_subtitle도 생성하세요.")
            else:
                lines.append(f"  ※ 간지 슬라이드: section_title(제목)은 필수입니다. 부제목 없음 → section_subtitle은 생성하지 마세요.")
        # 표/차트 오브젝트 안내
        ph_table_count = sum(1 for ph in placeholders if ph.get("role") == "table")
        ph_chart_count = sum(1 for ph in placeholders if ph.get("role") == "chart")
        if ph_table_count > 0:
            lines.append(f"  ※ 표(table) {ph_table_count}개 포함: 이 템플릿을 선택할 경우 반드시 table_data를 생성하세요.")
        if ph_chart_count > 0:
            lines.append(f"  ※ 차트(chart) {ph_chart_count}개 포함: 이 템플릿을 선택할 경우 반드시 chart_data를 생성하세요.")
        lines.append(f"  placeholder 목록:")
        if ph_desc:
            lines.extend(ph_desc)
        else:
            lines.append("    - (placeholder 없음)")
        lines.append("")

    return "\n".join(lines)


# ============ 리치 스키마 파싱 ============

def _parse_rich_schema(response_text: str, slides_meta: list[dict]) -> dict | None:
    """LLM 응답에서 리치 스키마 JSON을 파싱하고 placeholder 매핑으로 변환"""
    text = response_text.strip()
    print(f"[LLM] 응답 길이: {len(text)} chars")

    # ```json ... ``` 블록 추출
    if "```json" in text:
        start = text.index("```json") + 7
        try:
            end = text.index("```", start)
            text = text[start:end].strip()
        except ValueError:
            # 닫는 ``` 없음 - max_tokens 초과로 잘린 응답
            text = text[start:].strip()
            print(f"[LLM] Warning: JSON 블록이 닫히지 않음 (max_tokens 초과 가능성)")
    elif "```" in text:
        start = text.index("```") + 3
        try:
            end = text.index("```", start)
            text = text[start:end].strip()
        except ValueError:
            text = text[start:].strip()

    parsed = None

    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        for start_char, end_char in [("{", "}"), ("[", "]")]:
            if start_char in text:
                try:
                    s = text.index(start_char)
                    e = text.rindex(end_char) + 1
                    parsed = json.loads(text[s:e])
                    break
                except (json.JSONDecodeError, ValueError):
                    continue

    # JSON 파싱 실패 시 잘린 JSON 복구 시도
    if parsed is None:
        print(f"[LLM] JSON 파싱 실패 - 잘린 JSON 복구 시도")
        repaired = _try_repair_truncated_json(text)
        if repaired is not None:
            parsed = repaired
            print(f"[LLM] 잘린 JSON 복구 성공")
        else:
            print(f"[LLM] JSON 복구 실패 - 응답 앞 500자: {text[:500]}")
            print(f"[LLM] 응답 뒤 300자: {text[-300:]}")
            return None

    # 리치 스키마 형식: {"meta": {...}, "slides": [...], "sources": [...]}
    if isinstance(parsed, dict) and "slides" in parsed:
        schema_slides = parsed.get("slides", [])
        if isinstance(schema_slides, list) and schema_slides:
            # 첫 번째 항목에 "type" 필드가 있으면 리치 스키마
            if "type" in schema_slides[0]:
                # 목차(toc) items를 section 제목으로 보정
                _ensure_toc_items(schema_slides)
                # 리치 스키마 → placeholder 매핑 (내부에서 items 수 기반 템플릿 매칭 수행)
                mapped = _map_rich_schema_to_contents(schema_slides, slides_meta)
                return {
                    "slides": mapped,
                    "meta": parsed.get("meta", {}),
                    "sources": parsed.get("sources", []),
                    "raw_slides": schema_slides,
                }
            else:
                # 레거시 형식 (template_index + contents)
                validated = _validate_and_normalize(schema_slides, len(slides_meta))
                return {
                    "slides": validated,
                    "meta": parsed.get("meta", {}),
                    "sources": parsed.get("sources", []),
                }

    # 레거시 형식: [{"template_index": 0, "contents": {...}}]
    if isinstance(parsed, list):
        if parsed and "slide_index" in parsed[0]:
            converted = [
                {"template_index": item.get("slide_index", 0), "contents": item.get("contents", {})}
                for item in parsed
            ]
            return {"slides": converted, "meta": {}, "sources": []}
        if parsed and "template_index" in parsed[0]:
            validated = _validate_and_normalize(parsed, len(slides_meta))
            return {"slides": validated, "meta": {}, "sources": []}

    return None


def _try_repair_truncated_json(text: str) -> dict | None:
    """max_tokens 초과로 잘린 JSON 복구 시도

    슬라이드 배열이 중간에 잘린 경우, 완성된 슬라이드까지만 추출합니다.
    """
    # "slides" 배열 시작점 찾기
    if '"slides"' not in text:
        return None

    try:
        # { 부터 시작하는 JSON 찾기
        brace_start = text.index("{")
        partial = text[brace_start:]

        # slides 배열 내에서 완성된 마지막 슬라이드 객체 찾기
        # 패턴: }, { 또는 } ] 로 끝나는 슬라이드 경계
        # 역방향으로 완전한 } 찾기
        depth = 0
        last_valid_end = -1
        in_string = False
        escape_next = False

        for i, ch in enumerate(partial):
            if escape_next:
                escape_next = False
                continue
            if ch == '\\':
                escape_next = True
                continue
            if ch == '"':
                in_string = not in_string
                continue
            if in_string:
                continue
            if ch == '{':
                depth += 1
            elif ch == '}':
                depth -= 1
                if depth == 0:
                    # 최상위 객체가 닫힘 - 완전한 JSON
                    try:
                        return json.loads(partial[:i+1])
                    except json.JSONDecodeError:
                        pass

        # 최상위가 닫히지 않은 경우: slides 배열에서 마지막 완전한 슬라이드까지 추출
        # "slides": [ ... 에서 마지막 완성된 객체까지 잘라서 닫기
        slides_match = re.search(r'"slides"\s*:\s*\[', partial)
        if not slides_match:
            return None

        arr_start = slides_match.end()
        # 역방향으로 마지막 완전한 } 찾기 (배열 원소 경계)
        last_complete = partial.rfind('}')
        if last_complete <= arr_start:
            return None

        # slides 배열을 잘라서 닫기
        truncated_slides_str = partial[arr_start:last_complete + 1]
        # 마지막에 쉼표 제거
        truncated_slides_str = truncated_slides_str.rstrip().rstrip(',')

        # meta 추출 시도
        meta_match = re.search(r'"meta"\s*:\s*(\{[^}]*\})', partial)
        meta = {}
        if meta_match:
            try:
                meta = json.loads(meta_match.group(1))
            except json.JSONDecodeError:
                pass

        # slides 배열 파싱
        try:
            slides = json.loads(f"[{truncated_slides_str}]")
            if isinstance(slides, list) and len(slides) > 0:
                print(f"[LLM] 잘린 JSON에서 {len(slides)}개 슬라이드 복구")
                return {"meta": meta, "slides": slides, "sources": []}
        except json.JSONDecodeError:
            pass

    except (ValueError, IndexError):
        pass

    return None


def _ensure_minimum_items_for_slide(slide: dict, target_count: int):
    """단일 content 슬라이드의 items를 target_count까지 보충

    선택된 템플릿의 슬롯 수(effective_slots)를 target으로 받아,
    items가 부족한 경우 detail 텍스트를 분할하여 보충합니다.
    """
    items = slide.get("items", [])
    if len(items) >= target_count:
        return  # 이미 충분

    if len(items) == 1:
        # 1개 item의 detail을 문장 단위로 분할 시도
        item = items[0]
        detail = item.get("detail", "")

        # 문장 분할 (다양한 한국어/영어 문장 종결 패턴)
        sentences = [s.strip() for s in re.split(r'(?<=[.!?。])\s+', detail) if s.strip() and len(s.strip()) > 5]

        if len(sentences) >= target_count:
            new_items = []
            per_group = len(sentences) // target_count
            remainder = len(sentences) % target_count
            idx = 0
            for i in range(target_count):
                count = per_group + (1 if i < remainder else 0)
                group_sentences = sentences[idx:idx + count]
                new_detail = " ".join(group_sentences)
                new_heading = group_sentences[0][:30].rstrip('.!?。 ') if group_sentences else f"포인트 {i+1}"
                new_items.append({"heading": new_heading, "detail": new_detail})
                idx += count
            slide["items"] = new_items
        else:
            # 문장이 부족하면 줄바꿈/콤마 기준으로 분할 시도
            parts = [p.strip() for p in re.split(r'[\n,;·•\-]', detail) if p.strip() and len(p.strip()) > 5]
            if len(parts) >= target_count:
                new_items = []
                per_group = len(parts) // target_count
                remainder = len(parts) % target_count
                idx = 0
                for i in range(target_count):
                    count = per_group + (1 if i < remainder else 0)
                    group = parts[idx:idx + count]
                    new_detail = ". ".join(group)
                    new_heading = group[0][:30].rstrip('.!?。 ') if group else f"포인트 {i+1}"
                    new_items.append({"heading": new_heading, "detail": new_detail})
                    idx += count
                slide["items"] = new_items
            else:
                print(f"[LLM] Info: content slide items={len(items)}, template slots={target_count}, keeping as-is")

    elif len(items) == 2 and target_count >= 3:
        # 2개인 경우 3개까지 보충 시도 (긴 detail 분할)
        longest_idx = 0 if len(items[0].get("detail", "")) >= len(items[1].get("detail", "")) else 1
        longest_detail = items[longest_idx].get("detail", "")
        sentences = [s.strip() for s in re.split(r'(?<=[.!?。])\s+', longest_detail) if s.strip() and len(s.strip()) > 5]
        if len(sentences) >= 2:
            mid = len(sentences) // 2
            first_half = " ".join(sentences[:mid])
            second_half = " ".join(sentences[mid:])
            items[longest_idx]["detail"] = first_half
            new_item = {
                "heading": sentences[mid][:30].rstrip('.!?。 '),
                "detail": second_half,
            }
            items.insert(longest_idx + 1, new_item)
            slide["items"] = items


def _ensure_toc_items(schema_slides: list[dict]):
    """목차(toc) 슬라이드의 items를 section 슬라이드 제목으로 보정

    LLM이 toc items의 text를 비워두거나 toc 자체에 items가 없는 경우,
    뒤따르는 section 슬라이드들의 section_title을 수집하여 목차를 채웁니다.
    toc가 없고 section이 있으면 toc가 필요한 상황이므로 별도 처리하지 않습니다.
    """
    # section 슬라이드 제목 수집
    sections = []
    for slide in schema_slides:
        if slide.get("type") == "section":
            title = slide.get("section_title", "").strip()
            num = slide.get("section_num", "").strip()
            if title:
                sections.append({"num": num or f"{len(sections)+1:02d}", "text": title})

    if not sections:
        return

    for slide in schema_slides:
        if slide.get("type") != "toc":
            continue

        items = slide.get("items", [])

        # case 1: items가 없으면 section에서 생성
        if not items:
            slide["items"] = [{"num": s["num"], "text": s["text"]} for s in sections]
            print(f"[LLM] TOC: items 없음 → section {len(sections)}개에서 생성")
            continue

        # case 2: items가 있지만 text가 비어있으면 section에서 채우기
        needs_fix = any(not item.get("text", "").strip() for item in items)
        if needs_fix:
            if len(items) == len(sections):
                # 개수 동일 → 1:1 매핑
                for i, item in enumerate(items):
                    if not item.get("text", "").strip():
                        item["text"] = sections[i]["text"]
                        if not item.get("num", "").strip():
                            item["num"] = sections[i]["num"]
                print(f"[LLM] TOC: 빈 text를 section 제목으로 보정 ({len(items)}개)")
            else:
                # 개수 불일치 → section 기준으로 재생성
                slide["items"] = [{"num": s["num"], "text": s["text"]} for s in sections]
                print(f"[LLM] TOC: items {len(items)}개 ≠ section {len(sections)}개 → 재생성")


# ============ 리치 스키마 → placeholder 매핑 ============

def _map_rich_schema_to_contents(schema_slides: list[dict], slides_meta: list[dict]) -> list[dict]:
    """리치 스키마 슬라이드 배열을 template_index + contents 형식으로 변환"""
    num_templates = len(slides_meta)
    result = []

    # 본문(body) 템플릿을 effective_slots 기준으로 그룹화 (동일 용량 템플릿 라운드로빈용)
    _body_groups = {}  # effective_slots → [slide_index, ...]
    # 표/차트가 있는 본문 템플릿 인덱스 사전 수집
    _templates_with_chart = set()
    _templates_with_table = set()
    for sm in slides_meta:
        meta = sm.get("slide_meta", {})
        if meta.get("content_type") == "body":
            phs = sm.get("placeholders", [])
            ph_sub = sum(1 for ph in phs if ph.get("role") == "subtitle")
            ph_desc = sum(1 for ph in phs if ph.get("role") == "description")
            effective = min(ph_sub, ph_desc) if (ph_sub > 0 and ph_desc > 0) else max(ph_sub, ph_desc)
            _body_groups.setdefault(effective, []).append(sm["slide_index"])
            # 표/차트 포함 여부 기록
            if any(ph.get("role") == "chart" for ph in phs):
                _templates_with_chart.add(sm["slide_index"])
            if any(ph.get("role") == "table" for ph in phs):
                _templates_with_table.add(sm["slide_index"])

    # 그룹별 각 템플릿 리스트를 셔플 (랜덤 시작점)
    for slots in _body_groups:
        random.shuffle(_body_groups[slots])

    _group_counters = {slots: 0 for slots in _body_groups}

    for slide in schema_slides:
        slide_type = slide.get("type", "content")
        template_idx = slide.get("template_index")

        items_count = len(slide.get("items", [])) if slide_type == "content" else 0

        # template_index 유효성 검증
        if template_idx is not None and num_templates > 0:
            template_idx = max(0, min(template_idx, num_templates - 1))

        # content 슬라이드: items 수에 맞는 최적 템플릿 강제 매칭
        # (LLM이 잘못된 template_index를 선택해도 items_count 기반으로 최적 템플릿 자동 선택)
        has_chart_data = bool(slide.get("chart_data")) if slide_type == "content" else False
        has_table_data = bool(slide.get("table_data")) if slide_type == "content" else False

        if slide_type == "content" and items_count > 0:
            template_idx = _find_best_template_for_type(slide_type, slides_meta, items_count)
        elif template_idx is None:
            # 비-content 슬라이드: template_index가 없으면 자동 매칭
            template_idx = _find_best_template_for_type(slide_type, slides_meta, items_count)

        # 표/차트 데이터가 있는 슬라이드는 해당 오브젝트가 있는 템플릿으로 재매칭
        if slide_type == "content" and (has_chart_data or has_table_data):
            need_chart = has_chart_data and template_idx not in _templates_with_chart
            need_table = has_table_data and template_idx not in _templates_with_table
            if need_chart or need_table:
                # 표/차트가 있는 본문 템플릿 중에서 가장 적합한 것 선택
                best_idx = None
                best_score = -1
                for sm in slides_meta:
                    if sm.get("slide_meta", {}).get("content_type") != "body":
                        continue
                    si = sm["slide_index"]
                    match = False
                    if need_chart and si in _templates_with_chart:
                        match = True
                    if need_table and si in _templates_with_table:
                        match = True
                    if not match:
                        continue
                    # items 수용 가능성 점수 계산
                    phs = sm.get("placeholders", [])
                    ps = sum(1 for ph in phs if ph.get("role") == "subtitle")
                    pd = sum(1 for ph in phs if ph.get("role") == "description")
                    eff = min(ps, pd) if (ps > 0 and pd > 0) else max(ps, pd)
                    score = 10 if eff == items_count else (5 if eff >= items_count else 1)
                    if score > best_score:
                        best_score = score
                        best_idx = si
                if best_idx is not None:
                    template_idx = best_idx
                    print(f"[LLM] 표/차트 데이터 → 템플릿 {template_idx} 재매칭")

        # 동일 용량 본문 템플릿이 여러 개일 때 라운드로빈으로 돌아가며 사용
        # (표/차트 재매칭된 경우는 라운드로빈 건너뜀)
        if slide_type == "content" and not (has_chart_data or has_table_data):
            for sm in slides_meta:
                if sm["slide_index"] == template_idx:
                    phs = sm.get("placeholders", [])
                    ph_sub = sum(1 for ph in phs if ph.get("role") == "subtitle")
                    ph_desc = sum(1 for ph in phs if ph.get("role") == "description")
                    eff = min(ph_sub, ph_desc) if (ph_sub > 0 and ph_desc > 0) else max(ph_sub, ph_desc)
                    group = _body_groups.get(eff, [])
                    if len(group) > 1:
                        counter = _group_counters.get(eff, 0)
                        template_idx = group[counter % len(group)]
                        _group_counters[eff] = counter + 1
                    break
        elif slide_type == "content" and (has_chart_data or has_table_data):
            # 표/차트 템플릿은 라운드로빈에서 제외 (이미 재매칭됨)
            pass

        # 해당 템플릿의 placeholder 정보 찾기
        target_meta = None
        for sm in slides_meta:
            if sm["slide_index"] == template_idx:
                target_meta = sm
                break

        if not target_meta and slides_meta:
            target_meta = slides_meta[0]
            template_idx = target_meta["slide_index"]

        if not target_meta:
            continue

        # toc 슬라이드: 슬롯 수 초과 items 잘라내기 (LLM이 제한을 무시한 경우 안전장치)
        if slide_type == "toc":
            phs = target_meta.get("placeholders", [])
            ph_sub = sum(1 for ph in phs if ph.get("role") == "subtitle")
            ph_desc = sum(1 for ph in phs if ph.get("role") == "description")
            toc_slots = max(ph_sub, ph_desc)
            toc_items = slide.get("items", [])
            if toc_slots > 0 and len(toc_items) > toc_slots:
                print(f"[LLM] TOC: items {len(toc_items)}개 → 템플릿 슬롯 {toc_slots}개로 잘라냄")
                slide["items"] = toc_items[:toc_slots]

        # content 슬라이드: 선택된 템플릿의 슬롯 수에 맞게 items 조정
        if slide_type == "content":
            phs = target_meta.get("placeholders", [])
            ph_sub = sum(1 for ph in phs if ph.get("role") == "subtitle")
            ph_desc = sum(1 for ph in phs if ph.get("role") == "description")
            effective_slots = min(ph_sub, ph_desc) if (ph_sub > 0 and ph_desc > 0) else max(ph_sub, ph_desc)
            items = slide.get("items", [])
            # 슬롯보다 items가 많으면 잘라내기 (초과 items 안전장치)
            if effective_slots > 0 and len(items) > effective_slots:
                print(f"[LLM] Content: items {len(items)}개 → 템플릿 슬롯 {effective_slots}개로 잘라냄")
                slide["items"] = items[:effective_slots]
            # 슬롯보다 items가 적을 때만 보충 시도 (최소 슬롯 수 충족 목표)
            elif 0 < len(items) < effective_slots:
                _ensure_minimum_items_for_slide(slide, effective_slots)

        # 역할별 콘텐츠 매핑
        contents = _map_slide_content_to_placeholders(slide, target_meta)

        entry = {
            "template_index": template_idx,
            "contents": contents,
        }

        # content 슬라이드의 items 배열 보존 (프론트엔드 구조화 렌더링용)
        if slide.get("type") == "content" and slide.get("items"):
            entry["items"] = slide["items"]

        # 표/차트 데이터 보존
        if slide.get("table_data"):
            entry["table_data"] = slide["table_data"]
        if slide.get("chart_data"):
            entry["chart_data"] = slide["chart_data"]

        result.append(entry)

    return result


def _map_slide_content_to_placeholders(slide: dict, template_meta: dict) -> dict:
    """리치 스키마 슬라이드의 콘텐츠를 템플릿 placeholder에 매핑

    관리자가 오브젝트에 설정한 역할(role)에 따라 콘텐츠를 배치:
    - role=title → 슬라이드 제목
    - role=governance → 거버넌스/섹션번호
    - role=subtitle → 부제목/섹션부제목/meta_line
    - role=body → 본문 텍스트/메시지
    - role=description → 세부 항목 (items)
    """
    contents = {}
    placeholders = template_meta.get("placeholders", [])
    slide_type = slide.get("type", "content")

    # 역할별 placeholder 분류
    role_map = {}  # role → placeholder_name (단일)
    desc_placeholders = []  # description 역할은 여러 개 가능
    subtitle_placeholders = []  # subtitle 역할도 여러 개 가능
    number_placeholders = []  # number 역할은 여러 개 가능 (순서 번호 자동 매핑)

    for ph in placeholders:
        role = ph.get("role", "")
        name = ph.get("placeholder", "")
        if not name:
            continue
        if role == "description":
            desc_placeholders.append(name)
        elif role == "subtitle":
            subtitle_placeholders.append(name)
        elif role == "number":
            number_placeholders.append(name)
        elif role not in role_map:
            role_map[role] = name

    # ── 타입별 매핑 ──

    if slide_type == "title":
        _map_if("title", role_map, contents, slide.get("title", ""))
        if subtitle_placeholders and slide.get("subtitle"):
            contents[subtitle_placeholders[0]] = slide["subtitle"]
        _map_if("body", role_map, contents, slide.get("meta_line", ""))
        _map_if("governance", role_map, contents, slide.get("meta_line", ""))
        if desc_placeholders and slide.get("meta_line"):
            contents[desc_placeholders[0]] = slide["meta_line"]

    elif slide_type == "toc":
        _map_if("title", role_map, contents, slide.get("title", "목차"))
        items = slide.get("items", [])
        has_number_obj = len(number_placeholders) > 0

        # number placeholder에 순서 번호 자동 매핑 (두 자리: 01, 02, ...)
        for i, item in enumerate(items):
            if i < len(number_placeholders):
                contents[number_placeholders[i]] = item.get("num", str(i + 1).zfill(2))

        if subtitle_placeholders and desc_placeholders and items:
            for i, item in enumerate(items):
                num = item.get("num", str(i + 1).zfill(2))
                text = item.get("text", "")
                if has_number_obj:
                    # 숫자 오브젝트가 있으면 subtitle/desc에 텍스트만 배치
                    if i < len(subtitle_placeholders):
                        contents[subtitle_placeholders[i]] = text
                    if i < len(desc_placeholders):
                        contents[desc_placeholders[i]] = text
                else:
                    # 숫자 오브젝트가 없으면 "01. 목차명" 형식으로 번호 자동 추가
                    if i < len(subtitle_placeholders):
                        contents[subtitle_placeholders[i]] = num
                    if i < len(desc_placeholders):
                        contents[desc_placeholders[i]] = f"{num}. {text}" if num else text
        elif subtitle_placeholders and items:
            for i, item in enumerate(items):
                if i >= len(subtitle_placeholders):
                    break
                num = item.get("num", str(i + 1).zfill(2))
                text = item.get("text", "")
                if has_number_obj:
                    contents[subtitle_placeholders[i]] = text
                else:
                    contents[subtitle_placeholders[i]] = f"{num}. {text}" if num else text
        elif desc_placeholders and items:
            for i, item in enumerate(items):
                if i >= len(desc_placeholders):
                    break
                num = item.get("num", str(i + 1).zfill(2))
                text = item.get("text", "")
                if has_number_obj:
                    contents[desc_placeholders[i]] = text
                else:
                    contents[desc_placeholders[i]] = f"{num}. {text}" if num else text
        elif items:
            toc_lines = []
            for i, item in enumerate(items):
                num = item.get("num", str(i + 1).zfill(2))
                text = item.get("text", "")
                if has_number_obj:
                    toc_lines.append(text)
                else:
                    toc_lines.append(f"{num}. {text}" if num else text)
            if "body" in role_map:
                contents[role_map["body"]] = "\n".join(toc_lines)

    elif slide_type == "section":
        # 간지 슬라이드: 제목은 필수 매핑
        section_title = slide.get("section_title", "")
        if "title" in role_map:
            if section_title:
                contents[role_map["title"]] = section_title
        elif section_title:
            # title role이 없어도 첫 번째 빈 placeholder에 제목 매핑
            for ph in placeholders:
                name = ph.get("placeholder", "")
                role = ph.get("role", "")
                if name and name not in contents and role not in ("number",):
                    contents[name] = section_title
                    break

        # 부제목: subtitle placeholder가 있을 때만 매핑
        if subtitle_placeholders and slide.get("section_subtitle"):
            contents[subtitle_placeholders[0]] = slide["section_subtitle"]

        _map_if("governance", role_map, contents, slide.get("section_num", ""))
        _map_if("body", role_map, contents, slide.get("section_subtitle", ""))
        # number placeholder에 section_num 매핑
        if number_placeholders and slide.get("section_num"):
            contents[number_placeholders[0]] = slide["section_num"]

    elif slide_type == "content":
        # 제목, 거버넌스는 무조건 매핑 (필드가 있으면)
        _map_if("title", role_map, contents, slide.get("title", ""))
        _map_if("governance", role_map, contents, slide.get("governance", ""))

        items = slide.get("items", [])

        # number placeholder에 순서 번호 자동 매핑 (01, 02, 03, ...)
        for i in range(len(items)):
            if i < len(number_placeholders):
                contents[number_placeholders[i]] = str(i + 1).zfill(2)

        # 부제목/설명을 순서대로 1:1 매핑
        # items[i].heading → subtitle_placeholders[i]
        # items[i].detail  → desc_placeholders[i]
        for i, item in enumerate(items):
            heading = item.get("heading", "")
            detail = item.get("detail", item.get("body", ""))

            # 부제목 placeholder에 heading 매핑
            if i < len(subtitle_placeholders) and heading:
                contents[subtitle_placeholders[i]] = heading

            # 설명 placeholder에 detail 매핑
            if i < len(desc_placeholders) and detail:
                contents[desc_placeholders[i]] = detail

        # 부제목/설명 placeholder가 없는 경우 fallback:
        # body placeholder에 전체 내용 합쳐서 배치
        if not subtitle_placeholders and not desc_placeholders and items:
            text_blocks = []
            for item in items:
                heading = item.get("heading", "")
                detail = item.get("detail", item.get("body", ""))
                if heading and detail:
                    text_blocks.append(f"{heading}\n{detail}")
                elif heading or detail:
                    text_blocks.append(heading or detail)

            if "body" in role_map:
                contents[role_map["body"]] = "\n\n".join(text_blocks)
            else:
                # placeholder가 전혀 없으면 첫 번째 빈 placeholder에 합치기
                for ph in placeholders:
                    name = ph.get("placeholder", "")
                    if name and name not in contents:
                        contents[name] = "\n\n".join(text_blocks)
                        break

        # 표(table) 데이터 매핑
        table_phs = [ph["placeholder"] for ph in placeholders if ph.get("role") == "table"]
        if table_phs and slide.get("table_data"):
            td = slide["table_data"]
            if isinstance(td, dict):
                headers = td.get("headers", [])
                rows = td.get("rows", [])
                data_2d = [headers] + rows if headers else rows
                if data_2d and table_phs:
                    contents[table_phs[0]] = {
                        "data": data_2d,
                        "rows": len(data_2d),
                        "cols": len(data_2d[0]) if data_2d else 0,
                    }
            elif isinstance(td, list):
                for i, tdi in enumerate(td):
                    if i >= len(table_phs):
                        break
                    headers = tdi.get("headers", [])
                    rows = tdi.get("rows", [])
                    data_2d = [headers] + rows if headers else rows
                    if data_2d:
                        contents[table_phs[i]] = {
                            "data": data_2d,
                            "rows": len(data_2d),
                            "cols": len(data_2d[0]) if data_2d else 0,
                        }

        # 차트(chart) 데이터 매핑
        chart_phs = [ph["placeholder"] for ph in placeholders if ph.get("role") == "chart"]
        if chart_phs and slide.get("chart_data"):
            cd = slide["chart_data"]
            if isinstance(cd, dict):
                if chart_phs:
                    contents[chart_phs[0]] = {
                        "chart_type": cd.get("chart_type", "bar"),
                        "title": cd.get("title", ""),
                        "chart_data": cd.get("chart_data", {}),
                    }
            elif isinstance(cd, list):
                for i, cdi in enumerate(cd):
                    if i >= len(chart_phs):
                        break
                    contents[chart_phs[i]] = {
                        "chart_type": cdi.get("chart_type", "bar"),
                        "title": cdi.get("title", ""),
                        "chart_data": cdi.get("chart_data", {}),
                    }

        # sources → 별도 placeholder가 없으므로 무시 (meta에 저장)

    elif slide_type == "closing":
        _map_if("title", role_map, contents, slide.get("title", "감사합니다"))
        _map_if("body", role_map, contents, slide.get("message", ""))
        if subtitle_placeholders and slide.get("contact"):
            contents[subtitle_placeholders[0]] = slide["contact"]
        if desc_placeholders and slide.get("contact"):
            contents[desc_placeholders[0]] = slide["contact"]

    # 빈 placeholder에 대한 처리:
    # - subtitle/description/number: 매핑 안 된 경우 contents에 포함하지 않음
    #   → _build_gen_objects에서 generated_text=None 판정 → 오브젝트 제거
    # - 그 외 역할 (title/governance/body 등): 빈 문자열로 유지
    for ph in placeholders:
        name = ph.get("placeholder", "")
        role = ph.get("role", "")
        if name and name not in contents and role not in ("number", "subtitle", "description"):
            contents[name] = ""

    return contents


def _map_if(role: str, role_map: dict, contents: dict, value: str):
    """역할이 존재하면 매핑"""
    if role in role_map and value:
        contents[role_map[role]] = value


def _find_best_template_for_type(
    slide_type: str, slides_meta: list[dict], items_count: int = 0
) -> int:
    """슬라이드 타입에 가장 적합한 템플릿 인덱스 찾기

    content 슬라이드의 경우 subtitle과 description placeholder가 모두
    items 개수를 수용할 수 있는 템플릿을 우선 선택합니다.
    동일 점수의 후보가 여러 개이면 랜덤으로 선택합니다.
    """
    target_content_type = SCHEMA_TYPE_TO_CONTENT_TYPE.get(slide_type, "body")

    candidates = []  # (score, slide_index)
    best_score = -1

    for sm in slides_meta:
        meta = sm.get("slide_meta", {})
        content_type = meta.get("content_type", "body")
        score = 0

        # content_type 정확 일치
        if content_type == target_content_type:
            score += 10

        # layout 일치 보너스
        layout = meta.get("layout", "")
        if slide_type == "content" and layout in ("two_column", "grid", "single_column"):
            score += 3
        elif slide_type == "toc" and layout in ("numbered_list", "list"):
            score += 3
        elif slide_type == "section" and layout == "divider":
            score += 3

        # content 슬라이드: subtitle/description placeholder 개수와 items 개수 매칭
        if slide_type == "content" and items_count > 0:
            placeholders = sm.get("placeholders", [])
            ph_desc_count = sum(1 for ph in placeholders if ph.get("role") == "description")
            ph_sub_count = sum(1 for ph in placeholders if ph.get("role") == "subtitle")

            # subtitle과 description 둘 다 있으면, 둘 다 items를 수용해야 함 (min 기준)
            # 한쪽만 있으면 그쪽 기준
            if ph_sub_count > 0 and ph_desc_count > 0:
                effective_slots = min(ph_sub_count, ph_desc_count)
            else:
                effective_slots = max(ph_sub_count, ph_desc_count)

            if effective_slots == items_count:
                score += 8  # 정확히 일치 (total: 18)
            elif effective_slots >= items_count:
                score += 5  # placeholder가 더 많음, 수용 가능 (total: 15)
            elif effective_slots > 0 and effective_slots < items_count:
                score -= 8  # placeholder 부족, body(2) > non-body(0) 유지 (total: 2)

            # 표/차트가 있는 템플릿은 일반 선택 시 감점 (표/차트 데이터 없이 배정되면 기본값 노출)
            has_data_obj = any(ph.get("role") in ("table", "chart") for ph in placeholders)
            if has_data_obj:
                score -= 3  # 표/차트 데이터 없는 슬라이드에는 비선호

        if score > best_score:
            best_score = score
            candidates = [sm["slide_index"]]
        elif score == best_score:
            candidates.append(sm["slide_index"])

    return random.choice(candidates) if candidates else 0


# ============ 유효성 검증 (레거시 호환) ============

def _validate_and_normalize(slides: list[dict], num_templates: int) -> list[dict]:
    """template_index 유효성 검증 및 정규화"""
    result = []
    for slide in slides:
        template_idx = slide.get("template_index", 0)
        if num_templates > 0:
            template_idx = max(0, min(template_idx, num_templates - 1))
        result.append({
            "template_index": template_idx,
            "contents": slide.get("contents", {}),
        })
    return result


# ============ 폴백 콘텐츠 생성 ============

async def generate_single_slide_content(
    resources_text: str,
    instruction: str,
    slide_meta: dict,
    lang: str = "ko",
    current_content: dict = None,
) -> dict:
    """단일 슬라이드용 텍스트 콘텐츠 생성/수정

    current_content가 있으면 기존 내용을 바탕으로 사용자 지침에 따라 수정합니다.
    없으면 새로 생성합니다.

    Returns:
        {"contents": {placeholder: text}, "items": [{"heading": ..., "detail": ...}]}
    """
    if not settings.ANTHROPIC_API_KEY or settings.ANTHROPIC_API_KEY == "your_anthropic_api_key_here":
        # 폴백: placeholder마다 지침 텍스트 반환
        contents = {}
        for p in slide_meta.get("placeholders", []):
            contents[p["placeholder"]] = instruction or "내용 없음"
        return {"contents": contents, "items": []}

    placeholders = slide_meta.get("placeholders", [])
    meta = slide_meta.get("slide_meta", {})
    desc_count = meta.get("description_count", 3)

    placeholder_desc = "\n".join(
        f'  - "{p["placeholder"]}" (역할: {p["role"]})'
        for p in placeholders
    )

    lang_instruction = {
        "ko": "한국어로 작성하세요.",
        "en": "Write in English.",
        "ja": "日本語で作成してください。",
        "zh": "请用中文撰写。",
    }.get(lang, "한국어로 작성하세요.")

    # 기존 내용이 있는 경우: 수정 모드
    has_existing = current_content and (current_content.get("contents") or current_content.get("items"))

    if has_existing:
        existing_contents = current_content.get("contents", {})
        existing_items = current_content.get("items", [])

        existing_text = ""
        if existing_contents:
            existing_text += "### 현재 텍스트 내용:\n"
            for k, v in existing_contents.items():
                if v:
                    existing_text += f'  - {k}: "{v}"\n'
        if existing_items:
            existing_text += "### 현재 항목(items):\n"
            for i, item in enumerate(existing_items):
                existing_text += f'  {i+1}. heading: "{item.get("heading", "")}", detail: "{item.get("detail", "")}"\n'

        system_prompt = f"""당신은 프레젠테이션 콘텐츠 편집 전문가입니다.
사용자의 지침에 따라 현재 슬라이드의 내용을 수정합니다.
{lang_instruction}

## 중요 규칙:
- 사용자가 특정 부분만 수정 요청하면 해당 부분만 변경하고 나머지는 유지하세요.
- "제목을 바꿔줘" → title 역할의 텍스트만 변경
- "항목을 추가해줘" → 기존 items를 유지하고 새 항목 추가
- "삭제해줘" → 해당 항목 제거
- "전체를 다시 작성해줘" → 전체 새로 작성
- 지침이 명확하지 않으면 기존 내용을 최대한 유지하면서 개선하세요.
- **[필수] 응답 시 반드시 모든 placeholder에 대한 콘텐츠를 포함하세요. 특히 title(제목), subtitle(부제목), description(설명) 역할의 placeholder가 비어있으면 안 됩니다.**
- **[필수] items 배열에는 반드시 heading(소제목)과 detail(설명)을 모두 포함해야 합니다. heading이나 detail이 빈 문자열이면 안 됩니다.**

반드시 아래 JSON 형식으로만 응답하세요:
{{
  "contents": {{ "placeholder_name": "텍스트 내용", ... }},
  "items": [ {{ "heading": "소제목", "detail": "설명 내용" }}, ... ]
}}
"""

        user_prompt = f"""## 현재 슬라이드 내용 (수정 대상)
{existing_text}

## 참고 리소스
{resources_text[:6000]}

## 사용자 수정 지침
{instruction}

## 슬라이드 정보
- 콘텐츠 타입: {meta.get('content_type', 'body')}
- 설명 항목 수: {desc_count}개

## Placeholder 목록
{placeholder_desc}

위 현재 내용을 사용자 지침에 따라 수정하여 JSON으로 응답하세요.
items 배열의 heading은 subtitle, detail은 description에 대응합니다.
**반드시 title(제목), items의 heading(부제목)과 detail(설명)을 모두 포함하여 응답하세요. 빈 값이 없어야 합니다.**
"""
    else:
        system_prompt = f"""당신은 프레젠테이션 콘텐츠 전문가입니다.
주어진 리소스와 지침을 바탕으로 슬라이드 1장의 텍스트를 생성합니다.
{lang_instruction}

## 중요 규칙:
- **[필수] 모든 placeholder에 대한 콘텐츠를 생성하세요. title(제목), subtitle(부제목), description(설명) 역할의 placeholder를 빠뜨리지 마세요.**
- **[필수] items 배열에는 반드시 heading(소제목)과 detail(설명)을 모두 포함해야 합니다. heading이나 detail이 빈 문자열이면 안 됩니다.**
- governance(거버넌스)가 있으면 슬라이드 내용을 요약한 문장을 작성하세요.

반드시 아래 JSON 형식으로만 응답하세요:
{{
  "contents": {{ "placeholder_name": "텍스트 내용", ... }},
  "items": [ {{ "heading": "소제목", "detail": "설명 내용" }}, ... ]
}}
"""

        user_prompt = f"""## 리소스 (참고 자료)
{resources_text[:8000]}

## 사용자 지침
{instruction}

## 이 슬라이드 정보
- 콘텐츠 타입: {meta.get('content_type', 'body')}
- 레이아웃: {meta.get('layout', '')}
- 설명 항목 수: {desc_count}개

## Placeholder 목록 (모든 placeholder에 대해 contents를 생성하세요)
{placeholder_desc}

items 배열에는 heading(소제목)과 detail(설명)을 {desc_count}개 생성하세요.
subtitle 역할의 placeholder는 items의 heading에 대응하고, description 역할은 items의 detail에 대응합니다.
**반드시 title(제목), items의 heading(부제목)과 detail(설명)을 모두 포함하여 응답하세요. 빈 값이 없어야 합니다.**
"""

    try:
        single_model = await get_prompt_model("slide_generation_system") or settings.ANTHROPIC_MODEL
        result = await _call_claude_api(system_prompt, user_prompt, model=single_model)
        print(f"[LLM] 단일 슬라이드 응답 길이: {len(result)} chars (model: {single_model})")
        parsed = _extract_json(result)
        if parsed and "contents" in parsed:
            return parsed
        print(f"[LLM] JSON 파싱 실패, 원본: {result[:500]}")
        raise ValueError("AI 응답을 파싱할 수 없습니다")
    except Exception as e:
        print(f"[LLM] 단일 슬라이드 텍스트 생성 실패: {e}")
        raise


def _extract_json(text: str) -> dict:
    """텍스트에서 JSON 객체 추출"""
    import re
    # ```json ... ``` 블록 먼저 시도
    m = re.search(r'```(?:json)?\s*(\{[\s\S]*?\})\s*```', text)
    if m:
        try:
            return json.loads(m.group(1))
        except Exception:
            pass
    # 직접 JSON 파싱 시도
    m = re.search(r'\{[\s\S]*\}', text)
    if m:
        try:
            return json.loads(m.group(0))
        except Exception:
            pass
    return {}


def _fallback_content(resources_text: str, slides_meta: list[dict]) -> list[dict]:
    """API 키 미설정 또는 오류 시 템플릿 타입 기반 기본 콘텐츠 생성"""
    result = []
    if not slides_meta:
        return result

    # 텍스트를 문단 단위로 분할
    paragraphs = [p.strip() for p in resources_text.split("\n\n") if p.strip()]
    if not paragraphs:
        paragraphs = [p.strip() for p in resources_text.split("\n") if p.strip()]
    if not paragraphs:
        paragraphs = [resources_text[:200]] if resources_text else ["내용 없음"]

    # 템플릿 타입별 인덱스 수집
    title_idx = None
    toc_idx = None
    section_idx = None
    body_idx = None
    closing_idx = None

    for sm in slides_meta:
        ct = sm.get("slide_meta", {}).get("content_type", "body")
        if ct == "title_slide" and title_idx is None:
            title_idx = sm["slide_index"]
        elif ct == "toc" and toc_idx is None:
            toc_idx = sm["slide_index"]
        elif ct == "section_divider" and section_idx is None:
            section_idx = sm["slide_index"]
        elif ct == "body" and body_idx is None:
            body_idx = sm["slide_index"]
        elif ct == "closing" and closing_idx is None:
            closing_idx = sm["slide_index"]

    if body_idx is None:
        body_idx = slides_meta[0]["slide_index"]

    # 타이틀 슬라이드
    if title_idx is not None:
        result.append({
            "template_index": title_idx,
            "contents": _fill_placeholders(slides_meta, title_idx, paragraphs[0] if paragraphs else ""),
        })

    # 본문 슬라이드 (문단을 분배)
    body_paras = paragraphs[1:] if (title_idx is not None and len(paragraphs) > 1) else paragraphs
    num_body_slides = max(1, min(len(body_paras), 5))
    chunks = _distribute_paragraphs(body_paras, num_body_slides)

    for chunk in chunks:
        result.append({
            "template_index": body_idx,
            "contents": _fill_placeholders(slides_meta, body_idx, chunk),
        })

    # 마무리 슬라이드
    if closing_idx is not None:
        result.append({
            "template_index": closing_idx,
            "contents": _fill_placeholders(slides_meta, closing_idx, ""),
        })

    return result


def _fill_placeholders(slides_meta: list[dict], template_idx: int, chunk_text: str) -> dict:
    """주어진 템플릿의 placeholder를 역할에 맞게 채우기"""
    target_meta = None
    for sm in slides_meta:
        if sm["slide_index"] == template_idx:
            target_meta = sm
            break

    if not target_meta:
        return {}

    placeholders = target_meta.get("placeholders", [])
    contents = {}

    for ph in placeholders:
        name = ph.get("placeholder", "")
        role = ph.get("role", "")

        if not name:
            continue

        if role == "title":
            first_line = chunk_text.split("\n")[0] if chunk_text else "제목"
            contents[name] = first_line[:50]
        elif role == "governance":
            contents[name] = f"{settings.SOLUTION_NAME} 자동 생성"
        elif role == "subtitle":
            contents[name] = chunk_text[:80] if chunk_text else ""
        else:
            contents[name] = chunk_text[:300] if chunk_text else ""

    if not contents and placeholders:
        for ph in placeholders:
            name = ph.get("placeholder", "text")
            contents[name] = chunk_text[:300] if chunk_text else ""

    return contents


def _distribute_paragraphs(paragraphs: list[str], num_slides: int) -> list[str]:
    """문단 목록을 슬라이드 수에 맞게 균등 분배"""
    if num_slides <= 0:
        return []
    if len(paragraphs) <= num_slides:
        result = paragraphs + [""] * (num_slides - len(paragraphs))
        return result

    result = []
    chunk_size = len(paragraphs) / num_slides
    for i in range(num_slides):
        start = int(i * chunk_size)
        end = int((i + 1) * chunk_size)
        chunk = "\n".join(paragraphs[start:end])
        result.append(chunk)

    return result


# ============ 스트리밍 생성 ============

async def _build_generation_prompts(
    resources_text: str, instructions: str, slides_meta: list[dict], lang: str = "",
    slide_count: str = "auto",
) -> tuple[str, str]:
    """DB에서 프롬프트 템플릿을 로드하여 변수를 바인딩한 시스템/유저 프롬프트 생성"""
    slides_description = _build_slides_description(slides_meta)

    # 슬라이드 수 지침
    if slide_count and slide_count != "auto":
        slide_count_instruction = f"\n11. **[필수] 전체 슬라이드 수를 정확히 {slide_count}장으로 생성하세요.** title, toc, section, content, closing 모두 포함하여 총 {slide_count}장이어야 합니다."
    else:
        slide_count_instruction = ""

    output_lang = lang or (settings.SUPPORTED_LANGS[0] if settings.SUPPORTED_LANGS else "ko")
    lang_instruction = {
        "ko": "모든 콘텐츠를 한국어로 작성하세요.",
        "en": "Write all content in English.",
        "ja": "すべてのコンテンツを日本語で作成してください。",
        "zh": "请用中文撰写所有内容。",
    }.get(output_lang, "모든 콘텐츠를 한국어로 작성하세요.")

    # DB에서 프롬프트 템플릿 로드
    system_template = await get_prompt_content("slide_generation_system")
    user_template = await get_prompt_content("slide_generation_user")

    # 변수 바인딩
    system_prompt = system_template.format(
        lang_instruction=lang_instruction,
        slide_count_instruction=slide_count_instruction,
    )

    user_prompt = user_template.format(
        lang_instruction=lang_instruction,
        instructions=instructions if instructions else "특별한 지침 없음 - 리소스 내용을 보고서 형태로 정리해주세요.",
        resources_text=resources_text[:12000],
        slides_description=slides_description,
    )

    return system_prompt, user_prompt


async def _stream_claude_api(system_prompt: str, user_prompt: str, model: str = "", max_tokens: int = 0):
    """Claude API 스트리밍 호출 - text delta를 async yield"""
    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": settings.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    payload = {
        "model": model or settings.ANTHROPIC_MODEL,
        "system": system_prompt,
        "messages": [{"role": "user", "content": user_prompt}],
        "stream": True,
    }
    # 모델별 max_tokens 적용 (Anthropic API 필수 파라미터)
    effective_model = model or settings.ANTHROPIC_MODEL
    if max_tokens > 0:
        payload["max_tokens"] = max_tokens
    elif effective_model == settings.ANTHROPIC_OUTLINE_MODEL and settings.ANTHROPIC_OUTLINE_MAX_TOKENS > 0:
        payload["max_tokens"] = settings.ANTHROPIC_OUTLINE_MAX_TOKENS
    elif settings.ANTHROPIC_MAX_TOKENS > 0:
        payload["max_tokens"] = settings.ANTHROPIC_MAX_TOKENS
    else:
        payload["max_tokens"] = 4096  # 기본값

    # Extended output requires beta header
    if payload.get("max_tokens", 0) > 16384:
        headers["anthropic-beta"] = "interleaved-thinking-2025-05-14,output-128k-2025-02-19"

    async with httpx.AsyncClient(timeout=180.0) as client:
        async with client.stream("POST", url, json=payload, headers=headers) as response:
            response.raise_for_status()
            async for line in response.aiter_lines():
                if not line.startswith("data: "):
                    continue
                data_str = line[6:]
                if data_str.strip() == "[DONE]":
                    break
                try:
                    event = json.loads(data_str)
                    if event.get("type") == "content_block_delta":
                        delta = event.get("delta", {})
                        if delta.get("type") == "text_delta":
                            text = delta.get("text", "")
                            if text:
                                yield text
                except json.JSONDecodeError:
                    continue


async def generate_slide_content_stream(
    resources_text: str,
    instructions: str,
    slides_meta: list[dict],
    lang: str = "",
    slide_count: str = "auto",
):
    """스트리밍 방식 슬라이드 콘텐츠 생성

    Yields:
        ("delta", str)   - Claude에서 스트리밍된 텍스트 청크
        ("result", dict) - 최종 파싱 결과
    """
    if not settings.ANTHROPIC_API_KEY or settings.ANTHROPIC_API_KEY == "your_anthropic_api_key_here":
        yield ("result", {
            "slides": _fallback_content(resources_text, slides_meta),
            "meta": {},
            "sources": [],
        })
        return

    system_prompt, user_prompt = await _build_generation_prompts(
        resources_text, instructions, slides_meta, lang,
        slide_count=slide_count,
    )

    # 프롬프트별 모델 설정 조회
    prompt_model = await get_prompt_model("slide_generation_system")
    effective_model = prompt_model or settings.ANTHROPIC_OUTLINE_MODEL

    try:
        full_text = ""
        async for delta in _stream_claude_api(system_prompt, user_prompt, model=effective_model):
            full_text += delta
            yield ("delta", delta)

        print(f"[LLM] 스트리밍 완료 - 전체 응답 길이: {len(full_text)} chars (model: {effective_model})")
        parsed = _parse_rich_schema(full_text, slides_meta)
        if parsed:
            slide_count_result = len(parsed.get("slides", []))
            raw_count = len(parsed.get("raw_slides", []))
            print(f"[LLM] 파싱 성공 - slides: {slide_count_result}개, raw_slides: {raw_count}개")
            yield ("result", parsed)
        else:
            print(f"[LLM] 파싱 실패 → 폴백 콘텐츠 사용")
            yield ("result", {
                "slides": _fallback_content(resources_text, slides_meta),
                "meta": {},
                "sources": [],
            })
    except Exception as e:
        import traceback
        print(f"[LLM] Streaming 호출 실패: {e}")
        traceback.print_exc()
        yield ("result", {
            "slides": _fallback_content(resources_text, slides_meta),
            "meta": {},
            "sources": [],
        })


# ============ 엑셀 콘텐츠 생성 ============

async def generate_excel_content_stream(
    resources_text: str,
    instructions: str,
    lang: str = "",
    sheet_count: str = "auto",
):
    """스트리밍 방식 엑셀 콘텐츠 생성

    Yields:
        ("delta", str)   - Claude에서 스트리밍된 텍스트 청크
        ("result", dict) - 최종 파싱 결과: {"sheets": [...], "meta": {...}}
    """
    if not settings.ANTHROPIC_API_KEY or settings.ANTHROPIC_API_KEY == "your_anthropic_api_key_here":
        yield ("result", _excel_fallback(resources_text))
        return

    system_prompt, user_prompt = await _build_excel_generation_prompts(
        resources_text, instructions, lang, sheet_count=sheet_count,
    )

    # 프롬프트별 모델 설정 조회
    prompt_model = await get_prompt_model("excel_generation_system")
    effective_model = prompt_model or settings.ANTHROPIC_MODEL

    try:
        full_text = ""
        async for delta in _stream_claude_api(system_prompt, user_prompt, model=effective_model):
            full_text += delta
            yield ("delta", delta)

        print(f"[LLM-Excel] 스트리밍 완료 - 전체 응답 길이: {len(full_text)} chars (model: {effective_model})")
        parsed = _parse_excel_schema(full_text)
        if parsed:
            sheet_count_result = len(parsed.get("sheets", []))
            print(f"[LLM-Excel] 파싱 성공 - sheets: {sheet_count_result}개")
            yield ("result", parsed)
        else:
            print(f"[LLM-Excel] 파싱 실패 → 폴백 사용")
            yield ("result", _excel_fallback(resources_text))
    except Exception as e:
        import traceback
        print(f"[LLM-Excel] Streaming 호출 실패: {e}")
        traceback.print_exc()
        yield ("result", _excel_fallback(resources_text))


async def modify_excel_content_stream(
    current_data: dict,
    instruction: str,
    lang: str = "",
    target_sheet_index: int | None = None,
):
    """스트리밍 방식 엑셀 부분 수정

    Yields:
        ("delta", str)   - Claude에서 스트리밍된 텍스트 청크
        ("result", dict) - 최종 파싱 결과: {"sheets": [...], "meta": {...}}
    """
    if not settings.ANTHROPIC_API_KEY or settings.ANTHROPIC_API_KEY == "your_anthropic_api_key_here":
        yield ("result", current_data)
        return

    system_prompt, user_prompt = await _build_excel_modify_prompts(
        current_data, instruction, lang, target_sheet_index=target_sheet_index,
    )

    prompt_model = await get_prompt_model("excel_modify_system")
    effective_model = prompt_model or settings.ANTHROPIC_MODEL

    try:
        full_text = ""
        async for delta in _stream_claude_api(system_prompt, user_prompt, model=effective_model):
            full_text += delta
            yield ("delta", delta)

        print(f"[LLM-Excel-Modify] 스트리밍 완료 - 전체 응답 길이: {len(full_text)} chars (model: {effective_model})")
        parsed = _parse_excel_schema(full_text)
        if parsed:
            sheet_count_result = len(parsed.get("sheets", []))
            print(f"[LLM-Excel-Modify] 파싱 성공 - sheets: {sheet_count_result}개")
            yield ("result", parsed)
        else:
            print(f"[LLM-Excel-Modify] 파싱 실패 → 기존 데이터 유지")
            yield ("result", current_data)
    except Exception as e:
        import traceback
        print(f"[LLM-Excel-Modify] Streaming 호출 실패: {e}")
        traceback.print_exc()
        yield ("result", current_data)


async def _build_excel_modify_prompts(
    current_data: dict, instruction: str, lang: str = "",
    target_sheet_index: int | None = None,
) -> tuple[str, str]:
    """엑셀 수정용 프롬프트 빌드

    target_sheet_index가 지정되면 해당 시트만 전송하여 토큰 절약.
    """
    output_lang = lang or (settings.SUPPORTED_LANGS[0] if settings.SUPPORTED_LANGS else "ko")
    lang_instruction = {
        "ko": "모든 콘텐츠를 한국어로 작성하세요.",
        "en": "Write all content in English.",
        "ja": "すべてのコンテンツを日本語で作成してください。",
        "zh": "请用中文撰写所有内容。",
    }.get(output_lang, "모든 콘텐츠를 한국어로 작성하세요.")

    system_template = await get_prompt_content("excel_modify_system")
    user_template = await get_prompt_content("excel_modify_user")

    system_prompt = system_template.format(
        lang_instruction=lang_instruction,
    )

    sheets = current_data.get("sheets", [])

    # 시트 구조 변경 키워드 감지 (추가/삭제/이동 등)
    _sheet_struct_keywords = [
        "시트 추가", "시트를 추가", "시트 삭제", "시트를 삭제",
        "시트 제거", "시트를 제거", "시트 이동", "시트를 이동",
        "시트 복사", "시트를 복사", "새 시트", "새로운 시트",
        "마지막에 시트", "마지막 시트", "앞에 시트", "뒤에 시트",
        "sheet add", "add sheet", "new sheet", "remove sheet", "delete sheet",
        "시트를 만들", "시트 만들", "시트를 생성", "시트 생성",
        "시트를 넣", "시트 넣", "시트를 붙", "시트 붙",
    ]
    _is_sheet_structure_change = any(kw in instruction.lower() for kw in _sheet_struct_keywords)

    # 타겟 시트만 전송 (토큰 절약) — 단, 시트 구조 변경 요청이면 전체 전송
    if (target_sheet_index is not None
            and 0 <= target_sheet_index < len(sheets)
            and not _is_sheet_structure_change):
        target_sheet = sheets[target_sheet_index]
        # 타겟 시트 데이터만 포함, 다른 시트 이름만 참조로 제공
        other_sheet_names = [s.get("name", f"Sheet{i+1}") for i, s in enumerate(sheets) if i != target_sheet_index]
        send_data = {
            "meta": current_data.get("meta", {}),
            "target_sheet_index": target_sheet_index,
            "target_sheet": target_sheet,
        }
        if other_sheet_names:
            send_data["other_sheets_reference"] = other_sheet_names
        print(f"[LLM-Excel-Modify] 타겟 시트만 전송: index={target_sheet_index}, name={target_sheet.get('name', '?')}")
    else:
        if _is_sheet_structure_change:
            print(f"[LLM-Excel-Modify] 시트 구조 변경 감지 → 전체 시트 전송 ({len(sheets)}개)")
        send_data = current_data

    current_excel_json = json.dumps(send_data, ensure_ascii=False, indent=2)
    if len(current_excel_json) > 30000:
        current_excel_json = current_excel_json[:30000] + "\n... (데이터 일부 생략)"

    user_prompt = user_template.format(
        lang_instruction=lang_instruction,
        instruction=instruction,
        current_excel_data=current_excel_json,
    )

    return system_prompt, user_prompt


async def _build_excel_generation_prompts(
    resources_text: str, instructions: str, lang: str = "",
    sheet_count: str = "auto",
) -> tuple[str, str]:
    """엑셀 생성용 프롬프트 빌드"""
    if sheet_count and sheet_count != "auto":
        sheet_count_instruction = f"\n9. **[필수] 시트를 정확히 {sheet_count}개 생성하세요.**"
    else:
        sheet_count_instruction = ""

    output_lang = lang or (settings.SUPPORTED_LANGS[0] if settings.SUPPORTED_LANGS else "ko")
    lang_instruction = {
        "ko": "모든 콘텐츠를 한국어로 작성하세요.",
        "en": "Write all content in English.",
        "ja": "すべてのコンテンツを日本語で作成してください。",
        "zh": "请用中文撰写所有内容。",
    }.get(output_lang, "모든 콘텐츠를 한국어로 작성하세요.")

    system_template = await get_prompt_content("excel_generation_system")
    user_template = await get_prompt_content("excel_generation_user")

    system_prompt = system_template.format(
        lang_instruction=lang_instruction,
        sheet_count_instruction=sheet_count_instruction,
    )

    user_prompt = user_template.format(
        lang_instruction=lang_instruction,
        instructions=instructions if instructions else "리소스 내용을 분석하여 핵심 데이터를 표 형태로 정리해주세요.",
        resources_text=resources_text[:12000],
    )

    return system_prompt, user_prompt


def _parse_excel_schema(response_text: str) -> dict | None:
    """LLM 응답에서 엑셀 JSON 스키마를 파싱"""
    text = response_text.strip()

    # ```json ... ``` 블록 추출
    if "```json" in text:
        start = text.index("```json") + 7
        try:
            end = text.index("```", start)
            text = text[start:end].strip()
        except ValueError:
            text = text[start:].strip()
    elif "```" in text:
        start = text.index("```") + 3
        try:
            end = text.index("```", start)
            text = text[start:end].strip()
        except ValueError:
            text = text[start:].strip()

    parsed = None
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        for start_char, end_char in [("{", "}"), ("[", "]")]:
            if start_char in text:
                try:
                    s = text.index(start_char)
                    e = text.rindex(end_char) + 1
                    parsed = json.loads(text[s:e])
                    break
                except (json.JSONDecodeError, ValueError):
                    continue

    if parsed is None:
        return None

    if isinstance(parsed, dict):
        sheets = parsed.get("sheets", [])
        if isinstance(sheets, list) and sheets:
            for sheet in sheets:
                _coerce_sheet_types(sheet)
                _normalize_charts(sheet)
            return {
                "meta": parsed.get("meta", {}),
                "sheets": sheets,
            }

    return None


def _coerce_sheet_types(sheet: dict):
    """시트의 rows 내 값을 적절한 타입으로 변환"""
    rows = sheet.get("rows", [])
    for row_idx, row in enumerate(rows):
        if not isinstance(row, list):
            continue
        for col_idx, val in enumerate(row):
            if isinstance(val, str):
                try:
                    rows[row_idx][col_idx] = int(val)
                    continue
                except (ValueError, TypeError):
                    pass
                try:
                    rows[row_idx][col_idx] = float(val)
                    continue
                except (ValueError, TypeError):
                    pass


def _normalize_charts(sheet: dict):
    """차트 정의 검증 및 정규화"""
    charts = sheet.get("charts")
    if not charts or not isinstance(charts, list):
        sheet.pop("charts", None)
        return

    columns = sheet.get("columns", [])
    rows = sheet.get("rows", [])
    num_cols = len(columns)
    num_rows = len(rows)

    if num_cols == 0 or num_rows == 0:
        sheet.pop("charts", None)
        return

    valid_types = {"bar", "line", "pie", "area", "scatter", "doughnut", "radar"}
    normalized = []

    for chart in charts[:3]:  # 시트당 최대 3개
        if not isinstance(chart, dict):
            continue
        if chart.get("type") not in valid_types:
            continue

        dr = chart.get("data_range", {})
        if not isinstance(dr, dict) or not dr.get("series"):
            continue

        # labels_column 범위 검증
        labels_col = dr.get("labels_column", 0)
        if not isinstance(labels_col, int) or labels_col < 0 or labels_col >= num_cols:
            dr["labels_column"] = 0

        # series 열 인덱스 범위 검증
        valid_series = []
        for s in dr["series"]:
            if not isinstance(s, dict):
                continue
            col = s.get("column", -1)
            if isinstance(col, int) and 0 <= col < num_cols:
                valid_series.append(s)
        if not valid_series:
            continue
        dr["series"] = valid_series

        # pie/doughnut은 시리즈 1개만
        if chart["type"] in ("pie", "doughnut"):
            dr["series"] = dr["series"][:1]

        # row 범위 정규화
        dr.setdefault("row_start", 0)
        if dr["row_start"] < 0:
            dr["row_start"] = 0
        if dr.get("row_end") is not None:
            dr["row_end"] = min(dr["row_end"], num_rows - 1)

        # options 기본값
        chart.setdefault("options", {})
        normalized.append(chart)

    if normalized:
        sheet["charts"] = normalized
    else:
        sheet.pop("charts", None)


def _excel_fallback(resources_text: str) -> dict:
    """AI 호출 실패 시 폴백 엑셀 데이터"""
    lines = resources_text.strip().split("\n")[:20]
    rows = []
    for i, line in enumerate(lines):
        text = line.strip()
        if text:
            rows.append([i + 1, text])

    return {
        "meta": {"title": "데이터 정리", "description": "리소스 내용 정리"},
        "sheets": [
            {
                "name": "데이터",
                "columns": ["번호", "내용"],
                "rows": rows if rows else [[1, "데이터가 없습니다"]],
            }
        ],
    }


# ============ Word 문서 콘텐츠 생성 ============

async def generate_docx_content_stream(
    resources_text: str,
    instructions: str,
    lang: str = "",
    section_count: str = "auto",
    template_structure: list = None,
):
    """스트리밍 방식 Word 문서 콘텐츠 생성

    Yields:
        ("delta", str)   - Claude에서 스트리밍된 텍스트 청크
        ("result", dict) - 최종 파싱 결과: {"sections": [...], "meta": {...}}
    """
    if not settings.ANTHROPIC_API_KEY or settings.ANTHROPIC_API_KEY == "your_anthropic_api_key_here":
        yield ("result", _docx_fallback(resources_text))
        return

    system_prompt, user_prompt = await _build_docx_generation_prompts(
        resources_text, instructions, lang, section_count=section_count,
        template_structure=template_structure,
    )

    prompt_model = await get_prompt_model("docx_generation_system")
    effective_model = prompt_model or settings.ANTHROPIC_MODEL

    # 페이지 수에 따라 max_tokens 동적 조절
    max_tokens = 16384
    if section_count and section_count != "auto":
        try:
            pages = int(section_count)
            if pages >= 20:
                max_tokens = 32768
            elif pages >= 10:
                max_tokens = 24576
        except ValueError:
            pass

    try:
        full_text = ""
        async for delta in _stream_claude_api(system_prompt, user_prompt, model=effective_model, max_tokens=max_tokens):
            full_text += delta
            yield ("delta", delta)

        print(f"[LLM-Docx] 스트리밍 완료 - 전체 응답 길이: {len(full_text)} chars (model: {effective_model})")
        parsed = _parse_docx_schema(full_text)
        if parsed:
            section_count_result = len(parsed.get("sections", []))
            print(f"[LLM-Docx] 파싱 성공 - sections: {section_count_result}개")
            yield ("result", parsed)
        else:
            print(f"[LLM-Docx] 파싱 실패 → 폴백 사용")
            yield ("result", _docx_fallback(resources_text))
    except Exception as e:
        import traceback
        print(f"[LLM-Docx] Streaming 호출 실패: {e}")
        traceback.print_exc()
        yield ("result", _docx_fallback(resources_text))


async def _build_docx_generation_prompts(
    resources_text: str, instructions: str, lang: str = "",
    section_count: str = "auto",
    template_structure: list = None,
) -> tuple[str, str]:
    """Word 문서 생성용 프롬프트 빌드"""
    if section_count and section_count != "auto":
        section_count_instruction = f"\n9. **[필수] 약 {section_count}페이지 분량의 문서를 생성하세요. A4 기준 1페이지는 약 500~600자입니다. 총 분량이 약 {section_count}페이지가 되도록 섹션 수와 각 섹션의 내용 길이를 충분히 조절하세요.**"
    else:
        section_count_instruction = ""

    output_lang = lang or (settings.SUPPORTED_LANGS[0] if settings.SUPPORTED_LANGS else "ko")
    lang_instruction = {
        "ko": "모든 콘텐츠를 한국어로 작성하세요.",
        "en": "Write all content in English.",
        "ja": "すべてのコンテンツを日本語で作成してください。",
        "zh": "请用中文撰写所有内容。",
    }.get(output_lang, "모든 콘텐츠를 한국어로 작성하세요.")

    # 템플릿 구조가 있으면 시스템 프롬프트에 추가 지침 삽입
    template_instruction = ""
    if template_structure:
        is_table_based = any(sec.get("type") == "table_cell" for sec in template_structure)

        sections_desc = []
        for i, sec in enumerate(template_structure):
            title = sec.get("title", "")
            placeholder = sec.get("placeholder", "")
            level = sec.get("level", 1)
            if is_table_based:
                desc = f'  섹션 {i+1}: (level: {level}) - 안내: "{placeholder}"'
            else:
                desc = f'  섹션 {i+1}: "{title}" (level: {level})'
                if placeholder:
                    desc += f" - 안내: {placeholder}"
            sections_desc.append(desc)

        if is_table_based:
            template_instruction = f"""

## [최우선] 템플릿 양식 준수 규칙
사용자가 Word 템플릿 양식(테이블 기반)을 제공했습니다. **이 규칙은 다른 모든 규칙보다 우선합니다.**

### 필수 규칙:
1. 아래 각 섹션의 안내 텍스트(placeholder)를 참고하여, 리소스에서 해당 섹션에 맞는 내용을 추출하여 채워넣으세요.
2. 섹션 수를 **정확히 {len(template_structure)}개**로 유지하세요. 섹션을 추가하거나 삭제하지 마세요.
3. 각 섹션의 title은 안내 텍스트를 **그대로** 사용하세요 (매칭에 사용됩니다).
4. **meta.title과 meta.description은 빈 문자열("")로 두세요.**
5. 각 섹션의 content는 Markdown 형식으로 작성하되, 간결하고 읽기 쉽게 작성하세요.
6. 리소스에 해당 내용이 없으면 맥락에 맞게 합리적으로 작성하세요.

### 템플릿 섹션 구조:
{chr(10).join(sections_desc)}
"""
        else:
            template_instruction = f"""

## [최우선] 템플릿 양식 준수 규칙
사용자가 Word 템플릿 양식을 제공했습니다. **이 규칙은 다른 모든 규칙보다 우선합니다.**

### 필수 규칙:
1. 템플릿에 정의된 섹션 제목(title)을 **글자 하나 바꾸지 말고 정확히 그대로** 사용하세요.
2. 템플릿에 정의된 섹션 순서를 **절대** 변경하지 마세요.
3. 템플릿에 **없는 섹션을 임의로 추가하지 마세요.** sections 배열에는 아래 템플릿 섹션만 포함합니다.
4. **meta.title과 meta.description은 빈 문자열("")로 두세요.** 템플릿에 이미 제목이 포함되어 있습니다.
5. **level 0 섹션(문서 타이틀)의 content는 빈 문자열("")로 두세요.** 타이틀은 템플릿 서식이 그대로 사용됩니다.
6. level 1 이상 섹션의 안내 텍스트(placeholder)를 참고하여, 리소스 자료에서 해당 섹션에 맞는 내용을 추출하여 채워넣으세요.
7. 리소스 자료에서 해당 섹션에 맞는 내용이 없으면, 맥락에 맞게 합리적으로 작성하세요.

### 템플릿 섹션 구조:
{chr(10).join(sections_desc)}
"""

    system_template = await get_prompt_content("docx_generation_system")
    user_template = await get_prompt_content("docx_generation_user")

    system_prompt = system_template.format(
        lang_instruction=lang_instruction,
        section_count_instruction=section_count_instruction,
    )

    # 템플릿 지침을 시스템 프롬프트 끝에 추가
    if template_instruction:
        system_prompt += template_instruction

    user_prompt = user_template.format(
        lang_instruction=lang_instruction,
        instructions=instructions if instructions else "리소스 내용을 분석하여 체계적인 문서로 정리해주세요.",
        resources_text=resources_text[:12000],
    )

    # 템플릿이 있으면 사용자 프롬프트에도 리마인더 추가
    if template_structure:
        is_table_tpl = any(sec.get("type") == "table_cell" for sec in template_structure)
        section_titles = [sec.get("title", "") for sec in template_structure]
        if is_table_tpl:
            user_prompt += f"\n\n## [최우선] 템플릿 양식 준수\n위 리소스를 분석하여 다음 템플릿 섹션에 맞게 내용을 채워넣으세요.\n**정확히 {len(template_structure)}개 섹션을 생성하세요. 각 섹션의 title은 아래 안내 텍스트를 그대로 사용하세요.**\n**meta.title과 meta.description은 반드시 빈 문자열(\"\")로 두세요.**\n" + "\n".join(f'- "{t}"' for t in section_titles)
        else:
            user_prompt += f"\n\n## [최우선] 템플릿 양식 준수\n위 리소스를 분석하여 다음 템플릿 섹션에 맞게 내용을 채워넣으세요.\n**아래 섹션 제목을 글자 하나 바꾸지 말고 정확히 동일하게 사용하세요. 절대 변경/추가/삭제하지 마세요.**\n**meta.title과 meta.description은 반드시 빈 문자열(\"\")로 두세요.**\n**level 0 섹션의 content는 빈 문자열(\"\")로 두세요.**\n" + "\n".join(f'- "{t}"' for t in section_titles)

    return system_prompt, user_prompt


def _parse_docx_schema(response_text: str) -> dict | None:
    """LLM 응답에서 Word 문서 JSON 스키마를 파싱"""
    text = response_text.strip()

    # ```json ... ``` 블록 추출 (마지막 ``` 사용 - 내부 chart 블록 간섭 방지)
    if "```json" in text:
        start = text.index("```json") + 7
        try:
            end = text.rindex("```")
            if end > start:
                text = text[start:end].strip()
            else:
                text = text[start:].strip()
        except ValueError:
            text = text[start:].strip()
    elif "```" in text:
        start = text.index("```") + 3
        try:
            end = text.rindex("```")
            if end > start:
                text = text[start:end].strip()
            else:
                text = text[start:].strip()
        except ValueError:
            text = text[start:].strip()

    parsed = None
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        for start_char, end_char in [("{", "}"), ("[", "]")]:
            if start_char in text:
                try:
                    s = text.index(start_char)
                    e = text.rindex(end_char) + 1
                    parsed = json.loads(text[s:e])
                    break
                except (json.JSONDecodeError, ValueError):
                    continue

    if parsed is None:
        return None

    if isinstance(parsed, dict):
        sections = parsed.get("sections", [])
        if isinstance(sections, list) and sections:
            return {
                "meta": parsed.get("meta", {}),
                "sections": sections,
            }

    return None


def _docx_fallback(resources_text: str) -> dict:
    """AI 호출 실패 시 폴백 Word 데이터"""
    lines = resources_text.strip().split("\n")[:30]
    content = "\n".join(line.strip() for line in lines if line.strip())

    return {
        "meta": {"title": "문서 정리", "description": "리소스 내용 정리"},
        "sections": [
            {
                "title": "내용 정리",
                "level": 1,
                "content": content if content else "데이터가 없습니다.",
            }
        ],
    }


# ============ 텍스트 리라이트 (선택 영역 수정) ============

async def rewrite_text_stream(
    selected_text: str,
    instructions: str,
    lang: str = "",
    context_text: str = "",
):
    """선택된 텍스트를 AI로 리라이트 (스트리밍)

    Yields:
        ("delta", str) - 스트리밍 텍스트 청크
        ("result", str) - 최종 완성 텍스트
    """
    if not settings.ANTHROPIC_API_KEY or settings.ANTHROPIC_API_KEY == "your_anthropic_api_key_here":
        yield ("result", selected_text)
        return

    output_lang = lang or (settings.SUPPORTED_LANGS[0] if settings.SUPPORTED_LANGS else "ko")
    lang_instruction = {
        "ko": "모든 콘텐츠를 한국어로 작성하세요.",
        "en": "Write all content in English.",
        "ja": "すべてのコンテンツを日本語で作成してください。",
        "zh": "请用中文撰写所有内容。",
    }.get(output_lang, "모든 콘텐츠를 한국어로 작성하세요.")

    system_template = await get_prompt_content("rewrite_system")
    user_template = await get_prompt_content("rewrite_user")

    system_prompt = system_template.format(lang_instruction=lang_instruction)
    user_prompt = user_template.format(
        lang_instruction=lang_instruction,
        instructions=instructions,
        selected_text=selected_text[:8000],
        context_text=context_text[:12000] if context_text else "(주변 문맥 없음)",
    )

    prompt_model = await get_prompt_model("rewrite_system")
    effective_model = prompt_model or settings.ANTHROPIC_MODEL

    try:
        full_text = ""
        async for delta in _stream_claude_api(system_prompt, user_prompt, model=effective_model, max_tokens=8192):
            full_text += delta
            yield ("delta", delta)

        print(f"[LLM-Rewrite] 완료 - 응답 길이: {len(full_text)} chars (model: {effective_model})")
        yield ("result", full_text)
    except Exception as e:
        import traceback
        print(f"[LLM-Rewrite] 호출 실패: {e}")
        traceback.print_exc()
        yield ("result", selected_text)


# ============ HTML 리포트 콘텐츠 생성 ============


async def _build_html_report_prompts(
    resources_text: str, instructions: str, skill: dict,
    lang: str = "", page_count: str = "auto",
) -> tuple[str, str]:
    """HTML 리포트 생성용 프롬프트 빌드"""
    if page_count and page_count != "auto":
        page_count_instruction = f"\n9. **[필수] 페이지를 정확히 {page_count}개 생성하세요.**"
    else:
        page_count_instruction = ""

    output_lang = lang or (settings.SUPPORTED_LANGS[0] if settings.SUPPORTED_LANGS else "ko")
    lang_instruction = {
        "ko": "모든 콘텐츠를 한국어로 작성하세요.",
        "en": "Write all content in English.",
        "ja": "すべてのコンテンツを日本語で作成してください。",
        "zh": "请用中文撰写所有内容。",
    }.get(output_lang, "모든 콘텐츠를 한국어로 작성하세요.")

    skill_prompt = skill.get("skill_prompt", "") or ""
    theme = skill.get("theme", "corporate") or "corporate"

    system_template = await get_prompt_content("html_report_generation_system")
    user_template = await get_prompt_content("html_report_generation_user")

    system_prompt = system_template.format(
        lang_instruction=lang_instruction,
        page_count_instruction=page_count_instruction,
        skill_prompt=skill_prompt if skill_prompt else "특별한 스킬 지침 없음",
        theme=theme,
    )

    user_prompt = user_template.format(
        lang_instruction=lang_instruction,
        instructions=instructions if instructions else "리소스 내용을 분석하여 전문적인 리포트로 정리해주세요.",
        resources_text=resources_text[:12000],
        skill_prompt=skill_prompt if skill_prompt else "특별한 스킬 지침 없음",
    )

    return system_prompt, user_prompt


def _parse_html_report_schema(response_text: str) -> dict | None:
    """LLM 응답에서 HTML 리포트 JSON 스키마를 파싱"""
    text = response_text.strip()

    # ```json ... ``` 블록 추출
    if "```json" in text:
        start = text.index("```json") + 7
        try:
            end = text.rindex("```")
            if end > start:
                text = text[start:end].strip()
            else:
                text = text[start:].strip()
        except ValueError:
            text = text[start:].strip()
    elif "```" in text:
        start = text.index("```") + 3
        try:
            end = text.rindex("```")
            if end > start:
                text = text[start:end].strip()
            else:
                text = text[start:].strip()
        except ValueError:
            text = text[start:].strip()

    parsed = None
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        for start_char, end_char in [("{", "}"), ("[", "]")]:
            if start_char in text:
                try:
                    s = text.index(start_char)
                    e = text.rindex(end_char) + 1
                    parsed = json.loads(text[s:e])
                    break
                except (json.JSONDecodeError, ValueError):
                    continue

    if parsed is None:
        # 잘린 JSON 복구 시도
        repaired = _try_repair_truncated_json(text)
        if repaired is not None and "pages" in repaired:
            parsed = repaired

    if parsed is None:
        return None

    if isinstance(parsed, dict):
        pages = parsed.get("pages", [])
        if isinstance(pages, list) and pages:
            return {
                "meta": parsed.get("meta", {}),
                "pages": pages,
            }

    return None


async def generate_html_report_stream(
    resources_text: str,
    instructions: str,
    skill: dict,
    lang: str = "",
    page_count: str = "auto",
):
    """[Legacy] 스트리밍 방식 HTML 리포트 콘텐츠 생성 (전체 한 번에 생성)

    2-phase 방식의 generate_html_report_outline_stream() + generate_html_page_content()로 대체됨.

    Yields:
        ("delta", str)   - Claude에서 스트리밍된 텍스트 청크
        ("result", dict) - 최종 파싱 결과: {"pages": [...], "meta": {...}}
    """
    if not settings.ANTHROPIC_API_KEY or settings.ANTHROPIC_API_KEY == "your_anthropic_api_key_here":
        yield ("result", _html_report_fallback(resources_text))
        return

    system_prompt, user_prompt = await _build_html_report_prompts(
        resources_text, instructions, skill, lang, page_count=page_count,
    )

    prompt_model = await get_prompt_model("html_report_generation_system")
    effective_model = prompt_model or settings.ANTHROPIC_MODEL

    # 페이지 수에 따라 max_tokens 동적 조절
    max_tokens = 16384
    if page_count and page_count != "auto":
        try:
            pages_num = int(page_count)
            if pages_num >= 15:
                max_tokens = 32768
            elif pages_num >= 8:
                max_tokens = 24576
        except ValueError:
            pass

    try:
        full_text = ""
        async for delta in _stream_claude_api(system_prompt, user_prompt, model=effective_model, max_tokens=max_tokens):
            full_text += delta
            yield ("delta", delta)

        print(f"[LLM-HTML] 스트리밍 완료 - 전체 응답 길이: {len(full_text)} chars (model: {effective_model})")
        parsed = _parse_html_report_schema(full_text)
        if parsed:
            page_count_result = len(parsed.get("pages", []))
            print(f"[LLM-HTML] 파싱 성공 - pages: {page_count_result}개")
            yield ("result", parsed)
        else:
            print(f"[LLM-HTML] 파싱 실패 → 폴백 사용")
            yield ("result", _html_report_fallback(resources_text))
    except Exception as e:
        import traceback
        print(f"[LLM-HTML] Streaming 호출 실패: {e}")
        traceback.print_exc()
        yield ("result", _html_report_fallback(resources_text))


def _html_report_fallback(resources_text: str) -> dict:
    """AI 호출 실패 시 폴백 HTML 리포트 데이터"""
    content = resources_text[:500] if resources_text else "데이터가 없습니다."
    return {
        "meta": {"title": "리포트", "description": "리소스 내용 정리"},
        "pages": [
            {
                "order": 1,
                "title": "내용 정리",
                "html_content": f'<div style="width:960px;height:540px;padding:40px;font-family:Arial,sans-serif;"><h1 style="color:#333;">리포트</h1><p>{content}</p></div>',
            }
        ],
    }


# ============ 2-Phase HTML 리포트 생성 ============


async def _build_html_report_outline_prompts(
    resources_text: str, instructions: str, skill_prompt: str,
    lang: str = "", page_count: str = "auto",
) -> tuple[str, str]:
    """HTML 리포트 아웃라인 생성용 프롬프트 빌드"""
    if page_count and page_count != "auto":
        page_count_instruction = f"\n9. **[필수] 반드시 pages 배열에 정확히 {page_count}개의 페이지 객체를 생성하세요. 절대 1개로 합치지 마세요.**"
    else:
        page_count_instruction = "\n9. **[필수] 반드시 pages 배열에 최소 3개 이상의 페이지 객체를 생성하세요. 절대 1개로 합치지 마세요. 내용이 풍부하면 5~8페이지로 분할합니다.**"

    output_lang = lang or (settings.SUPPORTED_LANGS[0] if settings.SUPPORTED_LANGS else "ko")
    lang_instruction = {
        "ko": "모든 콘텐츠를 한국어로 작성하세요.",
        "en": "Write all content in English.",
        "ja": "すべてのコンテンツを日本語で作成してください。",
        "zh": "请用中文撰写所有内容。",
    }.get(output_lang, "모든 콘텐츠를 한국어로 작성하세요.")

    system_template = await get_prompt_content("html_report_outline_system")
    user_template = await get_prompt_content("html_report_outline_user")

    system_prompt = system_template.format(
        lang_instruction=lang_instruction,
        page_count_instruction=page_count_instruction,
        skill_prompt=skill_prompt if skill_prompt else "특별한 스킬 지침 없음",
    )

    user_prompt = user_template.format(
        lang_instruction=lang_instruction,
        instructions=instructions if instructions else "리소스 내용을 분석하여 전문적인 리포트 아웃라인을 설계해주세요.",
        resources_text=resources_text[:12000],
        skill_prompt=skill_prompt if skill_prompt else "특별한 스킬 지침 없음",
    )

    return system_prompt, user_prompt


async def _build_html_css_prompts(
    outline_data: dict, skill_prompt: str,
    lang: str = "", theme: str = "corporate",
) -> tuple[str, str]:
    """HTML 리포트 공통 CSS 생성용 프롬프트 빌드"""
    output_lang = lang or (settings.SUPPORTED_LANGS[0] if settings.SUPPORTED_LANGS else "ko")
    lang_instruction = {
        "ko": "모든 콘텐츠를 한국어로 작성하세요.",
        "en": "Write all content in English.",
        "ja": "すべてのコンテンツを日本語で作成してください。",
        "zh": "请用中文撰写所有内容。",
    }.get(output_lang, "모든 콘텐츠를 한국어로 작성하세요.")

    outline_json = json.dumps(outline_data, ensure_ascii=False, indent=2)

    system_template = await get_prompt_content("html_report_css_system")
    user_template = await get_prompt_content("html_report_css_user")

    system_prompt = system_template.format(
        lang_instruction=lang_instruction,
        skill_prompt=skill_prompt if skill_prompt else "특별한 스킬 지침 없음",
        theme=theme,
    )

    user_prompt = user_template.format(
        outline_json=outline_json,
        lang_instruction=lang_instruction,
        skill_prompt=skill_prompt if skill_prompt else "특별한 스킬 지침 없음",
        theme=theme,
    )

    return system_prompt, user_prompt


async def generate_html_report_css_stream(
    outline_data: dict,
    skill_prompt: str,
    lang: str = "",
    theme: str = "corporate",
):
    """Phase 2: 공통 CSS 생성 (스트리밍)

    아웃라인을 기반으로 전체 리포트에서 사용할 공통 CSS를 생성합니다.
    빠른 모델(ANTHROPIC_OUTLINE_MODEL)을 사용합니다.

    Yields:
        ("delta", str)   - CSS 스트리밍 청크
        ("result", str)  - 최종 CSS 문자열
    """
    if not settings.ANTHROPIC_API_KEY or settings.ANTHROPIC_API_KEY == "your_anthropic_api_key_here":
        yield ("result", ".rpt-page { width:100%; padding:40px; font-family:'Malgun Gothic',Arial,sans-serif; box-sizing:border-box; }")
        return

    system_prompt, user_prompt = await _build_html_css_prompts(
        outline_data, skill_prompt, lang, theme,
    )

    prompt_model = await get_prompt_model("html_report_css_system")
    effective_model = prompt_model or settings.ANTHROPIC_OUTLINE_MODEL

    try:
        full_text = ""
        async for delta in _stream_claude_api(system_prompt, user_prompt, model=effective_model, max_tokens=4096):
            full_text += delta
            yield ("delta", delta)

        # CSS 코드 블록 래퍼 제거
        css_text = full_text.strip()
        if css_text.startswith("```css"):
            css_text = css_text[6:]
        elif css_text.startswith("```"):
            css_text = css_text[3:]
        if css_text.endswith("```"):
            css_text = css_text[:-3]
        css_text = css_text.strip()

        print(f"[LLM-HTML-CSS] CSS 생성 완료 - 길이: {len(css_text)} chars (model: {effective_model})")
        yield ("result", css_text)
    except Exception as e:
        import traceback
        print(f"[LLM-HTML-CSS] CSS 생성 실패: {e}")
        traceback.print_exc()
        yield ("result", ".rpt-page { width:100%; padding:40px; font-family:'Malgun Gothic',Arial,sans-serif; box-sizing:border-box; }")


async def _build_html_page_prompts(
    resources_text: str, instructions: str, skill_prompt: str,
    page_info: dict, all_pages_outline: dict,
    lang: str = "", theme: str = "corporate", common_css: str = "",
) -> tuple[str, str]:
    """HTML 리포트 단일 페이지 생성용 프롬프트 빌드"""
    output_lang = lang or (settings.SUPPORTED_LANGS[0] if settings.SUPPORTED_LANGS else "ko")
    lang_instruction = {
        "ko": "모든 콘텐츠를 한국어로 작성하세요.",
        "en": "Write all content in English.",
        "ja": "すべてのコンテンツを日本語で作成してください。",
        "zh": "请用中文撰写所有内容。",
    }.get(output_lang, "모든 콘텐츠를 한국어로 작성하세요.")

    all_pages = all_pages_outline.get("pages", [])
    meta = all_pages_outline.get("meta", {})
    page_order = page_info.get("order", 1)
    total_pages = len(all_pages)

    # 이전 페이지 제목 목록 (현재 페이지 전까지)
    previous_titles = []
    for p in all_pages:
        if p.get("order", 0) < page_order:
            previous_titles.append(p.get("title", ""))
    previous_page_titles = ", ".join(previous_titles) if previous_titles else "없음 (첫 번째 페이지)"

    # key_points를 문자열로 변환
    key_points = page_info.get("key_points", [])
    key_points_str = "\n".join(f"  - {kp}" for kp in key_points) if key_points else "없음"

    # 리소스 텍스트: 페이지 순서에 따라 다른 구간의 리소스를 전달
    if total_pages > 1:
        chunk_size = max(4000, len(resources_text) // total_pages)
        page_idx = page_order - 1
        start = page_idx * (len(resources_text) // total_pages)
        resources_excerpt = resources_text[start:start + chunk_size]
        if not resources_excerpt.strip():
            resources_excerpt = resources_text[:chunk_size]
    else:
        resources_excerpt = resources_text[:8000]

    # 공통 CSS에서 클래스 목록 추출
    css_classes_list = re.findall(r'\.(rpt-[\w-]+)', common_css) if common_css else []
    css_classes_unique = list(dict.fromkeys(css_classes_list))  # 중복 제거 (순서 유지)
    css_classes_str = ", ".join(f".{c}" for c in css_classes_unique) if css_classes_unique else "없음 (인라인 스타일 사용)"

    system_template = await get_prompt_content("html_page_generation_system")
    user_template = await get_prompt_content("html_page_generation_user")

    system_prompt = system_template.format(
        lang_instruction=lang_instruction,
        skill_prompt=skill_prompt if skill_prompt else "특별한 스킬 지침 없음",
        theme=theme,
    )

    user_prompt = user_template.format(
        lang_instruction=lang_instruction,
        page_title=page_info.get("title", ""),
        page_summary=page_info.get("summary", ""),
        key_points=key_points_str,
        report_title=meta.get("title", "리포트"),
        page_order=page_order,
        total_pages=total_pages,
        previous_page_titles=previous_page_titles,
        resources_text=resources_excerpt,
        instructions=instructions if instructions else "리소스 내용을 바탕으로 전문적인 페이지를 생성해주세요.",
        css_classes=css_classes_str,
    )

    return system_prompt, user_prompt


def _parse_html_report_outline(response_text: str) -> dict | None:
    """LLM 응답에서 HTML 리포트 아웃라인 JSON을 파싱"""
    text = response_text.strip()

    # ```json ... ``` 블록 추출
    if "```json" in text:
        start = text.index("```json") + 7
        try:
            end = text.rindex("```")
            if end > start:
                text = text[start:end].strip()
            else:
                text = text[start:].strip()
        except ValueError:
            text = text[start:].strip()
    elif "```" in text:
        start = text.index("```") + 3
        try:
            end = text.rindex("```")
            if end > start:
                text = text[start:end].strip()
            else:
                text = text[start:].strip()
        except ValueError:
            text = text[start:].strip()

    parsed = None
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        # 전략 1: { ... } 또는 [ ... ] 경계 찾기
        for start_char, end_char in [("{", "}"), ("[", "]")]:
            if start_char in text:
                try:
                    s = text.index(start_char)
                    e = text.rindex(end_char) + 1
                    parsed = json.loads(text[s:e])
                    break
                except (json.JSONDecodeError, ValueError):
                    continue

        # 전략 2: "pages" 키를 찾아 배열 직접 추출
        if parsed is None and '"pages"' in text:
            try:
                # "pages": [...] 패턴 매칭 (중첩 대괄호 처리)
                pages_idx = text.index('"pages"')
                bracket_start = text.index('[', pages_idx)
                depth = 0
                end_idx = bracket_start
                for i in range(bracket_start, len(text)):
                    if text[i] == '[':
                        depth += 1
                    elif text[i] == ']':
                        depth -= 1
                        if depth == 0:
                            end_idx = i + 1
                            break
                pages_json = text[bracket_start:end_idx]
                pages_list = json.loads(pages_json)
                if isinstance(pages_list, list) and pages_list:
                    # meta도 추출 시도
                    meta = {}
                    if '"meta"' in text:
                        try:
                            meta_idx = text.index('"meta"')
                            meta_brace = text.index('{', meta_idx)
                            d2 = 0
                            for j in range(meta_brace, len(text)):
                                if text[j] == '{': d2 += 1
                                elif text[j] == '}':
                                    d2 -= 1
                                    if d2 == 0:
                                        meta = json.loads(text[meta_brace:j+1])
                                        break
                        except Exception:
                            pass
                    parsed = {"meta": meta, "pages": pages_list}
            except Exception as e:
                print(f"[LLM-HTML-Outline] pages 직접 추출 실패: {e}")

    if parsed is None:
        print(f"[LLM-HTML-Outline] 파싱 완전 실패 - 원본 텍스트 앞 500자: {text[:500]}")
        return None

    if isinstance(parsed, dict):
        pages = parsed.get("pages", [])
        if isinstance(pages, list) and pages:
            return {
                "meta": parsed.get("meta", {}),
                "pages": pages,
            }

    # parsed가 list인 경우 (pages 배열만 반환한 경우)
    if isinstance(parsed, list) and parsed:
        return {
            "meta": {},
            "pages": parsed,
        }

    print(f"[LLM-HTML-Outline] 파싱 결과에 pages 없음 - parsed type: {type(parsed)}")
    return None


async def generate_html_report_outline_stream(
    resources_text: str,
    instructions: str,
    skill_prompt: str,
    lang: str = "",
    page_count: str = "auto",
):
    """Phase 1: 스트리밍 방식 HTML 리포트 아웃라인 생성

    아웃라인(페이지 제목 + 요약 + 핵심 포인트)만 생성합니다.
    빠른 모델(ANTHROPIC_OUTLINE_MODEL)을 사용하여 신속하게 구조를 설계합니다.

    Yields:
        ("delta", str)   - Claude에서 스트리밍된 텍스트 청크
        ("result", dict) - 최종 파싱 결과: {"meta": {...}, "pages": [...]}
    """
    if not settings.ANTHROPIC_API_KEY or settings.ANTHROPIC_API_KEY == "your_anthropic_api_key_here":
        yield ("result", {
            "meta": {"title": "리포트", "description": ""},
            "pages": [
                {"order": 1, "title": "내용 정리", "summary": "리소스 내용을 정리합니다.", "key_points": ["핵심 내용"]}
            ],
        })
        return

    system_prompt, user_prompt = await _build_html_report_outline_prompts(
        resources_text, instructions, skill_prompt, lang, page_count=page_count,
    )

    # 아웃라인용 모델 (빠른 모델 사용)
    prompt_model = await get_prompt_model("html_report_outline_system")
    effective_model = prompt_model or settings.ANTHROPIC_OUTLINE_MODEL

    try:
        full_text = ""
        async for delta in _stream_claude_api(system_prompt, user_prompt, model=effective_model, max_tokens=4096):
            full_text += delta
            yield ("delta", delta)

        print(f"[LLM-HTML-Outline] 스트리밍 완료 - 전체 응답 길이: {len(full_text)} chars (model: {effective_model})")
        parsed = _parse_html_report_outline(full_text)
        if parsed:
            page_count_result = len(parsed.get("pages", []))
            print(f"[LLM-HTML-Outline] 파싱 성공 - pages: {page_count_result}개")

            # 자동 모드에서 1페이지만 생성된 경우 → 재시도 (명시적 페이지 수 지정)
            if page_count_result <= 1 and (not page_count or page_count == "auto"):
                print(f"[LLM-HTML-Outline] 1페이지만 생성됨 → 3페이지로 재시도")
                yield ("retry", "재시도 중...")
                retry_system, retry_user = await _build_html_report_outline_prompts(
                    resources_text, instructions, skill_prompt, lang, page_count="3",
                )
                retry_text = ""
                async for delta2 in _stream_claude_api(retry_system, retry_user, model=effective_model, max_tokens=4096):
                    retry_text += delta2
                    yield ("retry_delta", delta2)
                retry_parsed = _parse_html_report_outline(retry_text)
                if retry_parsed and len(retry_parsed.get("pages", [])) > 1:
                    print(f"[LLM-HTML-Outline] 재시도 성공 - pages: {len(retry_parsed['pages'])}개")
                    parsed = retry_parsed
                else:
                    print(f"[LLM-HTML-Outline] 재시도 실패 - 원본 유지 ({page_count_result}페이지)")

            yield ("result", parsed)
        else:
            print(f"[LLM-HTML-Outline] 파싱 실패 → 재시도")
            print(f"[LLM-HTML-Outline] 원본 응답 앞 500자: {full_text[:500]}")
            # 파싱 실패 시 재시도 (명시적 페이지 수로)
            yield ("retry", "아웃라인 재구성 중...")
            retry_system, retry_user = await _build_html_report_outline_prompts(
                resources_text, instructions, skill_prompt, lang, page_count="3",
            )
            retry_text = ""
            async for delta2 in _stream_claude_api(retry_system, retry_user, model=effective_model, max_tokens=4096):
                retry_text += delta2
                yield ("retry_delta", delta2)
            retry_parsed = _parse_html_report_outline(retry_text)
            if retry_parsed and retry_parsed.get("pages"):
                print(f"[LLM-HTML-Outline] 재시도 성공 - pages: {len(retry_parsed['pages'])}개")
                yield ("result", retry_parsed)
            else:
                print(f"[LLM-HTML-Outline] 재시도도 실패 → 폴백 사용")
                if retry_text:
                    print(f"[LLM-HTML-Outline] 재시도 응답 앞 500자: {retry_text[:500]}")
                yield ("result", {
                    "meta": {"title": "리포트", "description": ""},
                    "pages": [
                        {"order": 1, "title": "내용 정리", "summary": "리소스 내용을 정리합니다.", "key_points": ["핵심 내용"]}
                    ],
                })
    except Exception as e:
        import traceback
        print(f"[LLM-HTML-Outline] Streaming 호출 실패: {e}")
        traceback.print_exc()
        yield ("result", {
            "meta": {"title": "리포트", "description": ""},
            "pages": [
                {"order": 1, "title": "내용 정리", "summary": "리소스 내용을 정리합니다.", "key_points": ["핵심 내용"]}
            ],
        })


async def generate_html_page_content(
    resources_text: str,
    instructions: str,
    skill_prompt: str,
    page_info: dict,
    all_pages_outline: dict,
    lang: str = "",
    theme: str = "corporate",
    common_css: str = "",
):
    """Phase 2: 단일 HTML 리포트 페이지 콘텐츠 생성 (스트리밍)

    아웃라인의 각 페이지에 대해 개별적으로 호출하여 HTML 콘텐츠를 생성합니다.
    스트리밍 방식으로 실시간 진행을 표시합니다.

    Yields:
        ("delta", str)   - Claude에서 스트리밍된 텍스트 청크
        ("result", str)  - 최종 HTML 문자열
    """
    if not settings.ANTHROPIC_API_KEY or settings.ANTHROPIC_API_KEY == "your_anthropic_api_key_here":
        fallback_html = f'<div style="width:960px;height:540px;padding:40px;font-family:Arial,sans-serif;"><h1>{page_info.get("title", "페이지")}</h1><p>{page_info.get("summary", "")}</p></div>'
        yield ("result", fallback_html)
        return

    system_prompt, user_prompt = await _build_html_page_prompts(
        resources_text, instructions, skill_prompt,
        page_info, all_pages_outline, lang, theme, common_css=common_css,
    )

    # 페이지 생성용 모델 (고품질 모델 사용)
    prompt_model = await get_prompt_model("html_page_generation_system")
    effective_model = prompt_model or settings.ANTHROPIC_MODEL

    try:
        full_text = ""
        async for delta in _stream_claude_api(system_prompt, user_prompt, model=effective_model, max_tokens=8192):
            full_text += delta
            yield ("delta", delta)

        page_order = page_info.get("order", "?")
        print(f"[LLM-HTML-Page] 페이지 {page_order} 스트리밍 완료 - 길이: {len(full_text)} chars (model: {effective_model})")
        yield ("result", full_text)
    except Exception as e:
        import traceback
        print(f"[LLM-HTML-Page] 페이지 {page_info.get('order', '?')} 생성 실패: {e}")
        traceback.print_exc()
        fallback_html = f'<div style="width:960px;height:540px;padding:40px;font-family:Arial,sans-serif;"><h1>{page_info.get("title", "페이지")}</h1><p>{page_info.get("summary", "")}</p></div>'
        yield ("result", fallback_html)
