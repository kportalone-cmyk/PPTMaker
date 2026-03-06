"""
Claude Opus LLM 서비스 - 리치 스키마 기반 슬라이드 콘텐츠 생성

리소스 내용을 분석하여 구조화된 프레젠테이션 스키마를 생성하고,
관리자가 설정한 템플릿 오브젝트(제목/거버넌스/부제목/내용 등)에 매핑합니다.
"""

import json
import re
import random
import httpx
from config import settings
from routers.prompt import get_prompt_content


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

    try:
        result = await _call_claude_api(system_prompt, user_prompt, model=settings.ANTHROPIC_OUTLINE_MODEL)
        print(f"[LLM] API 응답 길이: {len(result)} chars (model: {settings.ANTHROPIC_OUTLINE_MODEL})")
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
    # 모델별 max_tokens 적용
    effective_model = model or settings.ANTHROPIC_MODEL
    if effective_model == settings.ANTHROPIC_OUTLINE_MODEL and settings.ANTHROPIC_OUTLINE_MAX_TOKENS > 0:
        payload["max_tokens"] = settings.ANTHROPIC_OUTLINE_MAX_TOKENS
    elif settings.ANTHROPIC_MAX_TOKENS > 0:
        payload["max_tokens"] = settings.ANTHROPIC_MAX_TOKENS

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
            # 슬라이드에 표시되는 placeholder 수와 관계없이 아웃라인용으로 최소 3개 항목 요청
            min_items = max(effective_items, 3)
            max_items = max(effective_items, 4)
            lines.append(f"  ※ 본문 슬라이드: items를 최소 {min_items}개~최대 {max_items}개 생성하세요. 슬라이드에는 {effective_items}개까지 표시되며 나머지는 아웃라인에 표시됩니다. 절대 1개만 생성하지 마세요.")
        elif content_type == "section_divider":
            # 간지 슬라이드: 부제목 필드 유무 안내
            has_subtitle_ph = ph_sub_actual > 0
            if has_subtitle_ph:
                lines.append(f"  ※ 간지 슬라이드: section_title(제목)은 필수입니다. 부제목 있음 → section_subtitle도 생성하세요.")
            else:
                lines.append(f"  ※ 간지 슬라이드: section_title(제목)은 필수입니다. 부제목 없음 → section_subtitle은 생성하지 마세요.")
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
    for sm in slides_meta:
        meta = sm.get("slide_meta", {})
        if meta.get("content_type") == "body":
            phs = sm.get("placeholders", [])
            ph_sub = sum(1 for ph in phs if ph.get("role") == "subtitle")
            ph_desc = sum(1 for ph in phs if ph.get("role") == "description")
            effective = min(ph_sub, ph_desc) if (ph_sub > 0 and ph_desc > 0) else max(ph_sub, ph_desc)
            _body_groups.setdefault(effective, []).append(sm["slide_index"])

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

            # content 슬라이드: subtitle/description 수가 items보다 적으면 재매칭
            if slide_type == "content" and items_count > 0:
                chosen_meta = None
                for sm in slides_meta:
                    if sm["slide_index"] == template_idx:
                        chosen_meta = sm
                        break
                if chosen_meta:
                    phs = chosen_meta.get("placeholders", [])
                    ph_desc_count = sum(1 for ph in phs if ph.get("role") == "description")
                    ph_sub_count = sum(1 for ph in phs if ph.get("role") == "subtitle")
                    # subtitle과 description 모두 items 수를 수용해야 함
                    effective_slots = min(ph_sub_count, ph_desc_count) if (ph_sub_count > 0 and ph_desc_count > 0) else max(ph_sub_count, ph_desc_count)
                    if effective_slots < items_count:
                        better_idx = _find_best_template_for_type(
                            slide_type, slides_meta, items_count
                        )
                        if better_idx != template_idx:
                            template_idx = better_idx

        # template_index가 없거나 유효하지 않으면 자동 매칭
        if template_idx is None:
            template_idx = _find_best_template_for_type(slide_type, slides_meta, items_count)

        # 동일 용량 본문 템플릿이 여러 개일 때 라운드로빈으로 돌아가며 사용
        if slide_type == "content":
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

        # content 슬라이드: 선택된 템플릿의 슬롯 수에 맞게 items 보충
        if slide_type == "content":
            phs = target_meta.get("placeholders", [])
            ph_sub = sum(1 for ph in phs if ph.get("role") == "subtitle")
            ph_desc = sum(1 for ph in phs if ph.get("role") == "description")
            effective_slots = min(ph_sub, ph_desc) if (ph_sub > 0 and ph_desc > 0) else max(ph_sub, ph_desc)
            items = slide.get("items", [])
            # 슬롯보다 items가 적을 때만 보충 시도 (최소 슬롯 수 충족 목표)
            if 0 < len(items) < effective_slots:
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

        # number placeholder에 순서 번호 자동 매핑
        for i, item in enumerate(items):
            if i < len(number_placeholders):
                contents[number_placeholders[i]] = str(i + 1)

        if subtitle_placeholders and desc_placeholders and items:
            # subtitle에 번호, description에 항목 텍스트 분리 배치
            for i, item in enumerate(items):
                num = item.get("num", str(i + 1))
                text = item.get("text", "")
                if i < len(subtitle_placeholders):
                    contents[subtitle_placeholders[i]] = num
                if i < len(desc_placeholders):
                    contents[desc_placeholders[i]] = text
        elif subtitle_placeholders and items:
            # subtitle placeholder에 "num. text" 배치
            for i, item in enumerate(items):
                if i >= len(subtitle_placeholders):
                    break
                num = item.get("num", "")
                text = item.get("text", "")
                contents[subtitle_placeholders[i]] = f"{num}. {text}" if num else text
        elif desc_placeholders and items:
            # description placeholder에 "num. text" 배치
            for i, item in enumerate(items):
                if i >= len(desc_placeholders):
                    break
                num = item.get("num", "")
                text = item.get("text", "")
                contents[desc_placeholders[i]] = f"{num}. {text}" if num else text
        elif items:
            toc_lines = []
            for item in items:
                num = item.get("num", "")
                text = item.get("text", "")
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

        # number placeholder에 순서 번호 자동 매핑 (1, 2, 3, ...)
        for i in range(len(items)):
            if i < len(number_placeholders):
                contents[number_placeholders[i]] = str(i + 1)

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
                score += 8  # 정확히 일치
            elif effective_slots >= items_count:
                score += 5  # placeholder가 더 많음 (수용 가능)
            elif effective_slots > 0 and effective_slots < items_count:
                score -= 5  # placeholder 부족 (강한 패널티)

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
            contents[name] = "PPTMaker 자동 생성"
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


async def _stream_claude_api(system_prompt: str, user_prompt: str, model: str = ""):
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
    # 모델별 max_tokens 적용
    effective_model = model or settings.ANTHROPIC_MODEL
    if effective_model == settings.ANTHROPIC_OUTLINE_MODEL and settings.ANTHROPIC_OUTLINE_MAX_TOKENS > 0:
        payload["max_tokens"] = settings.ANTHROPIC_OUTLINE_MAX_TOKENS
    elif settings.ANTHROPIC_MAX_TOKENS > 0:
        payload["max_tokens"] = settings.ANTHROPIC_MAX_TOKENS

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

    try:
        full_text = ""
        async for delta in _stream_claude_api(system_prompt, user_prompt, model=settings.ANTHROPIC_OUTLINE_MODEL):
            full_text += delta
            yield ("delta", delta)

        print(f"[LLM] 스트리밍 완료 - 전체 응답 길이: {len(full_text)} chars (model: {settings.ANTHROPIC_OUTLINE_MODEL})")
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
