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

    # 슬라이드 카탈로그 설명 생성
    slides_description = _build_slides_description(slides_meta)

    # 슬라이드 수 지침
    if slide_count and slide_count != "auto":
        slide_count_instruction = f"\n11. **[필수] 전체 슬라이드 수를 정확히 {slide_count}장으로 생성하세요.** title, toc, section, content, closing 모두 포함하여 총 {slide_count}장이어야 합니다."
    else:
        slide_count_instruction = ""

    # 언어별 프롬프트 분기
    output_lang = lang or (settings.SUPPORTED_LANGS[0] if settings.SUPPORTED_LANGS else "ko")
    lang_instruction = {
        "ko": "모든 콘텐츠를 한국어로 작성하세요.",
        "en": "Write all content in English.",
        "ja": "すべてのコンテンツを日本語で作成してください。",
        "zh": "请用中文撰写所有内容。",
    }.get(output_lang, "모든 콘텐츠를 한국어로 작성하세요.")

    system_prompt = f"""당신은 기업용 프레젠테이션 구조 설계 및 콘텐츠 전문가입니다.
주어진 리소스 자료를 분석하여 전문적인 프레젠테이션을 설계합니다.

## 슬라이드 타입
1. **title** - 타이틀 슬라이드 (프레젠테이션 시작)
   필드: title, subtitle, meta_line
2. **toc** - 목차 슬라이드
   필드: title, items[] (각 항목: num, text)
3. **section** - 섹션 구분 간지
   필드: section_num, section_title, section_subtitle
4. **content** - 본문 콘텐츠 슬라이드
   필드: title(제목), governance(거버넌스/섹션태그), items[] (heading=부제목 + detail=설명, 순서대로 매핑), sources[]
5. **closing** - 마무리 슬라이드
   필드: title, message, contact

## 콘텐츠 작성 규칙
1. 제목(title)은 간결하고 임팩트 있게 작성합니다 (최대 30자).
2. **governance는 해당 슬라이드의 부제목과 설명 내용을 전체적으로 요약한 문장을 작성합니다 (20~50자).** 단순한 섹션 이름이 아니라, 슬라이드 전체 내용의 핵심을 한 문장으로 압축하세요.
3. **[필수] 본문(content) 슬라이드의 구조**:
   - 본문 슬라이드는 반드시 title(제목), governance(거버넌스), items[](부제목+설명 쌍) 를 생성합니다.
   - 각 item은 heading(부제목, 키워드 1~5단어) + detail(설명, 1~3문장) 구조입니다.
   - heading은 템플릿의 "부제목" 필드에, detail은 "설명" 필드에 순서대로 매핑됩니다.
   - **카탈로그의 "items를 정확히 N개 생성하세요" 안내를 반드시 따르세요.** 부제목과 설명 필드 수 중 작은 값이 표현 가능한 최대 items 수입니다.
   - 최소 3개, 최대 4개의 items를 생성하세요. items를 1개만 작성하면 안 됩니다.
   - 각 item의 heading은 서로 다른 관점/주제를 다뤄야 합니다.
4. 예시 - subtitle_count=3, description_count=3인 content 슬라이드:
   {{"type":"content","template_index":3,"title":"디지털 전환 핵심 전략","governance":"클라우드 전환, 데이터 분석, 업무 자동화를 통한 디지털 혁신 추진",
     "items":[
       {{"heading":"클라우드 마이그레이션","detail":"기존 온프레미스 인프라를 클라우드로 전환하여 운영 비용을 30% 절감하고 확장성을 확보합니다."}},
       {{"heading":"데이터 기반 의사결정","detail":"빅데이터 분석 플랫폼을 구축하여 실시간 시장 동향 파악과 고객 행동 예측이 가능해집니다."}},
       {{"heading":"업무 자동화 도입","detail":"RPA와 AI를 활용한 반복 업무 자동화로 직원 생산성을 40% 이상 향상시킬 수 있습니다."}}
     ]}}
5. sources가 있으면 출처를 명시합니다.

## 구조 설계 규칙
1. 권장 순서: title → toc → (section → content 슬라이드들)... → closing
2. 각 섹션마다 section(간지) 슬라이드 + 1~3개의 content 슬라이드를 배치합니다.
   **[필수] 목차(toc) 슬라이드의 items 텍스트는 반드시 section 슬라이드의 section_title과 정확히 일치해야 합니다.** 예: section이 3개면 toc items도 3개이며, 각 text가 해당 section_title과 동일합니다.
3. **[필수] 전체 슬라이드가 8장 이하일 경우, 목차(toc)와 섹션 간지(section)는 생략하세요.** 구성: title → content 슬라이드들 → closing. 9장 이상일 때만 toc와 section을 포함합니다.
4. 리소스 내용의 양과 복잡도에 맞게 슬라이드 수를 자유롭게 결정합니다.
5. template_index는 사용 가능한 템플릿 슬라이드 번호입니다 (카탈로그 참조).
6. 같은 template_index를 여러 번 사용할 수 있지만, **같은 타입의 템플릿이 여러 개 있으면 돌아가며 다양하게 사용하세요.** 예를 들어 본문 템플릿 3,4,5번이 모두 content 타입이면 3→4→5→3 순으로 번갈아 사용합니다.
7. 콘텐츠에 맞지 않는 템플릿은 사용하지 않아도 됩니다.
8. **[필수] 카탈로그의 "템플릿 타입 현황"을 반드시 확인하세요. 미등록(✗) 타입은 절대 생성하지 마세요.**
9. {lang_instruction}
10. 반드시 JSON 형식으로만 응답합니다.{slide_count_instruction}"""

    user_prompt = f"""아래 리소스 자료를 분석하여 프레젠테이션을 설계하고 콘텐츠를 생성해주세요.

## 출력 언어
{lang_instruction}

## 사용자 지침
{instructions if instructions else "특별한 지침 없음 - 리소스 내용을 보고서 형태로 정리해주세요."}

## 리소스 자료
{resources_text[:12000]}

## 사용 가능한 템플릿 슬라이드 카탈로그
{slides_description}

## 응답 형식
반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트는 포함하지 마세요.

```json
{{
    "meta": {{
        "title": "프레젠테이션 제목",
        "subtitle": "부제목",
        "author": "작성자/부서",
        "date": "날짜"
    }},
    "slides": [
        {{
            "type": "title",
            "template_index": 0,
            "title": "프레젠테이션 제목",
            "subtitle": "부제목",
            "meta_line": "작성자 | 날짜"
        }},
        {{
            "type": "toc",
            "template_index": 1,
            "title": "목차",
            "items": [
                {{"num": "01", "text": "섹션 제목"}},
                {{"num": "02", "text": "섹션 제목"}}
            ]
        }},
        {{
            "type": "section",
            "template_index": 2,
            "section_num": "01",
            "section_title": "섹션 제목",
            "section_subtitle": "섹션 부제목"
        }},
        {{
            "type": "content",
            "template_index": 3,
            "title": "슬라이드 제목",
            "governance": "핵심 포인트들의 내용을 종합적으로 요약한 문장 (20~50자)",
            "items": [
                {{"heading": "첫 번째 핵심 포인트", "detail": "첫 번째 포인트에 대한 상세 설명. 구체적인 수치나 사례를 포함하여 1~3문장으로 작성합니다."}},
                {{"heading": "두 번째 핵심 포인트", "detail": "두 번째 포인트에 대한 상세 설명. 논리적 근거와 함께 1~3문장으로 작성합니다."}},
                {{"heading": "세 번째 핵심 포인트", "detail": "세 번째 포인트에 대한 상세 설명. 결론이나 시사점을 1~3문장으로 작성합니다."}}
            ],
            "sources": ["출처1"]
        }},
        {{
            "type": "closing",
            "template_index": 4,
            "title": "감사합니다",
            "message": "마무리 메시지",
            "contact": "팀명 | 이메일"
        }}
    ],
    "sources": [
        {{"ref": "source_id", "title": "출처 제목"}}
    ]
}}
```"""

    try:
        result = await _call_claude_api(system_prompt, user_prompt)
        parsed = _parse_rich_schema(result, slides_meta)
        if parsed:
            return parsed
        fallback = _fallback_content(resources_text, slides_meta)
        return {"slides": fallback, "meta": {}, "sources": []}
    except Exception as e:
        print(f"[LLM] Claude API 호출 실패: {e}")
        fallback = _fallback_content(resources_text, slides_meta)
        return {"slides": fallback, "meta": {}, "sources": []}


async def _call_claude_api(system_prompt: str, user_prompt: str) -> str:
    """Claude API 호출 (httpx 비동기)"""
    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": settings.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    payload = {
        "model": settings.ANTHROPIC_MODEL,
        "max_tokens": settings.ANTHROPIC_MAX_TOKENS,
        "system": system_prompt,
        "messages": [
            {"role": "user", "content": user_prompt}
        ],
    }

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
        if content_type == "body" and effective_items > 0:
            lines.append(f"  ※ 본문 슬라이드: items를 정확히 {effective_items}개 생성하세요 (부제목 {ph_sub_actual}개, 설명 {ph_desc_actual}개 필드가 있으므로 {effective_items}개까지 표현 가능)")
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

    # ```json ... ``` 블록 추출
    if "```json" in text:
        start = text.index("```json") + 7
        end = text.index("```", start)
        text = text[start:end].strip()
    elif "```" in text:
        start = text.index("```") + 3
        end = text.index("```", start)
        text = text[start:end].strip()

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

    # 리치 스키마 형식: {"meta": {...}, "slides": [...], "sources": [...]}
    if isinstance(parsed, dict) and "slides" in parsed:
        schema_slides = parsed.get("slides", [])
        if isinstance(schema_slides, list) and schema_slides:
            # 첫 번째 항목에 "type" 필드가 있으면 리치 스키마
            if "type" in schema_slides[0]:
                # content 슬라이드의 items 최소 개수 보장
                _ensure_minimum_items(schema_slides, slides_meta)
                # 목차(toc) items를 section 제목으로 보정
                _ensure_toc_items(schema_slides)
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


def _ensure_minimum_items(schema_slides: list[dict], slides_meta: list[dict]):
    """content 슬라이드의 items가 부족한 경우 detail을 분할하여 보충

    LLM이 items를 1개만 생성한 경우, detail 텍스트를 문장 단위로 분할하여
    최소 description_count 또는 3개의 items를 생성합니다.
    """
    # slides_meta에서 description_count 조회용 맵
    meta_lookup = {}
    for sm in slides_meta:
        meta_lookup[sm["slide_index"]] = sm.get("slide_meta", {})

    for slide in schema_slides:
        if slide.get("type") != "content":
            continue

        items = slide.get("items", [])
        if len(items) >= 2:
            continue  # 이미 충분

        # 목표 items 수 결정
        template_idx = slide.get("template_index", 0)
        meta = meta_lookup.get(template_idx, {})
        target_count = max(meta.get("description_count", 3), 3)

        if len(items) == 1:
            # 1개 item의 detail을 문장 단위로 분할 시도
            item = items[0]
            detail = item.get("detail", "")
            heading = item.get("heading", "")

            # 문장 분할 (. 기준)
            sentences = [s.strip() for s in re.split(r'(?<=[.!?。])\s+', detail) if s.strip()]

            if len(sentences) >= target_count:
                # 문장을 target_count 그룹으로 분배하여 새 items 생성
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
                # 문장이 부족하면 원본 유지 (LLM이 생성한 것 그대로)
                print(f"[LLM] Warning: content slide items={len(items)}, target={target_count}")

        elif len(items) == 0:
            print(f"[LLM] Warning: content slide has no items")


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

    for ph in placeholders:
        role = ph.get("role", "")
        name = ph.get("placeholder", "")
        if not name:
            continue
        if role == "description":
            desc_placeholders.append(name)
        elif role == "subtitle":
            subtitle_placeholders.append(name)
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
        _map_if("title", role_map, contents, slide.get("section_title", ""))
        if subtitle_placeholders and slide.get("section_subtitle"):
            contents[subtitle_placeholders[0]] = slide["section_subtitle"]
        _map_if("governance", role_map, contents, slide.get("section_num", ""))
        _map_if("body", role_map, contents, slide.get("section_subtitle", ""))

    elif slide_type == "content":
        # 제목, 거버넌스는 무조건 매핑 (필드가 있으면)
        _map_if("title", role_map, contents, slide.get("title", ""))
        _map_if("governance", role_map, contents, slide.get("governance", ""))

        items = slide.get("items", [])

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

    # 빈 placeholder에 대한 처리: 매핑되지 않은 placeholder는 빈 문자열
    for ph in placeholders:
        name = ph.get("placeholder", "")
        if name and name not in contents:
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

def _build_generation_prompts(
    resources_text: str, instructions: str, slides_meta: list[dict], lang: str = "",
    slide_count: str = "auto",
) -> tuple[str, str]:
    """generate_slide_content와 동일한 시스템/유저 프롬프트 생성 (스트리밍용)"""
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

    system_prompt = f"""당신은 기업용 프레젠테이션 구조 설계 및 콘텐츠 전문가입니다.
주어진 리소스 자료를 분석하여 전문적인 프레젠테이션을 설계합니다.

## 슬라이드 타입
1. **title** - 타이틀 슬라이드 (프레젠테이션 시작)
   필드: title, subtitle, meta_line
2. **toc** - 목차 슬라이드
   필드: title, items[] (각 항목: num, text)
3. **section** - 섹션 구분 간지
   필드: section_num, section_title, section_subtitle
4. **content** - 본문 콘텐츠 슬라이드
   필드: title(제목), governance(거버넌스/섹션태그), items[] (heading=부제목 + detail=설명, 순서대로 매핑), sources[]
5. **closing** - 마무리 슬라이드
   필드: title, message, contact

## 콘텐츠 작성 규칙
1. 제목(title)은 간결하고 임팩트 있게 작성합니다 (최대 30자).
2. **governance는 해당 슬라이드의 부제목과 설명 내용을 전체적으로 요약한 문장을 작성합니다 (20~50자).** 단순한 섹션 이름이 아니라, 슬라이드 전체 내용의 핵심을 한 문장으로 압축하세요.
3. **[필수] 본문(content) 슬라이드의 구조**:
   - 본문 슬라이드는 반드시 title(제목), governance(거버넌스), items[](부제목+설명 쌍) 를 생성합니다.
   - 각 item은 heading(부제목, 키워드 1~5단어) + detail(설명, 1~3문장) 구조입니다.
   - heading은 템플릿의 "부제목" 필드에, detail은 "설명" 필드에 순서대로 매핑됩니다.
   - **카탈로그의 "items를 정확히 N개 생성하세요" 안내를 반드시 따르세요.** 부제목과 설명 필드 수 중 작은 값이 표현 가능한 최대 items 수입니다.
   - 최소 3개, 최대 4개의 items를 생성하세요. items를 1개만 작성하면 안 됩니다.
   - 각 item의 heading은 서로 다른 관점/주제를 다뤄야 합니다.
4. 예시 - subtitle_count=3, description_count=3인 content 슬라이드:
   {{"type":"content","template_index":3,"title":"디지털 전환 핵심 전략","governance":"클라우드 전환, 데이터 분석, 업무 자동화를 통한 디지털 혁신 추진",
     "items":[
       {{"heading":"클라우드 마이그레이션","detail":"기존 온프레미스 인프라를 클라우드로 전환하여 운영 비용을 30% 절감하고 확장성을 확보합니다."}},
       {{"heading":"데이터 기반 의사결정","detail":"빅데이터 분석 플랫폼을 구축하여 실시간 시장 동향 파악과 고객 행동 예측이 가능해집니다."}},
       {{"heading":"업무 자동화 도입","detail":"RPA와 AI를 활용한 반복 업무 자동화로 직원 생산성을 40% 이상 향상시킬 수 있습니다."}}
     ]}}
5. sources가 있으면 출처를 명시합니다.

## 구조 설계 규칙
1. 권장 순서: title → toc → (section → content 슬라이드들)... → closing
2. 각 섹션마다 section(간지) 슬라이드 + 1~3개의 content 슬라이드를 배치합니다.
   **[필수] 목차(toc) 슬라이드의 items 텍스트는 반드시 section 슬라이드의 section_title과 정확히 일치해야 합니다.** 예: section이 3개면 toc items도 3개이며, 각 text가 해당 section_title과 동일합니다.
3. **[필수] 전체 슬라이드가 8장 이하일 경우, 목차(toc)와 섹션 간지(section)는 생략하세요.** 구성: title → content 슬라이드들 → closing. 9장 이상일 때만 toc와 section을 포함합니다.
4. 리소스 내용의 양과 복잡도에 맞게 슬라이드 수를 자유롭게 결정합니다.
5. template_index는 사용 가능한 템플릿 슬라이드 번호입니다 (카탈로그 참조).
6. 같은 template_index를 여러 번 사용할 수 있지만, **같은 타입의 템플릿이 여러 개 있으면 돌아가며 다양하게 사용하세요.** 예를 들어 본문 템플릿 3,4,5번이 모두 content 타입이면 3→4→5→3 순으로 번갈아 사용합니다.
7. 콘텐츠에 맞지 않는 템플릿은 사용하지 않아도 됩니다.
8. **[필수] 카탈로그의 "템플릿 타입 현황"을 반드시 확인하세요. 미등록(✗) 타입은 절대 생성하지 마세요.**
9. {lang_instruction}
10. 반드시 JSON 형식으로만 응답합니다.{slide_count_instruction}"""

    user_prompt = f"""아래 리소스 자료를 분석하여 프레젠테이션을 설계하고 콘텐츠를 생성해주세요.

## 출력 언어
{lang_instruction}

## 사용자 지침
{instructions if instructions else "특별한 지침 없음 - 리소스 내용을 보고서 형태로 정리해주세요."}

## 리소스 자료
{resources_text[:12000]}

## 사용 가능한 템플릿 슬라이드 카탈로그
{slides_description}

## 응답 형식
반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트는 포함하지 마세요.

```json
{{
    "meta": {{
        "title": "프레젠테이션 제목",
        "subtitle": "부제목",
        "author": "작성자/부서",
        "date": "날짜"
    }},
    "slides": [
        {{
            "type": "content",
            "template_index": 3,
            "title": "슬라이드 제목",
            "governance": "핵심 포인트들의 내용을 종합적으로 요약한 문장 (20~50자)",
            "items": [
                {{"heading": "첫 번째 핵심 포인트", "detail": "상세 설명 1~3문장."}},
                {{"heading": "두 번째 핵심 포인트", "detail": "상세 설명 1~3문장."}},
                {{"heading": "세 번째 핵심 포인트", "detail": "상세 설명 1~3문장."}}
            ],
            "sources": ["출처1"]
        }}
    ],
    "sources": [
        {{"ref": "source_id", "title": "출처 제목"}}
    ]
}}
```"""

    return system_prompt, user_prompt


async def _stream_claude_api(system_prompt: str, user_prompt: str):
    """Claude API 스트리밍 호출 - text delta를 async yield"""
    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": settings.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    payload = {
        "model": settings.ANTHROPIC_MODEL,
        "max_tokens": settings.ANTHROPIC_MAX_TOKENS,
        "system": system_prompt,
        "messages": [{"role": "user", "content": user_prompt}],
        "stream": True,
    }

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

    system_prompt, user_prompt = _build_generation_prompts(
        resources_text, instructions, slides_meta, lang,
        slide_count=slide_count,
    )

    try:
        full_text = ""
        async for delta in _stream_claude_api(system_prompt, user_prompt):
            full_text += delta
            yield ("delta", delta)

        parsed = _parse_rich_schema(full_text, slides_meta)
        if parsed:
            yield ("result", parsed)
        else:
            yield ("result", {
                "slides": _fallback_content(resources_text, slides_meta),
                "meta": {},
                "sources": [],
            })
    except Exception as e:
        print(f"[LLM] Streaming 호출 실패: {e}")
        yield ("result", {
            "slides": _fallback_content(resources_text, slides_meta),
            "meta": {},
            "sources": [],
        })
