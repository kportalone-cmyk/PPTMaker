"""
Google Gemini 기반 인포그래픽 슬라이드 이미지 생성 서비스

google-genai SDK와 gemini-3.1-flash-image-preview 모델을 사용하여
각 슬라이드 아웃라인 기반 인포그래픽 이미지를 생성합니다.
"""

import uuid
import mimetypes
import asyncio
from pathlib import Path
from google import genai
from google.genai import types
from config import settings
from routers.prompt import get_prompt_content


UPLOAD_DIR = Path(settings.UPLOAD_DIR).resolve()
INFOGRAPHIC_DIR = UPLOAD_DIR / "infographics"
INFOGRAPHIC_DIR.mkdir(parents=True, exist_ok=True)

# Gemini 클라이언트 (lazy init)
_client = None

def _get_client():
    global _client
    if _client is None:
        _client = genai.Client(api_key=settings.GOOGLE_API_KEY)
    return _client


async def generate_infographic_image(
    slide_title: str,
    slide_content: str,
    slide_type: str = "content",
    style_hint: str = "",
    aspect_ratio: str = "16:9",
    slide_number: int = 1,
    total_slides: int = 1,
    presentation_title: str = "",
    infographic_ratio: int = 40,
    reference_image: bytes | None = None,
) -> tuple[str | None, bytes | None]:
    """
    Google Gemini API를 사용하여 인포그래픽 이미지를 생성합니다.

    Returns:
        (생성된 이미지의 URL 경로, 원본 이미지 바이트) 튜플. 실패 시 (None, None)
    """
    if not settings.GOOGLE_API_KEY:
        print("[Infographic] GOOGLE_API_KEY가 설정되지 않았습니다.")
        return None, None

    prompt = await _build_image_prompt(
        slide_title, slide_content, slide_type, style_hint, aspect_ratio,
        slide_number=slide_number, total_slides=total_slides,
        presentation_title=presentation_title,
        infographic_pct=infographic_ratio,
        has_reference=reference_image is not None,
    )

    print(f"\n{'='*60}")
    print(f"[Infographic] 슬라이드 이미지 생성 프롬프트 (참조이미지: {'있음' if reference_image else '없음'})")
    print(f"{'='*60}")
    print(prompt)
    print(f"{'='*60}\n")

    try:
        image_bytes, file_ext = await _call_gemini_image_api(prompt, reference_image=reference_image)
        if not image_bytes:
            return None, None

        # 이미지 파일 저장
        ext = file_ext or ".png"
        filename = f"{uuid.uuid4().hex}{ext}"
        filepath = INFOGRAPHIC_DIR / filename
        filepath.write_bytes(image_bytes)

        print(f"[Infographic] 이미지 저장 완료: {filepath}")
        return f"/uploads/infographics/{filename}", image_bytes

    except Exception as e:
        print(f"[Infographic] 이미지 생성 실패: {e}")
        import traceback
        traceback.print_exc()
        return None, None


async def _build_image_prompt(
    title: str,
    content: str,
    slide_type: str,
    style_hint: str,
    aspect_ratio: str,
    slide_number: int = 1,
    total_slides: int = 1,
    presentation_title: str = "",
    infographic_pct: int = 40,
    has_reference: bool = False,
) -> str:
    """파워포인트 슬라이드 이미지 생성을 위한 프롬프트 구성"""

    # 콘텐츠 요약 (이미지 생성에 적합한 분량으로 제한)
    content_summary = content[:600] if content else ""

    # 첫 번째 슬라이드: 배경/장식용 인포그래픽 이미지 (텍스트 없음, 타이틀은 별도 오브젝트로 오버레이)
    if slide_number == 1:
        infographic_ratio = (
            "⚡ THIS IS A COVER SLIDE BACKGROUND IMAGE — it must contain ABSOLUTELY NO TEXT whatsoever. "
            "No title, no subtitle, no labels, no captions, no numbers, no letters — ZERO text of any kind. "
            "Design requirements for this background image: "
            "- This image will be used as a BACKGROUND behind separately overlaid title/subtitle text objects. "
            "- The CENTER and UPPER area (top 50-60%) should be relatively CLEAN with only subtle, dark-toned "
            "  design elements (soft gradients, gentle geometric patterns, faint abstract shapes) so that white text "
            "  overlaid on top remains highly readable. "
            "- The LOWER portion (bottom 30-40%) and EDGES can have rich, vibrant infographic visual elements: "
            "  abstract data visualization graphics, geometric shapes, icon clusters, flowing lines, "
            "  gradient overlays, circuit-like patterns, and decorative design accents. "
            "- Use a dark, professional color scheme (deep navy, dark blue-gray gradients) as the base. "
            "- The overall feel should be a premium, professional presentation cover slide background. "
            "- Think of it as a high-end corporate keynote backdrop — elegant, modern, and visually striking "
            "  but designed to let overlaid text be the focal point."
        )
    else:
        text_pct = 100 - infographic_pct
        infographic_ratio = (
            f"This is a REPORT-STYLE summary slide. Design it as a concise executive briefing page: "
            f"- Use ~{infographic_pct}% infographic visual elements (icons, mini-charts, process arrows, "
            f"comparison cards, callout boxes, key metric highlights) and ~{text_pct}% text content. "
            "- Text content must be CONCISE and SUMMARIZED — use bullet points, short phrases, and key numbers. "
            "- Do NOT write long paragraphs or detailed explanations. "
            "- Present information in a structured, scannable report format: headings + short bullet points + visual data. "
            "- Emphasize key figures, percentages, and conclusions with visual callouts or bold formatting."
        )

    pres_context = f' for a presentation about "{presentation_title}"' if presentation_title else ""

    # 첫 번째 슬라이드는 완전히 다른 프롬프트 사용 (순수 배경 이미지)
    if slide_number == 1:
        prompt_template = await get_prompt_content("infographic_cover_image")
        prompt = prompt_template.format(
            pres_context=pres_context,
            title=title,
            infographic_ratio=infographic_ratio,
            aspect_ratio=aspect_ratio,
        )
    else:
        prompt_template = await get_prompt_content("infographic_content_image")
        prompt = prompt_template.format(
            pres_context=pres_context,
            title=title,
            content_summary=content_summary,
            infographic_ratio=infographic_ratio,
            aspect_ratio=aspect_ratio,
        )

    if style_hint:
        style_template = await get_prompt_content("infographic_style_override")
        prompt += style_template.format(style_hint=style_hint)

    # 참조 이미지가 있을 경우 스타일 일관성 지시 추가
    if has_reference:
        prompt += (
            "\n\n⚠️ CRITICAL — STYLE REFERENCE IMAGE PROVIDED:\n"
            "A reference slide image from this same presentation is attached. "
            "You MUST match its visual style EXACTLY:\n"
            "- Same background color/gradient\n"
            "- Same header bar design and color\n"
            "- Same font styles, sizes, and colors\n"
            "- Same icon style (line weight, color)\n"
            "- Same card/box design (border, radius, shadow)\n"
            "- Same color palette throughout\n"
            "- Same overall layout structure and spacing\n"
            "Only the CONTENT (text, data, icons) should differ — the DESIGN TEMPLATE must be identical."
        )

    return prompt


async def _call_gemini_image_api(prompt: str, reference_image: bytes | None = None) -> tuple[bytes | None, str | None]:
    """
    Google Gemini API (google-genai SDK)를 호출하여 이미지를 생성합니다.
    gemini-3.1-flash-image-preview 모델을 사용합니다.

    Args:
        prompt: 이미지 생성 프롬프트
        reference_image: 스타일 참조용 이미지 바이트 (있을 경우 프롬프트와 함께 전달)

    Returns:
        (image_bytes, file_extension) 또는 (None, None)
    """
    client = _get_client()

    # 프롬프트 파트 구성 (참조 이미지가 있으면 함께 전달)
    parts = []
    if reference_image:
        parts.append(types.Part.from_bytes(data=reference_image, mime_type="image/png"))
    parts.append(types.Part.from_text(text=prompt))

    contents = [
        types.Content(
            role="user",
            parts=parts,
        ),
    ]

    generate_config = types.GenerateContentConfig(
        thinking_config=types.ThinkingConfig(
            thinking_level="MINIMAL",
        ),
        image_config=types.ImageConfig(
            aspect_ratio="16:9",
            image_size="1K",
        ),
        response_modalities=["IMAGE", "TEXT"],
    )

    # 동기 SDK를 비동기 이벤트 루프에서 실행
    def _sync_generate():
        data_buffer = None
        file_ext = None
        for chunk in client.models.generate_content_stream(
            model="gemini-3.1-flash-image-preview",
            contents=contents,
            config=generate_config,
        ):
            if chunk.parts is None:
                continue
            for part in chunk.parts:
                if part.inline_data and part.inline_data.data:
                    data_buffer = part.inline_data.data
                    file_ext = mimetypes.guess_extension(part.inline_data.mime_type) or ".png"
                elif part.text:
                    print(f"[Infographic] Gemini 텍스트 응답: {part.text[:200]}")
        return data_buffer, file_ext

    loop = asyncio.get_event_loop()
    result = await loop.run_in_executor(None, _sync_generate)

    if result[0]:
        print(f"[Infographic] 이미지 데이터 수신: {len(result[0])} bytes")
    else:
        print("[Infographic] Gemini API 응답에 이미지 데이터가 없습니다.")

    return result




def _build_slide_content_text(slide: dict) -> str:
    """아웃라인 슬라이드에서 콘텐츠 텍스트를 추출"""
    content_parts = []
    if slide.get("subtitle"):
        content_parts.append(slide["subtitle"])

    for item in slide.get("items", []):
        heading = item.get("heading", "")
        detail = item.get("detail", "")
        if heading:
            content_parts.append(f"- {heading}: {detail}" if detail else f"- {heading}")

    contents = slide.get("contents", {})
    for key, val in contents.items():
        if isinstance(val, str) and val:
            content_parts.append(val)

    if slide.get("table_data"):
        td = slide["table_data"]
        headers = td.get("headers", [])
        rows = td.get("rows", [])
        if headers:
            content_parts.append(f"Table: {', '.join(str(h) for h in headers)}")
            for row in rows[:5]:
                content_parts.append(f"  {', '.join(str(c) for c in row)}")

    if slide.get("chart_data"):
        cd = slide["chart_data"]
        chart_type = cd.get("chart_type", "bar")
        chart_title = cd.get("title", "")
        content_parts.append(f"Chart ({chart_type}): {chart_title}")
        chart_d = cd.get("chart_data", {})
        labels = chart_d.get("labels", [])
        if labels:
            content_parts.append(f"Categories: {', '.join(str(l) for l in labels)}")

    return "\n".join(content_parts)


async def _generate_single_slide(idx, slide, style_hint, aspect_ratio, total_slides, presentation_title, infographic_ratio=40, reference_image: bytes | None = None):
    """단일 슬라이드 이미지를 생성"""
    title = slide.get("title", "") or slide.get("section_title", "") or f"슬라이드 {idx + 1}"
    slide_type = slide.get("type", "content")
    content_text = _build_slide_content_text(slide)

    print(f"[Infographic] 슬라이드 {idx + 1}/{total_slides} 생성 시작: {title}")
    image_url, image_bytes = await generate_infographic_image(
        slide_title=title,
        slide_content=content_text,
        slide_type=slide_type,
        style_hint=style_hint,
        aspect_ratio=aspect_ratio,
        slide_number=idx + 1,
        total_slides=total_slides,
        presentation_title=presentation_title,
        infographic_ratio=infographic_ratio,
        reference_image=reference_image,
    )

    # 첫 번째 슬라이드의 subtitle 추출
    subtitle = ""
    if idx == 0:
        subtitle = slide.get("subtitle", "") or ""
        if not subtitle:
            items = slide.get("items", [])
            if items and isinstance(items[0], dict):
                subtitle = items[0].get("heading", "")

    return {
        "index": idx,
        "image_url": image_url,
        "image_bytes": image_bytes,
        "title": title,
        "slide_type": slide_type,
        "subtitle": subtitle,
    }


async def generate_infographic_batch(
    outline_slides: list[dict],
    style_hint: str = "",
    aspect_ratio: str = "16:9",
    infographic_ratio: int = 40,
):
    """
    스타일 일관성을 위해 첫 콘텐츠 슬라이드를 먼저 생성하고,
    그 이미지를 참조로 나머지 슬라이드를 병렬 생성합니다.

    생성 순서:
      1) 커버 슬라이드 (idx=0) — 독립 생성
      2) 첫 콘텐츠 슬라이드 (idx=1) — 독립 생성, 이미지를 참조용으로 보관
      3) 나머지 콘텐츠 슬라이드 (idx=2+) — 2번 이미지를 참조로 병렬 생성

    Yields:
        {"index": int, "image_url": str|None, "title": str}
    """
    total = len(outline_slides)

    # 첫 번째 슬라이드에서 프레젠테이션 제목 추출
    presentation_title = ""
    if outline_slides:
        first = outline_slides[0]
        presentation_title = first.get("title", "") or first.get("section_title", "") or ""

    print(f"[Infographic] 총 {total}장 생성 시작 — 스타일 참조 방식 (주제: {presentation_title})")

    reference_image_bytes = None

    # ── Phase 1: 커버(0번) + 첫 콘텐츠(1번) 병렬 생성 ──
    phase1_indices = list(range(min(2, total)))
    phase1_tasks = []
    for idx in phase1_indices:
        task = asyncio.create_task(
            _generate_single_slide(
                idx, outline_slides[idx], style_hint, aspect_ratio,
                total, presentation_title, infographic_ratio,
            )
        )
        phase1_tasks.append(task)

    phase1_results = {}
    for coro in asyncio.as_completed(phase1_tasks):
        result = await coro
        idx = result["index"]
        phase1_results[idx] = result
        print(f"[Infographic] 슬라이드 {idx + 1}/{total} 완료 (Phase 1)")

    # 첫 콘텐츠 슬라이드(idx=1)의 이미지를 참조로 저장
    if 1 in phase1_results and phase1_results[1].get("image_bytes"):
        reference_image_bytes = phase1_results[1]["image_bytes"]
        print(f"[Infographic] 스타일 참조 이미지 확보 (슬라이드 2, {len(reference_image_bytes)} bytes)")

    # Phase 1 결과를 인덱스 순서대로 yield
    for idx in sorted(phase1_results.keys()):
        yield phase1_results[idx]

    # ── Phase 2: 나머지 슬라이드 병렬 생성 (참조 이미지 포함) ──
    if total <= 2:
        print(f"[Infographic] 전체 {total}장 생성 완료")
        return

    remaining_indices = list(range(2, total))
    print(f"[Infographic] 나머지 {len(remaining_indices)}장 병렬 생성 시작 (참조 이미지 {'있음' if reference_image_bytes else '없음'})")

    tasks = []
    for idx in remaining_indices:
        task = asyncio.create_task(
            _generate_single_slide(
                idx, outline_slides[idx], style_hint, aspect_ratio,
                total, presentation_title, infographic_ratio,
                reference_image=reference_image_bytes,
            )
        )
        tasks.append(task)

    # 완료되는 대로 수집하되, 인덱스 순서대로 yield
    results = {}
    next_yield = 2

    for coro in asyncio.as_completed(tasks):
        result = await coro
        idx = result["index"]
        results[idx] = result
        print(f"[Infographic] 슬라이드 {idx + 1}/{total} 완료 (Phase 2)")

        while next_yield < total and next_yield in results:
            yield results[next_yield]
            next_yield += 1

    print(f"[Infographic] 전체 {total}장 생성 완료")
