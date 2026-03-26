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
) -> str | None:
    """
    Google Gemini API를 사용하여 인포그래픽 이미지를 생성합니다.

    Returns:
        생성된 이미지의 URL 경로 (실패 시 None)
    """
    if not settings.GOOGLE_API_KEY:
        print("[Infographic] GOOGLE_API_KEY가 설정되지 않았습니다.")
        return None

    prompt = _build_image_prompt(
        slide_title, slide_content, slide_type, style_hint, aspect_ratio,
        slide_number=slide_number, total_slides=total_slides,
        presentation_title=presentation_title,
    )

    print(f"\n{'='*60}")
    print(f"[Infographic] 슬라이드 이미지 생성 프롬프트")
    print(f"{'='*60}")
    print(prompt)
    print(f"{'='*60}\n")

    try:
        image_bytes, file_ext = await _call_gemini_image_api(prompt)
        if not image_bytes:
            return None

        # 이미지 파일 저장
        ext = file_ext or ".png"
        filename = f"{uuid.uuid4().hex}{ext}"
        filepath = INFOGRAPHIC_DIR / filename
        filepath.write_bytes(image_bytes)

        print(f"[Infographic] 이미지 저장 완료: {filepath}")
        return f"/uploads/infographics/{filename}"

    except Exception as e:
        print(f"[Infographic] 이미지 생성 실패: {e}")
        import traceback
        traceback.print_exc()
        return None


def _build_image_prompt(
    title: str,
    content: str,
    slide_type: str,
    style_hint: str,
    aspect_ratio: str,
    slide_number: int = 1,
    total_slides: int = 1,
    presentation_title: str = "",
) -> str:
    """파워포인트 슬라이드 이미지 생성을 위한 프롬프트 구성"""

    # 콘텐츠 요약 (너무 길면 자르기)
    content_summary = content[:1500] if content else ""

    # 첫 번째 슬라이드: 전체 내용을 대표하는 타이틀 + 임팩트 인포그래픽 (90%)
    if slide_number == 1:
        infographic_ratio = (
            "⚡ THIS IS THE TITLE INFOGRAPHIC SLIDE — the single most important and visually impactful slide. "
            "It must serve as a VISUAL EXECUTIVE SUMMARY of the ENTIRE presentation. "
            "The viewer should understand the full topic, scope, and key takeaways just from this one slide. "
            "Design requirements for this slide ONLY: "
            "- ~90% infographic visual elements: large hero icons, key metric callout numbers (big bold text), "
            "  icon grids summarizing each section, mini process flow, comparison blocks, data highlights, "
            "  visual diagrams, statistical callouts, and impactful infographic layouts "
            "- Only ~10% text — use only the title and minimal keyword labels, NO paragraphs or sentences "
            "- Make it bold, dense with information, and visually striking like a professional infographic poster "
            "- Include 4-6 key data points or takeaways from the entire presentation as large visual callouts "
            "- Use the full slide area — this slide should feel rich, comprehensive, and visually dominant "
            "- The infographic elements should tell the story — viewers should grasp the entire presentation's message through visuals alone"
        )
    else:
        infographic_ratio = (
            "This is a REPORT-STYLE summary slide. Design it as a concise executive briefing page: "
            "- Use ~50% infographic visual elements (icons, mini-charts, process arrows, "
            "comparison cards, callout boxes, key metric highlights) and ~50% text content. "
            "- Text content must be CONCISE and SUMMARIZED — use bullet points, short phrases, and key numbers. "
            "- Do NOT write long paragraphs or detailed explanations. "
            "- Present information in a structured, scannable report format: headings + short bullet points + visual data. "
            "- Emphasize key figures, percentages, and conclusions with visual callouts or bold formatting."
        )

    pres_context = f' for a presentation about "{presentation_title}"' if presentation_title else ""

    prompt = f"""Generate a presentation slide image{pres_context}.

Slide Title: {title}

Slide content:
{content_summary}

{infographic_ratio}

===== MANDATORY TEMPLATE — EVERY SLIDE MUST USE THIS EXACT SAME DESIGN =====

LAYOUT (identical on ALL slides):
- Full-width dark navy (#1B2A4A) header bar at the very top, ~12% of slide height
- Slide title displayed inside the header bar in white (#FFFFFF) bold sans-serif text
- Thin #E2E8F0 separator line directly below the header bar
- White (#FFFFFF) content area below — NO gradients, NO patterns, NO textures, NO colored backgrounds
- Left/right margins: 5%, bottom margin: 5%
- ABSOLUTELY NO slide numbers, page numbers, "Slide X/Y" text, or any footer text anywhere

COLOR PALETTE (use ONLY these exact colors on ALL slides — NO exceptions):
- #1B2A4A — header bar background, section headings
- #FFFFFF — header text, content area background, card fills
- #334155 — all body text
- #2563EB — icons, chart bars, borders, arrows, accent elements
- #E2E8F0 — card borders, divider lines, subtle backgrounds
- #DBEAFE — highlight boxes, selected item backgrounds
- #64748B — captions, labels, secondary text
FORBIDDEN: Do NOT use red, orange, green, purple, pink, yellow, teal, amber, or ANY color not listed above.

TYPOGRAPHY (same on ALL slides):
- Sans-serif font family only (Pretendard, Noto Sans KR, or Arial)
- Header title: 28-32pt bold #FFFFFF
- Content headings: 18-20pt bold #1B2A4A
- Body: 14-16pt regular #334155
- Labels: 11-12pt #64748B

VISUAL ELEMENTS (same style on ALL slides):
- Icons: flat monoline, 2px stroke, #2563EB color only
- Cards: #FFFFFF fill, 1px #E2E8F0 border, 8px rounded corners
- Charts/graphs: #2563EB fills, #E2E8F0 grid lines
- Arrows/connectors: #2563EB, clean geometric

RULES:
- Widescreen {aspect_ratio}
- NO watermarks, NO placeholder text like "Lorem ipsum"
- If the title is in Korean, ALL text in the slide MUST be in Korean
- This slide must be visually IDENTICAL in template structure to every other slide in the deck

=========================================================================="""

    if style_hint:
        prompt += f"""

⚠️ HIGHEST PRIORITY — USER STYLE OVERRIDE:
The following user-specified style OVERRIDES all default design rules above.
If this style specifies different colors, backgrounds, layouts, or aesthetics, follow the user style INSTEAD.
But still keep the style CONSISTENT across ALL slides — do not vary between slides.

{style_hint}"""

    return prompt


async def _call_gemini_image_api(prompt: str) -> tuple[bytes | None, str | None]:
    """
    Google Gemini API (google-genai SDK)를 호출하여 이미지를 생성합니다.
    gemini-3.1-flash-image-preview 모델을 사용합니다.

    Returns:
        (image_bytes, file_extension) 또는 (None, None)
    """
    client = _get_client()

    contents = [
        types.Content(
            role="user",
            parts=[types.Part.from_text(text=prompt)],
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


async def _generate_single_slide(idx, slide, style_hint, aspect_ratio, total_slides, presentation_title):
    """단일 슬라이드 이미지를 생성"""
    title = slide.get("title", "") or slide.get("section_title", "") or f"슬라이드 {idx + 1}"
    slide_type = slide.get("type", "content")
    content_text = _build_slide_content_text(slide)

    print(f"[Infographic] 슬라이드 {idx + 1}/{total_slides} 생성 시작: {title}")
    image_url = await generate_infographic_image(
        slide_title=title,
        slide_content=content_text,
        slide_type=slide_type,
        style_hint=style_hint,
        aspect_ratio=aspect_ratio,
        slide_number=idx + 1,
        total_slides=total_slides,
        presentation_title=presentation_title,
    )

    return {
        "index": idx,
        "image_url": image_url,
        "title": title,
        "slide_type": slide_type,
    }


async def generate_infographic_batch(
    outline_slides: list[dict],
    style_hint: str = "",
    aspect_ratio: str = "16:9",
):
    """
    전체 슬라이드를 한번에 병렬 생성 → 완료되는 순서대로 yield (인덱스 순서 보장).

    Yields:
        {"index": int, "image_url": str|None, "title": str}
    """
    total = len(outline_slides)

    # 첫 번째 슬라이드에서 프레젠테이션 제목 추출
    presentation_title = ""
    if outline_slides:
        first = outline_slides[0]
        presentation_title = first.get("title", "") or first.get("section_title", "") or ""

    print(f"[Infographic] 총 {total}장 전체 병렬 생성 시작 (주제: {presentation_title})")

    # 모든 슬라이드를 동시에 시작
    tasks = []
    for idx, slide in enumerate(outline_slides):
        task = asyncio.create_task(
            _generate_single_slide(idx, slide, style_hint, aspect_ratio, total, presentation_title)
        )
        tasks.append(task)

    # 완료되는 대로 수집하되, 인덱스 순서대로 yield
    results = [None] * total
    next_yield = 0

    for coro in asyncio.as_completed(tasks):
        result = await coro
        idx = result["index"]
        results[idx] = result
        print(f"[Infographic] 슬라이드 {idx + 1}/{total} 완료")

        while next_yield < total and results[next_yield] is not None:
            yield results[next_yield]
            next_yield += 1

    print(f"[Infographic] 전체 {total}장 생성 완료")
