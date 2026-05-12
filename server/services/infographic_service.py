"""
인포그래픽 슬라이드 이미지 생성 서비스

.env `IMAGE_PROVIDER` 값에 따라 Google Gemini(nanobanana) 또는
OpenAI(gpt-image-2) 이미지 API를 호출합니다.
"""

import io
import uuid
import base64
import mimetypes
import asyncio
from pathlib import Path
from google import genai
from google.genai import types
from config import settings, google_key_rotator, openai_key_rotator
from routers.prompt import get_prompt_content


UPLOAD_DIR = Path(settings.UPLOAD_DIR).resolve()
INFOGRAPHIC_DIR = UPLOAD_DIR / "infographics"
INFOGRAPHIC_DIR.mkdir(parents=True, exist_ok=True)

# Gemini 클라이언트 (키별 캐시)
_clients = {}
# OpenAI 클라이언트 (키별 캐시)
_openai_clients = {}


def _image_provider() -> str:
    """현재 이미지 생성 프로바이더 (google | openai)"""
    return (settings.IMAGE_PROVIDER or "google").strip().lower()


def _image_provider_available() -> bool:
    """현재 선택된 프로바이더의 API 키가 설정되어 있는지 확인"""
    provider = _image_provider()
    if provider == "openai":
        return bool(settings.OPENAI_API_KEYS)
    return bool(settings.GOOGLE_API_KEYS)


def _get_client():
    """라운드 로빈으로 Google API 키를 선택하여 클라이언트 반환"""
    api_key = google_key_rotator.next()
    if api_key not in _clients:
        _clients[api_key] = genai.Client(api_key=api_key)
    return _clients[api_key]


def _get_openai_client():
    """라운드 로빈으로 OpenAI API 키를 선택하여 AsyncOpenAI 클라이언트 반환"""
    from openai import AsyncOpenAI
    api_key = openai_key_rotator.next()
    if api_key not in _openai_clients:
        _openai_clients[api_key] = AsyncOpenAI(api_key=api_key)
    return _openai_clients[api_key]


def _pad_to_16x9(image_bytes: bytes) -> bytes:
    """
    이미지를 16:9 비율로 letterbox 패딩한다.
    OpenAI gpt-image는 1536x1024(3:2) 등 16:9가 아닌 landscape만 지원하므로,
    슬라이드 캔버스(16:9)에서 잘리지 않도록 가장자리 색으로 좌우(또는 상하)에 여백을 덧댄다.

    이미 16:9(오차 1% 이내)이거나 열기 실패 시 원본을 그대로 돌려준다.
    """
    try:
        from PIL import Image
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    except Exception as e:
        print(f"[Infographic] 이미지 열기 실패 (패딩 스킵): {e}")
        return image_bytes

    w, h = img.size
    target_ratio = 16 / 9
    current_ratio = w / h
    if abs(current_ratio - target_ratio) / target_ratio < 0.01:
        return image_bytes

    if current_ratio < target_ratio:
        # 가로가 부족 → 좌우에 패딩 (왼쪽 열/오른쪽 열 색으로 채움)
        new_w = int(round(h * target_ratio))
        pad_total = new_w - w
        pad_left = pad_total // 2
        pad_right = pad_total - pad_left

        left_col = img.crop((0, 0, 1, h)).resize((pad_left, h))
        right_col = img.crop((w - 1, 0, w, h)).resize((pad_right, h))

        canvas = Image.new("RGB", (new_w, h))
        canvas.paste(left_col, (0, 0))
        canvas.paste(img, (pad_left, 0))
        canvas.paste(right_col, (pad_left + w, 0))
    else:
        # 세로가 부족 → 상하에 패딩
        new_h = int(round(w / target_ratio))
        pad_total = new_h - h
        pad_top = pad_total // 2
        pad_bottom = pad_total - pad_top

        top_row = img.crop((0, 0, w, 1)).resize((w, pad_top))
        bottom_row = img.crop((0, h - 1, w, h)).resize((w, pad_bottom))

        canvas = Image.new("RGB", (w, new_h))
        canvas.paste(top_row, (0, 0))
        canvas.paste(img, (0, pad_top))
        canvas.paste(bottom_row, (0, pad_top + h))

    buf = io.BytesIO()
    canvas.save(buf, format="PNG", optimize=True)
    out = buf.getvalue()
    print(f"[Infographic] 16:9 패딩 적용: {w}x{h} → {canvas.size[0]}x{canvas.size[1]}")
    return out


async def generate_infographic_image(
    slide_title: str,
    slide_content: str,
    slide_type: str = "content",
    style_hint: str = "",
    aspect_ratio: str = "16:9",
    slide_number: int = 1,
    total_slides: int = 1,
    presentation_title: str = "",
    infographic_ratio: int = 60,
    reference_image: bytes | None = None,
) -> tuple[str | None, bytes | None]:
    """
    Google Gemini API를 사용하여 인포그래픽 이미지를 생성합니다.

    Returns:
        (생성된 이미지의 URL 경로, 원본 이미지 바이트) 튜플. 실패 시 (None, None)
    """
    if not _image_provider_available():
        print(f"[Infographic] IMAGE_PROVIDER={_image_provider()} API 키가 설정되지 않았습니다.")
        return None, None

    prompt = await _build_image_prompt(
        slide_title, slide_content, slide_type, style_hint, aspect_ratio,
        slide_number=slide_number, total_slides=total_slides,
        presentation_title=presentation_title,
        infographic_pct=infographic_ratio,
        has_reference=reference_image is not None,
    )

    print(f"[Infographic] 슬라이드 이미지 생성 요청 (provider={_image_provider()}, 참조이미지: {'있음' if reference_image else '없음'})")

    try:
        image_bytes, file_ext = await _call_image_api(prompt, reference_image=reference_image)
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
    infographic_pct: int = 60,
    has_reference: bool = False,
) -> str:
    """파워포인트 슬라이드 이미지 생성을 위한 프롬프트 구성"""

    # 콘텐츠 요약 (이미지 생성에 적합한 분량으로 제한)
    content_summary = content[:600] if content else ""

    # 슬라이드 비율/지침을 DB에서 조회 (하드코딩 금지)
    if slide_number == 1:
        infographic_ratio_prompt = await get_prompt_content("infographic_cover_ratio")
    else:
        infographic_ratio_prompt = await get_prompt_content("infographic_content_ratio")

    # infographic_pct 값에 따라 비율 지침 추가
    text_pct = 100 - infographic_pct
    ratio_instruction = f"""
VISUAL vs TEXT RATIO — STRICTLY FOLLOW:
- Infographic visuals (icons, charts, diagrams, illustrations): {infographic_pct}% of the slide area
- Text content (headings, bullet points, descriptions): {text_pct}% of the slide area"""

    if infographic_pct <= 20:
        ratio_instruction += "\n- This is a TEXT-HEAVY slide. Minimal graphics, focus on readable text with clean layout."
    elif infographic_pct <= 40:
        ratio_instruction += "\n- Balanced layout with more text than graphics. Use icons/charts as accents."
    elif infographic_pct <= 60:
        ratio_instruction += "\n- Equal balance. Use infographic elements alongside concise text."
    elif infographic_pct <= 80:
        ratio_instruction += "\n- GRAPHIC-HEAVY slide. Large charts, diagrams, icons dominate. Text is minimal labels only."
    else:
        ratio_instruction += "\n- Almost ALL GRAPHICS. Minimal text — only short labels, numbers. Visual storytelling."

    infographic_ratio = infographic_ratio_prompt + ratio_instruction

    pres_context = f' for a presentation about "{presentation_title}"' if presentation_title else ""

    # 첫 번째 슬라이드: 심플한 인포그래픽 커버 (제목+부제목)
    if slide_number == 1:
        prompt_template = await get_prompt_content("infographic_cover_image")
        prompt = prompt_template.format(
            pres_context=pres_context,
            title=title,
            content_summary=content_summary,
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

    # 참조 이미지가 있을 경우 스타일 일관성 지시 추가 (DB에서 프롬프트 조회)
    if has_reference:
        ref_prompt = await get_prompt_content("infographic_reference_image")
        prompt += "\n\n" + ref_prompt

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

    # Thinking 설정 (.env GOOGLE_IMAGE_THINKING: none/LOW/MEDIUM/HIGH)
    thinking_level = settings.GOOGLE_IMAGE_THINKING.strip().upper()
    config_kwargs = {
        "image_config": types.ImageConfig(
            aspect_ratio="16:9",
            image_size="1K",
        ),
        "response_modalities": ["IMAGE", "TEXT"],
    }
    if thinking_level and thinking_level != "NONE":
        config_kwargs["thinking_config"] = types.ThinkingConfig(
            thinking_level=thinking_level,
        )
    generate_config = types.GenerateContentConfig(**config_kwargs)

    # 동기 SDK를 비동기 이벤트 루프에서 실행
    def _sync_generate():
        data_buffer = None
        file_ext = None
        for chunk in client.models.generate_content_stream(
            model=settings.GOOGLE_IMAGE_MODEL,
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


async def _call_openai_image_api(prompt: str, reference_image: bytes | None = None) -> tuple[bytes | None, str | None]:
    """
    OpenAI Image API (AsyncOpenAI SDK)를 호출하여 이미지를 생성합니다.
    .env `OPENAI_IMAGE_MODEL` 값의 모델(gpt-image-2 등)을 사용합니다.

    Args:
        prompt: 이미지 생성 프롬프트
        reference_image: 스타일 참조용 이미지 바이트 (있을 경우 images.edit 사용)

    Returns:
        (image_bytes, file_extension) 또는 (None, None)
    """
    client = _get_openai_client()
    model = settings.OPENAI_IMAGE_MODEL
    size = settings.OPENAI_IMAGE_SIZE
    quality = settings.OPENAI_IMAGE_QUALITY

    try:
        if reference_image:
            # 참조 이미지가 있으면 edit 엔드포인트 사용
            img_file = io.BytesIO(reference_image)
            img_file.name = "reference.png"
            response = await client.images.edit(
                model=model,
                image=img_file,
                prompt=prompt,
                size=size,
                quality=quality,
                n=1,
            )
        else:
            response = await client.images.generate(
                model=model,
                prompt=prompt,
                size=size,
                quality=quality,
                n=1,
            )
    except Exception as e:
        print(f"[Infographic] OpenAI 이미지 API 호출 실패: {e}")
        return None, None

    if not response.data:
        print("[Infographic] OpenAI API 응답에 이미지 데이터가 없습니다.")
        return None, None

    item = response.data[0]
    b64 = getattr(item, "b64_json", None)
    if b64:
        image_bytes = base64.b64decode(b64)
    else:
        url = getattr(item, "url", None)
        if not url:
            print("[Infographic] OpenAI 응답에 b64_json/url 모두 없음")
            return None, None
        try:
            import httpx
            async with httpx.AsyncClient(timeout=60.0) as hx:
                r = await hx.get(url)
                r.raise_for_status()
                image_bytes = r.content
        except Exception as e:
            print(f"[Infographic] OpenAI 이미지 URL 다운로드 실패: {e}")
            return None, None

    print(f"[Infographic] OpenAI 이미지 데이터 수신: {len(image_bytes)} bytes")

    # 슬라이드 캔버스(16:9)에서 잘리지 않도록 letterbox 패딩
    image_bytes = _pad_to_16x9(image_bytes)
    return image_bytes, ".png"


async def _call_image_api(prompt: str, reference_image: bytes | None = None) -> tuple[bytes | None, str | None]:
    """
    .env `IMAGE_PROVIDER` 값에 따라 Google Gemini 또는 OpenAI 이미지 API를 호출하는 디스패처.

    Returns:
        (image_bytes, file_extension) 또는 (None, None)
    """
    provider = _image_provider()
    if provider == "openai":
        return await _call_openai_image_api(prompt, reference_image=reference_image)
    return await _call_gemini_image_api(prompt, reference_image=reference_image)


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


async def _generate_single_slide(idx, slide, style_hint, aspect_ratio, total_slides, presentation_title, infographic_ratio=60, reference_image: bytes | None = None):
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
    infographic_ratio: int = 60,
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

    # ── Phase 2: 나머지 슬라이드 병렬 생성 (동시 3개 제한, 참조 이미지 포함) ──
    if total <= 2:
        print(f"[Infographic] 전체 {total}장 생성 완료")
        return

    remaining_indices = list(range(2, total))
    print(f"[Infographic] 나머지 {len(remaining_indices)}장 생성 시작 (동시 3개, 참조 이미지 {'있음' if reference_image_bytes else '없음'})")

    # 동시 실행 수 제한 (3개씩 배치)
    BATCH_SIZE = 5
    next_yield = 2
    results = {}

    for batch_start in range(0, len(remaining_indices), BATCH_SIZE):
        batch = remaining_indices[batch_start:batch_start + BATCH_SIZE]
        tasks = []
        for idx in batch:
            task = asyncio.create_task(
                _generate_single_slide(
                    idx, outline_slides[idx], style_hint, aspect_ratio,
                    total, presentation_title, infographic_ratio,
                    reference_image=reference_image_bytes,
                )
            )
            tasks.append(task)

        for coro in asyncio.as_completed(tasks):
            result = await coro
            idx = result["index"]
            results[idx] = result
            print(f"[Infographic] 슬라이드 {idx + 1}/{total} 완료 (Phase 2, batch {batch_start // BATCH_SIZE + 1})")

            while next_yield < total and next_yield in results:
                yield results[next_yield]
                next_yield += 1

    print(f"[Infographic] 전체 {total}장 생성 완료")


async def fix_slide_text_image(image_bytes: bytes) -> tuple[str | None, bytes | None]:
    """
    기존 슬라이드 이미지의 깨진 텍스트를 수정하여 재생성합니다.
    원본 이미지를 이미지 프로바이더에 보내 동일 디자인 + 수정된 텍스트로 재생성.
    """
    if not _image_provider_available():
        return None, None

    prompt = await get_prompt_content("fix_slide_text")

    print(f"[FixText] 텍스트 수정 이미지 재생성 시작 (provider={_image_provider()})")

    try:
        new_image_bytes, file_ext = await _call_image_api(
            prompt, reference_image=image_bytes,
        )
        if not new_image_bytes:
            return None, None

        ext = file_ext or ".png"
        filename = f"fixed_{uuid.uuid4().hex}{ext}"
        filepath = INFOGRAPHIC_DIR / filename
        filepath.write_bytes(new_image_bytes)

        print(f"[FixText] 텍스트 수정 이미지 저장 완료: {filepath}")
        return f"/uploads/infographics/{filename}", new_image_bytes
    except Exception as e:
        print(f"[FixText] 텍스트 수정 실패: {e}")
        import traceback
        traceback.print_exc()
        return None, None


async def edit_slide_image(image_bytes: bytes, instruction: str) -> tuple[str | None, bytes | None]:
    """
    사용자 지침에 따라 인포그래픽 슬라이드 이미지를 수정하여 재생성합니다.
    원본 이미지 + 사용자 지침을 이미지 프로바이더에 보내 수정된 이미지를 생성.
    """
    if not _image_provider_available():
        return None, None

    prompt_template = await get_prompt_content("edit_slide_image")
    prompt = prompt_template.format(instruction=instruction)

    print(f"[EditSlide] 슬라이드 이미지 수정 요청 (provider={_image_provider()}): {instruction[:80]}...")

    try:
        new_image_bytes, file_ext = await _call_image_api(
            prompt, reference_image=image_bytes,
        )
        if not new_image_bytes:
            return None, None

        ext = file_ext or ".png"
        filename = f"edited_{uuid.uuid4().hex}{ext}"
        filepath = INFOGRAPHIC_DIR / filename
        filepath.write_bytes(new_image_bytes)

        print(f"[EditSlide] 수정 이미지 저장 완료: {filepath}")
        return f"/uploads/infographics/{filename}", new_image_bytes
    except Exception as e:
        print(f"[EditSlide] 슬라이드 이미지 수정 실패: {e}")
        import traceback
        traceback.print_exc()
        return None, None


async def generate_bg_image(
    bg_prompt: str,
    style_hint: str = "",
    aspect_ratio: str = "16:9",
    reference_image: bytes | None = None,
) -> tuple[str | None, bytes | None]:
    """
    배경 이미지만 생성 (텍스트 없이 추상적/테마 배경).
    AI 슬라이드 모드에서 사용.
    """
    if not _image_provider_available():
        return None, None

    prompt_template = await get_prompt_content("ai_slide_bg_image")
    style_section = ""
    if style_hint:
        style_section = f"\nStyle guide: {style_hint}"

    prompt = prompt_template.format(
        bg_prompt=bg_prompt,
        style_hint=style_section,
        aspect_ratio=aspect_ratio,
    )

    if reference_image:
        prompt += (
            "\n\n⚠️ STYLE REFERENCE: A reference background image is attached. "
            "Match its color palette, mood, and visual style exactly. "
            "Only change the specific visual elements described above."
        )

    print(f"[AI Slide BG] 배경 이미지 생성 (provider={_image_provider()}): {bg_prompt[:80]}...")

    try:
        image_bytes, file_ext = await _call_image_api(prompt, reference_image=reference_image)
        if not image_bytes:
            return None, None

        ext = file_ext or ".png"
        filename = f"bg_{uuid.uuid4().hex}{ext}"
        filepath = INFOGRAPHIC_DIR / filename
        filepath.write_bytes(image_bytes)

        return f"/uploads/infographics/{filename}", image_bytes
    except Exception as e:
        print(f"[AI Slide BG] 배경 이미지 생성 실패: {e}")
        return None, None


async def generate_summary_infographic(
    summary_data: dict,
    style_hint: str = "",
    infographic_ratio: int = 60,
) -> tuple[str | None, bytes | None]:
    """
    한장 요약 인포그래픽 이미지를 생성합니다.
    Claude가 요약한 structured data를 Gemini 이미지로 변환.

    Args:
        summary_data: Claude가 생성한 요약 JSON (title, subtitle, sections, key_metrics, conclusion)
        style_hint: 스타일 힌트

    Returns:
        (이미지 URL, 이미지 bytes) 또는 (None, None)
    """
    if not _image_provider_available():
        print(f"[SummaryInfographic] IMAGE_PROVIDER={_image_provider()} API 키가 설정되지 않았습니다.")
        return None, None

    title = summary_data.get("title", "Summary")
    subtitle = summary_data.get("subtitle", "")
    sections = summary_data.get("sections", [])
    key_metrics = summary_data.get("key_metrics", [])
    conclusion = summary_data.get("conclusion", "")
    flow = summary_data.get("flow", {})
    color_scheme = summary_data.get("color_scheme", "auto")

    # sections 텍스트 구성 — 그래픽 요소 중심, 텍스트 최소화
    sections_text_parts = []
    for i, sec in enumerate(sections, 1):
        heading = sec.get("heading", f"Section {i}")
        icon = sec.get("icon_hint", "")
        visual_type = sec.get("visual_type", "stat_card")
        highlight = sec.get("highlight_value", "")

        # data_points (새 구조) 또는 points (구 구조) 지원
        data_points = sec.get("data_points", [])
        points = sec.get("points", [])

        section_line = f"{i}. [{icon} icon] {heading} → draw as [{visual_type}]"
        if highlight:
            section_line += f" — show {highlight} LARGE"

        # data_points 우선, 없으면 points fallback
        if data_points:
            dp_lines = [f"  {dp.get('label','')}: {dp.get('value','')}" for dp in data_points[:4]]
            sections_text_parts.append(section_line + "\n" + "\n".join(dp_lines))
        elif points:
            pt_lines = [f"  {p}" for p in points[:3]]
            sections_text_parts.append(section_line + "\n" + "\n".join(pt_lines))
        else:
            sections_text_parts.append(section_line)
    sections_text = "\n\n".join(sections_text_parts)

    # metrics 텍스트 구성 (색상/아이콘 힌트 포함)
    metrics_text = ""
    if key_metrics:
        metrics_lines = []
        for m in key_metrics:
            label = m.get("label", "")
            value = m.get("value", "")
            icon = m.get("icon_hint", "")
            color = m.get("color_hint", "blue")
            metrics_lines.append(f"  - [{icon} icon, {color} card] {label}: {value} (display value LARGE)")
        metrics_text = "Key Metric Cards (display as prominent stat cards with large numbers):\n" + "\n".join(metrics_lines)

    # flow 텍스트 구성
    flow_text = ""
    flow_type = flow.get("type", "none") if flow else "none"
    if flow_type and flow_type != "none":
        steps = flow.get("steps", [])
        flow_desc = flow.get("description", "")
        if steps:
            steps_str = " → ".join(steps)
            flow_text = f"Process/Flow Visualization ({flow_type}):\n  {steps_str}"
            if flow_desc:
                flow_text += f"\n  Description: {flow_desc}"
            flow_text += "\n  [Render as connected visual nodes with arrows, NOT plain text]"

    # 프롬프트 구성
    prompt_template = await get_prompt_content("summary_infographic_image")
    prompt = prompt_template.format(
        title=title,
        subtitle=subtitle,
        sections_text=sections_text,
        metrics_text=metrics_text,
        conclusion=conclusion,
        flow_text=flow_text,
        color_scheme=color_scheme,
    )

    if style_hint:
        style_template = await get_prompt_content("infographic_style_override")
        prompt += style_template.format(style_hint=style_hint)

    # 인포그래픽 비율 지침 추가
    text_pct = 100 - infographic_ratio
    ratio_instruction = f"""
VISUAL vs TEXT RATIO — STRICTLY FOLLOW:
- Infographic visuals (icons, charts, diagrams, illustrations): {infographic_ratio}% of the image area
- Text content (headings, bullet points, descriptions): {text_pct}% of the image area"""
    if infographic_ratio <= 20:
        ratio_instruction += "\n- This is a TEXT-HEAVY layout. Minimal graphics, focus on readable text with clean layout."
    elif infographic_ratio <= 40:
        ratio_instruction += "\n- Balanced layout with more text than graphics. Use icons/charts as accents."
    elif infographic_ratio <= 60:
        ratio_instruction += "\n- Equal balance. Use infographic elements alongside concise text."
    elif infographic_ratio <= 80:
        ratio_instruction += "\n- GRAPHIC-HEAVY layout. Large charts, diagrams, icons dominate. Text is minimal labels only."
    else:
        ratio_instruction += "\n- Almost ALL GRAPHICS. Minimal text — only short labels, numbers. Visual storytelling."
    prompt += "\n" + ratio_instruction

    print(f"[SummaryInfographic] 한장 요약 인포그래픽 이미지 생성 요청 (provider={_image_provider()})")

    try:
        image_bytes, file_ext = await _call_image_api(prompt)
        if not image_bytes:
            return None, None

        ext = file_ext or ".png"
        filename = f"summary_{uuid.uuid4().hex}{ext}"
        filepath = INFOGRAPHIC_DIR / filename
        filepath.write_bytes(image_bytes)

        print(f"[SummaryInfographic] 이미지 저장 완료: {filepath}")
        return f"/uploads/infographics/{filename}", image_bytes

    except Exception as e:
        print(f"[SummaryInfographic] 이미지 생성 실패: {e}")
        import traceback
        traceback.print_exc()
        return None, None
