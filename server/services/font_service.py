"""폰트 글리프 미리보기 PNG 생성 서비스

업로드된 폰트 파일(.ttf/.otf 등)에서 샘플 글리프를 렌더링한 PNG를 생성한다.
Structurer 멀티모달 LLM 호출 시 폰트의 시각적 특성을 LLM 컨텍스트에
주입하기 위한 핵심 빌딩 블록.
"""

import asyncio
import os
from typing import Optional

from PIL import Image, ImageDraw, ImageFont


# 캔버스 사양
_CANVAS_WIDTH = 512
_CANVAS_HEIGHT = 256
_BG_COLOR = (255, 255, 255)
_FG_COLOR = (11, 30, 63)  # #0B1E3F (브랜드 다크 네이비)
_LABEL_COLOR = (150, 150, 150)

# 미리보기 텍스트
_LATIN_SAMPLE = "AaBbCcDd 1234"
_HANGUL_SAMPLE = "가나다라 마바사"

# 폰트 크기 (px - PIL ImageFont 는 px 단위)
_LATIN_SIZE = 56
_HANGUL_SIZE = 40
_LABEL_SIZE = 14


def _safe_load_font(font_path: str, size: int) -> Optional[ImageFont.FreeTypeFont]:
    """폰트 로드 (실패 시 None)"""
    try:
        return ImageFont.truetype(font_path, size)
    except Exception as e:
        print(f"[FontService] 폰트 로드 실패 ({font_path}, size={size}): {e}")
        return None


def _supports_hangul(font: ImageFont.FreeTypeFont) -> bool:
    """폰트가 한글 글리프를 가지고 있는지 판단

    textbbox 결과로 '가나다' 글리프 너비가 의미 있는 값(>5px)인지 확인.
    한글 미지원 폰트는 보통 .notdef(빈 박스) 글리프만 가지므로 너비가 0에 가깝거나,
    오히려 폰트에 따라 박스 글리프로 인해 너비가 잡힐 수 있다 — 이 경우는
    cmap 조회로 한글 코드 포인트 매핑 여부를 추가 검사한다.
    """
    try:
        # PIL ImageFont 는 내부 freetype face 객체에 접근 가능
        face = getattr(font, "font", None)
        if face is not None and hasattr(face, "getsize"):
            # 가(U+AC00) 글리프 너비 측정
            try:
                w, _ = face.getsize("가")
                if w <= 1:
                    return False
            except Exception:
                pass
        # fontTools 가 있으면 cmap 으로 확정 검증
        try:
            from fontTools.ttLib import TTFont  # type: ignore
            tt = TTFont(font.path, lazy=True)
            try:
                for table in tt["cmap"].tables:
                    if 0xAC00 in table.cmap:  # '가'
                        return True
                return False
            finally:
                tt.close()
        except Exception:
            # fontTools 없거나 실패 — 휴리스틱(getbbox)으로 결정
            try:
                bbox = font.getbbox("가")
                width = bbox[2] - bbox[0]
                height = bbox[3] - bbox[1]
                return width > 5 and height > 5
            except Exception:
                return False
    except Exception:
        return False


def _measure_text_width(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont) -> tuple[int, int]:
    """텍스트 (width, height) 측정 (PIL 버전 호환)"""
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        return (bbox[2] - bbox[0], bbox[3] - bbox[1])
    except Exception:
        try:
            return font.getsize(text)
        except Exception:
            return (len(text) * 8, 16)


def _render_preview_sync(font_path: str, output_path: str, family: str = "") -> str:
    """동기 렌더링 (블로킹 PIL 호출). asyncio.to_thread 로 호출 권장."""
    canvas = Image.new("RGB", (_CANVAS_WIDTH, _CANVAS_HEIGHT), _BG_COLOR)
    draw = ImageDraw.Draw(canvas)

    latin_font = _safe_load_font(font_path, _LATIN_SIZE)
    if latin_font is None:
        # 폰트를 아예 못 읽으면 placeholder PNG 라도 만들어준다 (LLM 요청 흐름 깨지지 않게)
        try:
            fallback = ImageFont.load_default()
        except Exception:
            fallback = None
        if fallback is not None:
            draw.text((20, 110), "FONT LOAD FAILED", fill=(180, 60, 60), font=fallback)
        canvas.save(output_path, "PNG", optimize=True)
        return output_path

    # ── line 1: 영문/숫자 샘플 ──
    line1_w, line1_h = _measure_text_width(draw, _LATIN_SAMPLE, latin_font)
    line1_x = max(20, (_CANVAS_WIDTH - line1_w) // 2)
    line1_y = 40
    draw.text((line1_x, line1_y), _LATIN_SAMPLE, fill=_FG_COLOR, font=latin_font)

    # ── line 2: 한글 샘플 (지원 시) ──
    hangul_font = _safe_load_font(font_path, _HANGUL_SIZE)
    drew_hangul = False
    if hangul_font is not None and _supports_hangul(hangul_font):
        h_w, h_h = _measure_text_width(draw, _HANGUL_SAMPLE, hangul_font)
        if h_w > 5:  # 실제로 글리프가 그려질 수 있는 경우만
            h_x = max(20, (_CANVAS_WIDTH - h_w) // 2)
            h_y = line1_y + line1_h + 20
            draw.text((h_x, h_y), _HANGUL_SAMPLE, fill=_FG_COLOR, font=hangul_font)
            drew_hangul = True

    # ── family 라벨 (옵션) ──
    if family:
        label_font = _safe_load_font(font_path, _LABEL_SIZE)
        if label_font is None:
            try:
                label_font = ImageFont.load_default()
            except Exception:
                label_font = None
        if label_font is not None:
            label_text = f"family: {family}"
            l_w, l_h = _measure_text_width(draw, label_text, label_font)
            l_x = max(10, (_CANVAS_WIDTH - l_w) // 2)
            l_y = _CANVAS_HEIGHT - l_h - 10
            draw.text((l_x, l_y), label_text, fill=_LABEL_COLOR, font=label_font)

    canvas.save(output_path, "PNG", optimize=True)
    return output_path


async def generate_font_preview(font_path: str, output_path: str, family: str = "") -> str:
    """폰트 파일로 글리프 미리보기 PNG 생성 (async wrapper)

    Args:
        font_path: 입력 폰트 파일 절대경로 (.ttf/.otf/.ttc/.woff/.woff2)
        output_path: 저장할 PNG 파일 절대경로
        family: 폰트 family 명 (캔버스 하단 라벨에 표시, 옵션)

    Returns:
        저장된 PNG 파일 경로 (output_path 와 동일)

    Raises:
        OSError: output_path 의 부모 디렉토리가 존재하지 않거나 쓰기 권한이 없는 경우
    """
    parent = os.path.dirname(output_path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)

    # PIL 은 동기 라이브러리이므로 별도 스레드에서 실행 (이벤트 루프 블로킹 방지)
    return await asyncio.to_thread(_render_preview_sync, font_path, output_path, family)
