"""PPT 스타일 자산 서비스 (M3 빌더용)

배경 그라데이션 JPG 생성 + 슬라이드별 그라데이션 원점 매핑 + 간이 아이콘 PNG 생성.

참고 문서: 데스크탑 "PPT생성스.md" 4.3절 (배경 그라데이션) / 5절 (12개 패턴) /
6절 (헬퍼 함수). 좌표·색·원점값은 해당 문서 그대로 포팅.

외부 라이브러리 제약: numpy, Pillow 만 사용 (이미 requirements.txt 에 존재).
react-icons / cairosvg 등 추가 의존성 금지 → 아이콘은 PIL 로 원+문자 형태의
placeholder 만 생성하고 M5 에서 정식 SVG 세트로 교체 가능하도록 시그니처 안정 유지.
"""

from __future__ import annotations

import io
import os
from pathlib import Path
from typing import Optional

import numpy as np
from PIL import Image, ImageDraw, ImageFont


# ============ 색상 헬퍼 ============

def _hex_to_np(hex_str: str) -> np.ndarray:
    """'#RRGGBB' 또는 'RRGGBB' → np.float32 [r, g, b]"""
    if not isinstance(hex_str, str):
        return np.array([255.0, 255.0, 255.0], dtype=np.float32)
    s = hex_str.strip().lstrip("#")
    if len(s) != 6:
        return np.array([255.0, 255.0, 255.0], dtype=np.float32)
    try:
        r = int(s[0:2], 16)
        g = int(s[2:4], 16)
        b = int(s[4:6], 16)
    except ValueError:
        return np.array([255.0, 255.0, 255.0], dtype=np.float32)
    return np.array([r, g, b], dtype=np.float32)


def _hex_to_rgb_tuple(hex_str: str) -> tuple:
    """'#RRGGBB' → (r, g, b) int tuple"""
    arr = _hex_to_np(hex_str)
    return (int(arr[0]), int(arr[1]), int(arr[2]))


# ============ 그라데이션 코어 (.md 4.3 그대로 포팅) ============

def _make_meshgrid(W: int, H: int):
    xs = np.arange(W, dtype=np.float32)
    ys = np.arange(H, dtype=np.float32)
    X, Y = np.meshgrid(xs, ys)
    return X, Y


def _lerp(a: np.ndarray, b: np.ndarray, t: np.ndarray) -> np.ndarray:
    """선형 보간 - t는 (H,W) float, a/b는 (3,) float → (H,W,3)"""
    return a + (b - a) * t[..., None]


def _radial(
    W: int, H: int, X: np.ndarray, Y: np.ndarray,
    ox_n: float, oy_n: float,
    intensity: float, gamma: float,
    c_in: np.ndarray, c_out: np.ndarray,
) -> np.ndarray:
    """라디얼 그라데이션 - 정규화 원점(ox_n, oy_n) 기준 거리에 따라 c_in→c_out 보간"""
    cx, cy = W * ox_n, H * oy_n
    d = np.sqrt((X - cx) ** 2 + (Y - cy) ** 2)
    max_r = np.sqrt(W * W + H * H) * intensity
    t = np.clip(d / max_r, 0.0, 1.0) ** gamma
    return _lerp(c_in, c_out, t)


def _linear(
    W: int, H: int, X: np.ndarray, Y: np.ndarray,
    ox_n: float, oy_n: float,
    angle_deg: float,
    intensity: float, gamma: float,
    c_in: np.ndarray, c_out: np.ndarray,
) -> np.ndarray:
    """리니어 그라데이션 - 원점에서 angle_deg 방향으로 선형 진행"""
    rad = np.deg2rad(angle_deg)
    # 방향벡터
    dx, dy = np.cos(rad), np.sin(rad)
    cx, cy = W * ox_n, H * oy_n
    # 원점 기준 projection
    proj = (X - cx) * dx + (Y - cy) * dy
    diag = np.sqrt(W * W + H * H) * intensity
    t = np.clip((proj + diag / 2.0) / diag, 0.0, 1.0) ** gamma
    return _lerp(c_in, c_out, t)


def _save_jpg(arr: np.ndarray, path: str) -> None:
    arr = np.clip(arr, 0, 255).astype(np.uint8)
    Image.fromarray(arr).save(path, "JPEG", quality=92, optimize=True)


# ============ Public API ============

async def generate_background(
    output_path: str,
    canvas_size: tuple = (1920, 1080),
    style: str = "radial",
    origin: tuple = (0.3, 0.45),
    intensity: float = 0.85,
    gamma: float = 1.3,
    color_in: str = "#DBE8FE",
    color_out: str = "#FFFFFF",
    angle_deg: float = 0.0,
    dual_origin: Optional[tuple] = None,
) -> str:
    """numpy + PIL 그라데이션 JPG 를 output_path 에 생성. 경로 반환.

    style:
      - "radial": 단일 라디얼
      - "linear": 리니어 (angle_deg 방향, 기본 0 = 좌→우)
      - "solid":  단색 (color_in 사용)
      - "dual_radial": 이중 라디얼 (origin + dual_origin)

    동기 함수이지만 async 시그니처를 유지해 라우터에서 await 할 수 있게 한다.
    실제 계산은 numpy 벡터 연산으로 짧다(<1초 / 1920x1080).
    """
    W, H = canvas_size
    c_in = _hex_to_np(color_in)
    c_out = _hex_to_np(color_out)

    if style == "solid":
        arr = np.zeros((H, W, 3), dtype=np.float32)
        arr[:, :, :] = c_in
    elif style == "linear":
        X, Y = _make_meshgrid(W, H)
        arr = _linear(W, H, X, Y, origin[0], origin[1], angle_deg, intensity, gamma, c_in, c_out)
    elif style == "dual_radial":
        X, Y = _make_meshgrid(W, H)
        a1 = _radial(W, H, X, Y, origin[0], origin[1], intensity, gamma, c_in, c_out)
        if dual_origin is None:
            dual_origin = (1.0 - origin[0], 1.0 - origin[1])
        a2 = _radial(W, H, X, Y, dual_origin[0], dual_origin[1], intensity, gamma, c_in, c_out)
        # 두 라디얼을 min 합성 (더 가까운 쪽이 c_in 으로 우세)
        arr = np.minimum(a1, a2)
    else:  # "radial"
        X, Y = _make_meshgrid(W, H)
        arr = _radial(W, H, X, Y, origin[0], origin[1], intensity, gamma, c_in, c_out)

    # 디렉토리 보장
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    _save_jpg(arr, output_path)
    return output_path


# ============ 슬라이드 인덱스 → 그라데이션 원점 매핑 (.md 4.3) ============

# 본문 8방위 원점 (인덱스 mod 8)
_BODY_ORIGINS = [
    (0.10, 0.10),  # 좌상
    (0.50, 0.10),  # 상중
    (0.90, 0.10),  # 우상
    (0.10, 0.50),  # 좌중
    (0.90, 0.50),  # 우중
    (0.10, 0.90),  # 좌하
    (0.50, 0.90),  # 하중
    (0.90, 0.90),  # 우하
]


def slide_background_origin(slide_index: int, template_id: str) -> dict:
    """슬라이드 인덱스 + 패턴(template_id)에 따라 그라데이션 파라미터를 결정.

    반환 dict (generate_background() kwargs 호환):
      - style, origin, intensity, gamma, color_in, color_out, angle_deg, dual_origin
    실제 색 hex 는 호출 측에서 design_tokens 로 치환할 수 있도록 키 토큰
    ("primary"/"darker"/"light"/"white") 형태로 'color_in_token'/'color_out_token'
    필드를 함께 동봉한다. 호출 측은 design_tokens.colors[token] 로 lookup 한다.
    """
    tid = (template_id or "").strip()

    if tid == "cover":
        return {
            "style": "linear",
            "origin": (0.30, 0.45),
            "intensity": 1.4,
            "gamma": 1.2,
            "angle_deg": 35.0,
            "color_in_token": "primary",
            "color_out_token": "darker",
            "dual_origin": None,
        }
    if tid == "closing":
        return {
            "style": "linear",
            "origin": (0.85, 0.55),
            "intensity": 1.4,
            "gamma": 1.2,
            "angle_deg": 215.0,
            "color_in_token": "primary",
            "color_out_token": "darker",
            "dual_origin": None,
        }
    if tid == "toc":
        return {
            "style": "dual_radial",
            "origin": (0.95, 0.05),
            "dual_origin": (0.0, 1.0),
            "intensity": 0.85,
            "gamma": 1.3,
            "color_in_token": "light",
            "color_out_token": "white",
            "angle_deg": 0.0,
        }
    if tid == "chapter":
        # 좌→우 리니어 LIGHT → WHITE
        return {
            "style": "linear",
            "origin": (0.0, 0.5),
            "intensity": 1.2,
            "gamma": 1.2,
            "angle_deg": 0.0,
            "color_in_token": "light",
            "color_out_token": "white",
            "dual_origin": None,
        }

    # 본문: 8방위 원점 인덱스
    body_idx = max(0, slide_index - 1) % 8
    ox, oy = _BODY_ORIGINS[body_idx]
    return {
        "style": "radial",
        "origin": (ox, oy),
        "intensity": 0.85,
        "gamma": 1.3,
        "color_in_token": "light",
        "color_out_token": "white",
        "angle_deg": 0.0,
        "dual_origin": None,
    }


# ============ 간이 아이콘 PNG (M5 에서 정식 SVG 세트로 교체 예정) ============

# 30개 키 (M2 시스템 프롬프트와 동일)
_ICON_GLYPH = {
    "users":     "U",
    "atom":      "A",
    "ship":      "S",
    "bomb":      "B",
    "shield":    "D",
    "fire":      "F",
    "chart":     "C",
    "star":      "*",
    "lightning": "Z",
    "target":    "O",
    "brain":     "B",
    "gear":      "G",
    "lock":      "L",
    "leaf":      "Y",
    "globe":     "G",
    "rocket":    "R",
    "eye":       "E",
    "check":     "V",
    "warning":   "!",
    "idea":      "I",
    "clock":     "T",
    "money":     "$",
    "building":  "H",
    "network":   "N",
    "book":      "K",
    "cloud":     "C",
    "code":      "{",
    "database":  "D",
    "mobile":    "M",
    "palette":   "P",
}


def _try_load_font(size: int) -> ImageFont.FreeTypeFont:
    """폰트 로드 - 시스템 기본 → 실패 시 PIL default"""
    candidates = [
        "C:/Windows/Fonts/malgun.ttf",     # Windows Malgun Gothic
        "C:/Windows/Fonts/arialbd.ttf",    # Arial Bold
        "C:/Windows/Fonts/arial.ttf",      # Arial
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ]
    for p in candidates:
        try:
            if os.path.exists(p):
                return ImageFont.truetype(p, size)
        except Exception:
            continue
    try:
        return ImageFont.load_default()
    except Exception:
        return None  # type: ignore


def icon_placeholder(icon_key: str, color_hex: str, size: int = 256) -> bytes:
    """간이 아이콘 PNG bytes 반환.

    외부 라이브러리 없이 PIL 만으로 '원 + 한 글자' 형태의 placeholder 를 생성한다.
    M5 에서 정식 SVG 아이콘 세트로 교체할 때 함수 시그니처를 유지하면 호출부 무수정.

    icon_key:
      - _ICON_GLYPH 에 정의된 30개 키 → 매핑된 한 글자
      - 빈 문자열 또는 미정의 → 빈 원만 그림
    color_hex: 아이콘 색상 (원 + 글자 동일 색)
    size: 출력 PNG 한 변 px
    """
    img = Image.new("RGBA", (size, size), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    rgb = _hex_to_rgb_tuple(color_hex)
    fill = rgb + (255,)

    # 원 (배경 fill 으로 채운 솔리드 원)
    pad = int(size * 0.06)
    draw.ellipse((pad, pad, size - pad, size - pad), fill=fill)

    # 글자
    glyph = _ICON_GLYPH.get((icon_key or "").strip())
    if glyph:
        font = _try_load_font(int(size * 0.55))
        if font is not None:
            # 텍스트 박스 측정
            try:
                bbox = draw.textbbox((0, 0), glyph, font=font)
                tw = bbox[2] - bbox[0]
                th = bbox[3] - bbox[1]
                tx = (size - tw) // 2 - bbox[0]
                ty = (size - th) // 2 - bbox[1]
            except Exception:
                tw, th = draw.textsize(glyph, font=font)  # type: ignore
                tx = (size - tw) // 2
                ty = (size - th) // 2
            draw.text((tx, ty), glyph, fill=(255, 255, 255, 255), font=font)

    buf = io.BytesIO()
    img.save(buf, "PNG", optimize=True)
    return buf.getvalue()


__all__ = [
    "generate_background",
    "slide_background_origin",
    "icon_placeholder",
]
