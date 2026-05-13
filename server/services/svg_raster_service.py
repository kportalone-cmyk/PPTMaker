"""SVG → PNG 라스터라이저 (외부 의존성 없음).

Heroicons (single-path solid icons) 같은 단순 SVG 를 PIL 만으로 PNG 로 변환한다.
지원 path 커맨드: M/m, L/l, H/h, V/v, C/c, S/s, Q/q, T/t, A/a, Z/z.

전략:
  1. lxml 로 SVG 파싱 → 모든 <path> 의 'd' 와 viewBox 추출
  2. d 를 토큰화 → 베지어/아크를 짧은 선분 시퀀스로 flatten (sub-path 단위)
  3. PIL 의 ImageDraw.polygon 으로 fill-rule=evenodd 처리한 다음 RGBA PNG 출력

이 모듈은 ppt_asset_service.icon_placeholder() 에서 사용된다.
"""

from __future__ import annotations

import io
import math
import os
import re
from typing import List, Tuple

from lxml import etree
from PIL import Image, ImageDraw


# 베지어 곡선 1개를 몇 개의 선분으로 쪼갤지 (해상도 기준)
_BEZIER_STEPS = 16


# ============ path 토큰화 ============

_PATH_TOKEN_RE = re.compile(r"([MmLlHhVvCcSsQqTtAaZz])|(-?\d*\.?\d+(?:[eE][-+]?\d+)?)")


def _tokenize_d(d: str):
    """SVG path 'd' 문자열 → (command, [floats]) 시퀀스 yield."""
    cmd = None
    nums: List[float] = []

    def flush():
        nonlocal cmd, nums
        if cmd is not None:
            yield_args = (cmd, nums)
        else:
            yield_args = None
        cmd = None
        nums = []
        return yield_args

    last = None
    for m in _PATH_TOKEN_RE.finditer(d):
        cmd_tok, num_tok = m.group(1), m.group(2)
        if cmd_tok is not None:
            if last is not None:
                yield last
            last = (cmd_tok, [])
        elif num_tok is not None:
            if last is None:
                # number before any command — assume implicit M
                last = ("M", [])
            last[1].append(float(num_tok))
    if last is not None:
        yield last


# ============ 베지어 flatten ============

def _quad_point(t: float, p0, p1, p2):
    u = 1.0 - t
    x = u * u * p0[0] + 2 * u * t * p1[0] + t * t * p2[0]
    y = u * u * p0[1] + 2 * u * t * p1[1] + t * t * p2[1]
    return x, y


def _cubic_point(t: float, p0, p1, p2, p3):
    u = 1.0 - t
    x = (u ** 3) * p0[0] + 3 * (u ** 2) * t * p1[0] + 3 * u * (t ** 2) * p2[0] + (t ** 3) * p3[0]
    y = (u ** 3) * p0[1] + 3 * (u ** 2) * t * p1[1] + 3 * u * (t ** 2) * p2[1] + (t ** 3) * p3[1]
    return x, y


def _arc_to_cubics(x1, y1, rx, ry, phi, large_arc, sweep, x2, y2):
    """SVG arc → 베지어 큐빅 근사 시퀀스 (W3C SVG 1.1 부록 F.6 알고리즘)."""
    if x1 == x2 and y1 == y2:
        return []
    rx = abs(rx)
    ry = abs(ry)
    if rx == 0 or ry == 0:
        return [((x1, y1), (x1, y1), (x2, y2), (x2, y2))]

    cos_phi = math.cos(math.radians(phi))
    sin_phi = math.sin(math.radians(phi))

    # Step 1: 시작/끝점 회전
    dx = (x1 - x2) / 2.0
    dy = (y1 - y2) / 2.0
    x1p = cos_phi * dx + sin_phi * dy
    y1p = -sin_phi * dx + cos_phi * dy

    # 보정: rx,ry 가 너무 작으면 확대
    lam = (x1p ** 2) / (rx ** 2) + (y1p ** 2) / (ry ** 2)
    if lam > 1.0:
        s = math.sqrt(lam)
        rx *= s
        ry *= s

    # Step 2: 중심점
    num = (rx ** 2) * (ry ** 2) - (rx ** 2) * (y1p ** 2) - (ry ** 2) * (x1p ** 2)
    den = (rx ** 2) * (y1p ** 2) + (ry ** 2) * (x1p ** 2)
    if den == 0:
        coef = 0.0
    else:
        coef = math.sqrt(max(0.0, num / den))
    if large_arc == sweep:
        coef = -coef
    cxp = coef * (rx * y1p / ry)
    cyp = -coef * (ry * x1p / rx)

    cx = cos_phi * cxp - sin_phi * cyp + (x1 + x2) / 2.0
    cy = sin_phi * cxp + cos_phi * cyp + (y1 + y2) / 2.0

    # Step 3: 시작 각도 / sweep 각도
    def _angle(ux, uy, vx, vy):
        dot = ux * vx + uy * vy
        mag = math.sqrt((ux ** 2 + uy ** 2) * (vx ** 2 + vy ** 2))
        if mag == 0:
            return 0.0
        c = max(-1.0, min(1.0, dot / mag))
        ang = math.acos(c)
        if ux * vy - uy * vx < 0:
            ang = -ang
        return ang

    theta1 = _angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dtheta = _angle((x1p - cxp) / rx, (y1p - cyp) / ry,
                    (-x1p - cxp) / rx, (-y1p - cyp) / ry)
    if not sweep and dtheta > 0:
        dtheta -= 2 * math.pi
    elif sweep and dtheta < 0:
        dtheta += 2 * math.pi

    # Step 4: 90도 이하 segment 들로 분할 후 각각 큐빅 베지어 근사
    segments = max(1, int(math.ceil(abs(dtheta) / (math.pi / 2.0))))
    delta = dtheta / segments
    t = (8.0 / 3.0) * math.sin(delta / 4.0) ** 2 / math.sin(delta / 2.0) if delta != 0 else 0.0

    beziers = []
    theta = theta1
    cos_t = math.cos(theta)
    sin_t = math.sin(theta)
    px = cos_phi * rx * cos_t - sin_phi * ry * sin_t + cx
    py = sin_phi * rx * cos_t + cos_phi * ry * sin_t + cy

    for _ in range(segments):
        theta_next = theta + delta
        cos_tn = math.cos(theta_next)
        sin_tn = math.sin(theta_next)
        # tangent at theta
        tx1 = -cos_phi * rx * sin_t - sin_phi * ry * cos_t
        ty1 = -sin_phi * rx * sin_t + cos_phi * ry * cos_t
        c1x = px + t * tx1
        c1y = py + t * ty1
        # endpoint at theta_next
        nx = cos_phi * rx * cos_tn - sin_phi * ry * sin_tn + cx
        ny = sin_phi * rx * cos_tn + cos_phi * ry * sin_tn + cy
        # tangent at theta_next
        tx2 = -cos_phi * rx * sin_tn - sin_phi * ry * cos_tn
        ty2 = -sin_phi * rx * sin_tn + cos_phi * ry * cos_tn
        c2x = nx - t * tx2
        c2y = ny - t * ty2
        beziers.append(((px, py), (c1x, c1y), (c2x, c2y), (nx, ny)))
        px, py = nx, ny
        theta = theta_next
        cos_t = cos_tn
        sin_t = sin_tn

    return beziers


# ============ path 평탄화 ============

def _flatten_path(d: str) -> List[List[Tuple[float, float]]]:
    """SVG path 'd' → sub-path 별 점 리스트 (각 sub-path 는 폐쇄형 폴리곤)."""
    subpaths: List[List[Tuple[float, float]]] = []
    current: List[Tuple[float, float]] = []
    cx, cy = 0.0, 0.0
    start_x, start_y = 0.0, 0.0
    last_cubic_ctrl = None  # 마지막 큐빅 컨트롤 (S 명령용)
    last_quad_ctrl = None

    def _add_pt(x, y):
        nonlocal current, cx, cy
        current.append((x, y))
        cx, cy = x, y

    for cmd, args in _tokenize_d(d):
        upper = cmd.upper()
        relative = cmd.islower()

        # M 첫번째는 moveto, 이후는 lineto 로 처리 (SVG 표준)
        i = 0
        if upper == "M":
            # 새 sub-path 시작
            if current:
                subpaths.append(current)
                current = []
            if len(args) < 2:
                continue
            x, y = args[0], args[1]
            if relative:
                x += cx
                y += cy
            start_x, start_y = x, y
            _add_pt(x, y)
            last_cubic_ctrl = None
            last_quad_ctrl = None
            # 추가 좌표쌍은 lineto 처럼 처리
            i = 2
            while i + 1 < len(args):
                x2, y2 = args[i], args[i + 1]
                if relative:
                    x2 += cx
                    y2 += cy
                _add_pt(x2, y2)
                i += 2
        elif upper == "L":
            while i + 1 < len(args):
                x, y = args[i], args[i + 1]
                if relative:
                    x += cx
                    y += cy
                _add_pt(x, y)
                i += 2
            last_cubic_ctrl = None
            last_quad_ctrl = None
        elif upper == "H":
            while i < len(args):
                x = args[i]
                if relative:
                    x += cx
                _add_pt(x, cy)
                i += 1
            last_cubic_ctrl = None
            last_quad_ctrl = None
        elif upper == "V":
            while i < len(args):
                y = args[i]
                if relative:
                    y += cy
                _add_pt(cx, y)
                i += 1
            last_cubic_ctrl = None
            last_quad_ctrl = None
        elif upper == "C":
            while i + 5 < len(args):
                x1, y1, x2, y2, x, y = args[i:i + 6]
                if relative:
                    x1 += cx; y1 += cy
                    x2 += cx; y2 += cy
                    x += cx; y += cy
                p0 = (cx, cy)
                for s in range(1, _BEZIER_STEPS + 1):
                    t = s / _BEZIER_STEPS
                    pt = _cubic_point(t, p0, (x1, y1), (x2, y2), (x, y))
                    current.append(pt)
                cx, cy = x, y
                last_cubic_ctrl = (x2, y2)
                last_quad_ctrl = None
                i += 6
        elif upper == "S":
            while i + 3 < len(args):
                x2, y2, x, y = args[i:i + 4]
                if relative:
                    x2 += cx; y2 += cy
                    x += cx; y += cy
                # 이전 cubic 의 두번째 컨트롤점 반사
                if last_cubic_ctrl is not None:
                    x1 = 2 * cx - last_cubic_ctrl[0]
                    y1 = 2 * cy - last_cubic_ctrl[1]
                else:
                    x1, y1 = cx, cy
                p0 = (cx, cy)
                for s in range(1, _BEZIER_STEPS + 1):
                    t = s / _BEZIER_STEPS
                    pt = _cubic_point(t, p0, (x1, y1), (x2, y2), (x, y))
                    current.append(pt)
                cx, cy = x, y
                last_cubic_ctrl = (x2, y2)
                last_quad_ctrl = None
                i += 4
        elif upper == "Q":
            while i + 3 < len(args):
                x1, y1, x, y = args[i:i + 4]
                if relative:
                    x1 += cx; y1 += cy
                    x += cx; y += cy
                p0 = (cx, cy)
                for s in range(1, _BEZIER_STEPS + 1):
                    t = s / _BEZIER_STEPS
                    pt = _quad_point(t, p0, (x1, y1), (x, y))
                    current.append(pt)
                cx, cy = x, y
                last_quad_ctrl = (x1, y1)
                last_cubic_ctrl = None
                i += 4
        elif upper == "T":
            while i + 1 < len(args):
                x, y = args[i:i + 2]
                if relative:
                    x += cx; y += cy
                if last_quad_ctrl is not None:
                    x1 = 2 * cx - last_quad_ctrl[0]
                    y1 = 2 * cy - last_quad_ctrl[1]
                else:
                    x1, y1 = cx, cy
                p0 = (cx, cy)
                for s in range(1, _BEZIER_STEPS + 1):
                    t = s / _BEZIER_STEPS
                    pt = _quad_point(t, p0, (x1, y1), (x, y))
                    current.append(pt)
                cx, cy = x, y
                last_quad_ctrl = (x1, y1)
                last_cubic_ctrl = None
                i += 2
        elif upper == "A":
            while i + 6 < len(args):
                rx, ry, phi, large_arc, sweep, x, y = args[i:i + 7]
                if relative:
                    x += cx; y += cy
                beziers = _arc_to_cubics(cx, cy, rx, ry, phi,
                                          int(large_arc) == 1, int(sweep) == 1, x, y)
                for (p0, p1, p2, p3) in beziers:
                    for s in range(1, _BEZIER_STEPS + 1):
                        t = s / _BEZIER_STEPS
                        pt = _cubic_point(t, p0, p1, p2, p3)
                        current.append(pt)
                cx, cy = x, y
                last_cubic_ctrl = None
                last_quad_ctrl = None
                i += 7
        elif upper == "Z":
            # sub-path 닫기
            if current:
                if (cx, cy) != (start_x, start_y):
                    current.append((start_x, start_y))
                subpaths.append(current)
            current = []
            cx, cy = start_x, start_y
            last_cubic_ctrl = None
            last_quad_ctrl = None

    if current:
        subpaths.append(current)

    return subpaths


# ============ public: SVG → RGBA PIL.Image ============

def _parse_viewbox(svg_root) -> Tuple[float, float, float, float]:
    vb = svg_root.get("viewBox")
    if vb:
        parts = re.split(r"[\s,]+", vb.strip())
        if len(parts) == 4:
            try:
                return tuple(float(p) for p in parts)  # type: ignore
            except ValueError:
                pass
    w = float(svg_root.get("width", "24") or 24)
    h = float(svg_root.get("height", "24") or 24)
    return (0.0, 0.0, w, h)


def render_svg_to_png(svg_bytes: bytes, output_size: int = 256,
                      color_rgb: Tuple[int, int, int] = (0, 0, 0)) -> Image.Image:
    """SVG bytes → RGBA Image (output_size × output_size).

    모든 <path> 의 fill 영역을 color_rgb 로 채워 단일 색 아이콘으로 라스터라이즈.
    """
    try:
        # SVG 가 default ns 를 가짐 → tag 매칭이 까다로움. namespace 스트립.
        parser = etree.XMLParser(remove_comments=True)
        root = etree.fromstring(svg_bytes, parser=parser)
    except Exception:
        return Image.new("RGBA", (output_size, output_size), (0, 0, 0, 0))

    vbx, vby, vbw, vbh = _parse_viewbox(root)
    if vbw <= 0 or vbh <= 0:
        vbw, vbh = 24.0, 24.0

    img = Image.new("RGBA", (output_size, output_size), (0, 0, 0, 0))
    # 누적 마스크 — 여러 <path> 의 fill 영역을 OR 합집합으로 누적
    accum_mask = Image.new("L", (output_size, output_size), 0)

    scale = output_size / max(vbw, vbh)
    # viewBox 가 정사각 아닐 때 중앙 정렬
    ox = (output_size - vbw * scale) / 2.0
    oy = (output_size - vbh * scale) / 2.0

    # path 요소 모두 추출 (네임스페이스 무시)
    for elem in root.iter():
        tag = etree.QName(elem.tag).localname
        if tag != "path":
            continue
        d = elem.get("d")
        if not d:
            continue
        # fill="none" 은 그리지 않음
        fill_attr = (elem.get("fill") or "").strip().lower()
        if fill_attr == "none":
            continue
        subpaths = _flatten_path(d)
        fill_rule = (elem.get("fill-rule") or "nonzero").lower()

        # viewBox → output 좌표 변환
        transformed = []
        for sp in subpaths:
            if len(sp) < 3:
                continue
            pts = [((x - vbx) * scale + ox, (y - vby) * scale + oy) for (x, y) in sp]
            transformed.append(pts)

        if not transformed:
            continue

        # 이 <path> 의 fill 영역을 path_mask 에 그림
        if fill_rule == "evenodd":
            # 각 sub-path 를 XOR 누적 — 짝수 겹침 = 비움, 홀수 겹침 = 채움
            path_mask = Image.new("L", (output_size, output_size), 0)
            for pts in transformed:
                m2 = Image.new("L", (output_size, output_size), 0)
                ImageDraw.Draw(m2).polygon(pts, fill=255)
                path_mask = _xor_mask(path_mask, m2)
        else:
            # nonzero — 모든 sub-path 의 OR 합집합 (방향 기반 winding 은 미구현 근사)
            path_mask = Image.new("L", (output_size, output_size), 0)
            pd = ImageDraw.Draw(path_mask)
            for pts in transformed:
                pd.polygon(pts, fill=255)

        # 누적 마스크에 OR 합성
        accum_mask = _or_mask(accum_mask, path_mask)

    # 누적 마스크 → color overlay
    color_layer = Image.new("RGBA", (output_size, output_size), color_rgb + (255,))
    img.paste(color_layer, (0, 0), mask=accum_mask)
    return img


def _xor_mask(a: Image.Image, b: Image.Image) -> Image.Image:
    """L mode 두 마스크의 XOR 합성 (even-odd fill)."""
    import numpy as np
    arr_a = np.array(a, dtype=np.uint8)
    arr_b = np.array(b, dtype=np.uint8)
    # 둘 다 켜져 있으면 꺼짐, 한쪽만 켜져 있으면 켜짐
    out = ((arr_a > 0) ^ (arr_b > 0)).astype(np.uint8) * 255
    return Image.fromarray(out, mode="L")


def _or_mask(a: Image.Image, b: Image.Image) -> Image.Image:
    import numpy as np
    arr_a = np.array(a, dtype=np.uint8)
    arr_b = np.array(b, dtype=np.uint8)
    out = ((arr_a > 0) | (arr_b > 0)).astype(np.uint8) * 255
    return Image.fromarray(out, mode="L")


def svg_file_to_png_bytes(svg_path: str, size: int = 256,
                           color_rgb: Tuple[int, int, int] = (0, 0, 0)) -> bytes:
    """SVG 파일을 읽어 PNG bytes 로 반환."""
    if not os.path.isfile(svg_path):
        return b""
    with open(svg_path, "rb") as f:
        data = f.read()
    img = render_svg_to_png(data, output_size=size, color_rgb=color_rgb)
    buf = io.BytesIO()
    img.save(buf, "PNG", optimize=True)
    return buf.getvalue()


__all__ = ["render_svg_to_png", "svg_file_to_png_bytes"]
