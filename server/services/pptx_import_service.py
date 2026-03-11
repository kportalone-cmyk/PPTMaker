"""PPTX 파일을 파싱하여 템플릿 + 슬라이드를 MongoDB에 자동 생성하는 서비스"""

import os
import uuid
import logging
from datetime import datetime

from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

from services.mongo_service import get_db
from config import settings

logger = logging.getLogger(__name__)

# ============ 캔버스 / PPTX 상수 (ppt_service.py와 동일) ============

SLIDE_W_PX = 960
SLIDE_H_PX = 540
SLIDE_W_EMU = 12192000
SLIDE_H_EMU = 6858000

# ============ 역방향 도형 매핑 (MSO_SHAPE enum → 문자열) ============

_FORWARD_SHAPE_MAP = {
    "rectangle": MSO_SHAPE.RECTANGLE,
    "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
    "snip_1_rect": MSO_SHAPE.SNIP_1_RECTANGLE,
    "snip_2_diag_rect": MSO_SHAPE.SNIP_2_DIAG_RECTANGLE,
    "round_1_rect": MSO_SHAPE.ROUND_1_RECTANGLE,
    "round_2_diag_rect": MSO_SHAPE.ROUND_2_DIAG_RECTANGLE,
    "ellipse": MSO_SHAPE.OVAL,
    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "right_triangle": MSO_SHAPE.RIGHT_TRIANGLE,
    "parallelogram": MSO_SHAPE.PARALLELOGRAM,
    "trapezoid": MSO_SHAPE.TRAPEZOID,
    "diamond": MSO_SHAPE.DIAMOND,
    "pentagon": MSO_SHAPE.REGULAR_PENTAGON,
    "hexagon": MSO_SHAPE.HEXAGON,
    "heptagon": MSO_SHAPE.HEPTAGON,
    "octagon": MSO_SHAPE.OCTAGON,
    "decagon": MSO_SHAPE.DECAGON,
    "dodecagon": MSO_SHAPE.DODECAGON,
    "cross": MSO_SHAPE.CROSS,
    "donut": MSO_SHAPE.DONUT,
    "no_smoking": MSO_SHAPE.NO_SYMBOL,
    "block_arc": MSO_SHAPE.BLOCK_ARC,
    "heart": MSO_SHAPE.HEART,
    "lightning_bolt": MSO_SHAPE.LIGHTNING_BOLT,
    "sun": MSO_SHAPE.SUN,
    "moon": MSO_SHAPE.MOON,
    "cloud": MSO_SHAPE.CLOUD,
    "smiley_face": MSO_SHAPE.SMILEY_FACE,
    "folded_corner": MSO_SHAPE.FOLDED_CORNER,
    "frame": MSO_SHAPE.FRAME,
    "teardrop": MSO_SHAPE.TEAR,
    "plaque": MSO_SHAPE.PLAQUE,
    "brace_pair": MSO_SHAPE.DOUBLE_BRACE,
    "bracket_pair": MSO_SHAPE.DOUBLE_BRACKET,
    "right_arrow": MSO_SHAPE.RIGHT_ARROW,
    "left_arrow": MSO_SHAPE.LEFT_ARROW,
    "up_arrow": MSO_SHAPE.UP_ARROW,
    "down_arrow": MSO_SHAPE.DOWN_ARROW,
    "left_right_arrow": MSO_SHAPE.LEFT_RIGHT_ARROW,
    "up_down_arrow": MSO_SHAPE.UP_DOWN_ARROW,
    "quad_arrow": MSO_SHAPE.QUAD_ARROW,
    "notched_right_arrow": MSO_SHAPE.NOTCHED_RIGHT_ARROW,
    "chevron": MSO_SHAPE.CHEVRON,
    "home_plate": MSO_SHAPE.PENTAGON,
    "striped_right_arrow": MSO_SHAPE.STRIPED_RIGHT_ARROW,
    "bent_arrow": MSO_SHAPE.BENT_ARROW,
    "u_turn_arrow": MSO_SHAPE.U_TURN_ARROW,
    "circular_arrow": MSO_SHAPE.CIRCULAR_ARROW,
    "math_plus": MSO_SHAPE.MATH_PLUS,
    "math_minus": MSO_SHAPE.MATH_MINUS,
    "math_multiply": MSO_SHAPE.MATH_MULTIPLY,
    "math_divide": MSO_SHAPE.MATH_DIVIDE,
    "math_equal": MSO_SHAPE.MATH_EQUAL,
    "math_not_equal": MSO_SHAPE.MATH_NOT_EQUAL,
    "star_4_point": MSO_SHAPE.STAR_4_POINT,
    "star_5_point": MSO_SHAPE.STAR_5_POINT,
    "star_6_point": MSO_SHAPE.STAR_6_POINT,
    "star_8_point": MSO_SHAPE.STAR_8_POINT,
    "star_10_point": MSO_SHAPE.STAR_10_POINT,
    "star_12_point": MSO_SHAPE.STAR_12_POINT,
    "star_16_point": MSO_SHAPE.STAR_16_POINT,
    "star_24_point": MSO_SHAPE.STAR_24_POINT,
    "star_32_point": MSO_SHAPE.STAR_32_POINT,
    "explosion_1": MSO_SHAPE.EXPLOSION1,
    "explosion_2": MSO_SHAPE.EXPLOSION2,
    "wave": MSO_SHAPE.WAVE,
    "double_wave": MSO_SHAPE.DOUBLE_WAVE,
    "ribbon": MSO_SHAPE.DOWN_RIBBON,
    "wedge_rect_callout": MSO_SHAPE.RECTANGULAR_CALLOUT,
    "wedge_round_rect_callout": MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT,
    "wedge_ellipse_callout": MSO_SHAPE.OVAL_CALLOUT,
    "cloud_callout": MSO_SHAPE.CLOUD_CALLOUT,
    "border_callout_1": MSO_SHAPE.LINE_CALLOUT_1,
    "border_callout_2": MSO_SHAPE.LINE_CALLOUT_2,
    "border_callout_3": MSO_SHAPE.LINE_CALLOUT_3,
}

# MSO_SHAPE enum value → 문자열 이름 (역방향 매핑)
_REVERSE_SHAPE_MAP = {v: k for k, v in _FORWARD_SHAPE_MAP.items()}

# ============ TOC / Closing 키워드 ============

_TOC_KEYWORDS = {"목차", "contents", "table of contents", "index"}
_CLOSING_KEYWORDS = {"감사", "thank", "q&a", "질의", "문의"}


# ============ 유틸리티 함수 ============


def _emu_to_px(emu_val, is_horizontal=True):
    """EMU 값을 캔버스 px 좌표로 변환 (ppt_service.py의 역연산)"""
    if is_horizontal:
        return round(emu_val / SLIDE_W_EMU * SLIDE_W_PX, 1)
    return round(emu_val / SLIDE_H_EMU * SLIDE_H_PX, 1)


def _gen_obj_id():
    """admin.js 호환 오브젝트 ID 생성"""
    return "obj_" + uuid.uuid4().hex[:12]


def _safe_rgb(color_obj):
    """안전하게 색상 hex 문자열 추출 (테마 색상 등 예외 처리)"""
    try:
        if color_obj is not None and color_obj.rgb is not None:
            return f"#{color_obj.rgb}"
    except (AttributeError, TypeError, ValueError):
        pass
    return "#000000"


def _safe_font_color(font):
    """폰트에서 색상을 안전하게 추출 (color 속성 접근 자체가 예외를 던질 수 있음)"""
    try:
        color_obj = font.color
        return _safe_rgb(color_obj)
    except (AttributeError, TypeError, ValueError):
        return "#000000"


def _alignment_to_str(alignment):
    """PP_ALIGN enum → 문자열 변환"""
    if alignment == PP_ALIGN.CENTER:
        return "center"
    elif alignment == PP_ALIGN.RIGHT:
        return "right"
    return "left"


def _get_image_extension(content_type: str) -> str:
    """이미지 content_type → 파일 확장자 매핑"""
    type_map = {
        "image/png": "png",
        "image/jpeg": "jpg",
        "image/jpg": "jpg",
        "image/gif": "gif",
        "image/bmp": "bmp",
        "image/tiff": "tiff",
        "image/webp": "webp",
        "image/svg+xml": "svg",
        "image/x-emf": "emf",
        "image/x-wmf": "wmf",
    }
    return type_map.get(content_type, "png")


def _shape_covers_slide(left, top, width, height):
    """도형이 슬라이드 면적의 90% 이상을 차지하는지 확인 (EMU 값 기준)"""
    shape_area = width * height
    slide_area = SLIDE_W_EMU * SLIDE_H_EMU
    if slide_area == 0:
        return False
    coverage = shape_area / slide_area

    # 위치도 대략 원점 근처여야 함
    near_origin = (abs(left) < SLIDE_W_EMU * 0.05) and (abs(top) < SLIDE_H_EMU * 0.05)

    return coverage >= 0.9 and near_origin


def _save_image_blob(blob: bytes, content_type: str, subfolder: str = "images") -> str:
    """이미지 바이트를 디스크에 저장하고 URL 경로를 반환"""
    ext = _get_image_extension(content_type)
    filename = f"{uuid.uuid4().hex}.{ext}"
    save_dir = os.path.join(settings.UPLOAD_DIR, subfolder)
    os.makedirs(save_dir, exist_ok=True)
    save_path = os.path.join(save_dir, filename)

    with open(save_path, "wb") as f:
        f.write(blob)

    return f"/uploads/{subfolder}/{filename}"


# ============ 텍스트 스타일 추출 ============


def _extract_text_style(paragraph):
    """첫 번째 run에서 텍스트 스타일 추출 (기본값 적용)"""
    style = {
        "font_family": "Arial",
        "font_size": 16,
        "color": "#000000",
        "bold": False,
        "italic": False,
        "align": _alignment_to_str(paragraph.alignment),
    }

    if not paragraph.runs:
        return style

    run = paragraph.runs[0]
    font = run.font

    try:
        if font.name:
            style["font_family"] = font.name
    except (AttributeError, TypeError):
        pass

    try:
        if font.size is not None:
            # font.size는 Emu 단위 → Pt 변환: emu / 12700
            style["font_size"] = max(1, round(font.size / 12700))
    except (AttributeError, TypeError):
        pass

    style["color"] = _safe_font_color(font)

    try:
        style["bold"] = bool(font.bold)
    except (AttributeError, TypeError):
        pass

    try:
        style["italic"] = bool(font.italic)
    except (AttributeError, TypeError):
        pass

    return style


def _extract_text_content(text_frame):
    """텍스트 프레임의 모든 paragraph를 합쳐 텍스트 반환"""
    lines = []
    for para in text_frame.paragraphs:
        lines.append(para.text)
    return "\n".join(lines)


def _get_largest_font_size(text_frame):
    """텍스트 프레임에서 가장 큰 폰트 사이즈 (Pt)를 반환"""
    max_size = 0
    for para in text_frame.paragraphs:
        for run in para.runs:
            try:
                if run.font.size is not None:
                    pt_size = run.font.size / 12700
                    if pt_size > max_size:
                        max_size = pt_size
            except (AttributeError, TypeError):
                pass
    return max_size


# ============ 개별 도형 추출 ============


def _extract_text_shape(shape, z_index: int) -> dict:
    """텍스트 도형 오브젝트 추출"""
    tf = shape.text_frame
    text_content = _extract_text_content(tf)

    # 첫 번째 paragraph에서 스타일 추출
    first_para = tf.paragraphs[0] if tf.paragraphs else None
    text_style = _extract_text_style(first_para) if first_para else {
        "font_family": "Arial",
        "font_size": 16,
        "color": "#000000",
        "bold": False,
        "italic": False,
        "align": "left",
    }

    return {
        "obj_id": _gen_obj_id(),
        "obj_type": "text",
        "x": _emu_to_px(shape.left, is_horizontal=True),
        "y": _emu_to_px(shape.top, is_horizontal=False),
        "width": _emu_to_px(shape.width, is_horizontal=True),
        "height": _emu_to_px(shape.height, is_horizontal=False),
        "text_content": text_content,
        "text_style": text_style,
        "z_index": z_index,
        "role": None,       # 나중에 assign_roles에서 설정
        "placeholder": None,  # 나중에 assign_roles에서 설정
    }


def _extract_image_shape(shape, z_index: int) -> dict | None:
    """이미지 도형 오브젝트 추출 (이미지를 디스크에 저장)"""
    try:
        image = shape.image
        blob = image.blob
        content_type = image.content_type
        image_url = _save_image_blob(blob, content_type, subfolder="images")

        return {
            "obj_id": _gen_obj_id(),
            "obj_type": "image",
            "x": _emu_to_px(shape.left, is_horizontal=True),
            "y": _emu_to_px(shape.top, is_horizontal=False),
            "width": _emu_to_px(shape.width, is_horizontal=True),
            "height": _emu_to_px(shape.height, is_horizontal=False),
            "image_url": image_url,
            "image_fit": "contain",
            "z_index": z_index,
            "role": None,
            "placeholder": None,
        }
    except Exception as e:
        logger.warning(f"이미지 도형 추출 실패: {e}")
        return None


def _extract_auto_shape(shape, z_index: int) -> dict | None:
    """자동 도형 (AutoShape) 오브젝트 추출"""
    try:
        # MSO_SHAPE enum → 문자열 이름 변환
        shape_type_str = "rectangle"
        try:
            auto_shape_type = shape.auto_shape_type
            shape_type_str = _REVERSE_SHAPE_MAP.get(auto_shape_type, "rectangle")
        except (AttributeError, TypeError, ValueError):
            pass

        # 채우기 색상 추출
        fill_color = "#4A90D9"
        try:
            fill = shape.fill
            if fill.type is not None:
                fore_color = fill.fore_color
                fill_color = _safe_rgb(fore_color)
        except (AttributeError, TypeError, ValueError):
            pass

        # 테두리 색상 / 두께 추출
        stroke_color = "#333333"
        stroke_width = 2
        try:
            line = shape.line
            if line.color and line.color.rgb:
                stroke_color = f"#{line.color.rgb}"
        except (AttributeError, TypeError, ValueError):
            pass
        try:
            if shape.line.width is not None:
                stroke_width = max(0, round(shape.line.width / 12700))  # EMU → Pt
        except (AttributeError, TypeError, ValueError):
            pass

        obj = {
            "obj_id": _gen_obj_id(),
            "obj_type": "shape",
            "x": _emu_to_px(shape.left, is_horizontal=True),
            "y": _emu_to_px(shape.top, is_horizontal=False),
            "width": _emu_to_px(shape.width, is_horizontal=True),
            "height": _emu_to_px(shape.height, is_horizontal=False),
            "shape_style": {
                "shape_type": shape_type_str,
                "fill_color": fill_color,
                "fill_opacity": 1.0,
                "stroke_color": stroke_color,
                "stroke_width": stroke_width,
                "stroke_dash": "solid",
            },
            "z_index": z_index,
            "role": None,
            "placeholder": None,
        }

        # 텍스트가 포함된 도형 → 텍스트 오브젝트로 변환
        if shape.has_text_frame:
            text_content = _extract_text_content(shape.text_frame)
            if text_content.strip():
                first_para = shape.text_frame.paragraphs[0] if shape.text_frame.paragraphs else None
                text_style = _extract_text_style(first_para) if first_para else {
                    "font_family": "Arial", "font_size": 16, "color": "#000000",
                    "bold": False, "italic": False, "align": "left",
                }
                return {
                    "obj_id": obj["obj_id"],
                    "obj_type": "text",
                    "x": obj["x"],
                    "y": obj["y"],
                    "width": obj["width"],
                    "height": obj["height"],
                    "text_content": text_content,
                    "text_style": text_style,
                    "z_index": z_index,
                    "role": None,
                    "placeholder": None,
                }

        return obj
    except Exception as e:
        logger.warning(f"자동 도형 추출 실패: {e}")
        return None


def _extract_table_as_text(shape, z_index: int) -> dict | None:
    """테이블을 텍스트 오브젝트로 변환 (내용만 추출)"""
    try:
        table = shape.table
        rows_text = []
        for row in table.rows:
            cells_text = []
            for cell in row.cells:
                cells_text.append(cell.text.strip())
            rows_text.append(" | ".join(cells_text))

        text_content = "\n".join(rows_text)
        if not text_content.strip():
            return None

        return {
            "obj_id": _gen_obj_id(),
            "obj_type": "text",
            "x": _emu_to_px(shape.left, is_horizontal=True),
            "y": _emu_to_px(shape.top, is_horizontal=False),
            "width": _emu_to_px(shape.width, is_horizontal=True),
            "height": _emu_to_px(shape.height, is_horizontal=False),
            "text_content": text_content,
            "text_style": {
                "font_family": "Arial",
                "font_size": 12,
                "color": "#000000",
                "bold": False,
                "italic": False,
                "align": "left",
            },
            "z_index": z_index,
            "role": "description",
            "placeholder": None,
        }
    except Exception as e:
        logger.warning(f"테이블 추출 실패: {e}")
        return None


# ============ 배경 이미지 추출 ============


def _extract_slide_background(slide) -> str | None:
    """슬라이드 배경 이미지를 추출하여 저장, URL 반환"""
    # 방법 1: slide.background.fill에서 이미지 추출
    try:
        bg = slide.background
        fill = bg.fill
        # fill type이 picture인 경우
        if fill.type is not None:
            try:
                # XML에서 직접 blipFill 확인
                from pptx.oxml.ns import qn
                bg_elem = bg._element
                blip_fills = bg_elem.findall(f".//{qn('a:blip')}")
                if blip_fills:
                    for blip in blip_fills:
                        r_embed = blip.get(qn("r:embed"))
                        if r_embed:
                            part = slide.part.related_parts.get(r_embed)
                            if part:
                                blob = part.blob
                                content_type = part.content_type or "image/png"
                                return _save_image_blob(blob, content_type, subfolder="backgrounds")
            except Exception:
                pass
    except Exception:
        pass

    return None


def _check_first_shape_as_background(shapes, total_shapes: int) -> tuple[str | None, bool]:
    """첫 번째 도형이 슬라이드 전체를 덮는 이미지인지 확인
    Returns: (bg_url 또는 None, 배경으로 사용됨 여부)
    """
    if total_shapes == 0:
        return None, False

    first_shape = shapes[0]

    # 이미지인지 확인
    try:
        shape_type = first_shape.shape_type
        is_picture = (shape_type == MSO_SHAPE_TYPE.PICTURE)
    except (AttributeError, TypeError):
        is_picture = False

    if not is_picture:
        try:
            _ = first_shape.image
            is_picture = True
        except Exception:
            is_picture = False

    if not is_picture:
        return None, False

    # 슬라이드 면적의 90% 이상을 차지하는지 확인
    try:
        if _shape_covers_slide(first_shape.left, first_shape.top,
                                first_shape.width, first_shape.height):
            image = first_shape.image
            blob = image.blob
            content_type = image.content_type
            bg_url = _save_image_blob(blob, content_type, subfolder="backgrounds")
            return bg_url, True
    except Exception:
        pass

    return None, False


# ============ 역할 할당 (Role Assignment) ============


def _assign_roles(objects: list[dict], content_type: str = "body"):
    """텍스트 오브젝트에 role과 placeholder를 자동 할당

    슬라이드 유형별 역할 할당 전략:
    - cover/title_slide : 가장 큰 폰트 → 제목, 그 다음 → 부제목, 나머지 → 내용
    - toc              : 가장 큰 폰트 → 제목, 나머지 → 내용 (목차 항목)
    - section_divider  : 가장 큰 폰트 → 제목, 나머지 → 부제목
    - body             : 가장 큰 폰트 → 제목, 두 번째 큰 폰트 → 부제목, 나머지 → 내용
    - closing          : 가장 큰 폰트 → 제목, 나머지 → 부제목

    공통 규칙:
    - Y 위치 기준으로 공간 순서를 고려
    - 폰트 크기로 중요도 판단 (큰 폰트 = 제목급)
    - 상단의 아주 작은 텍스트(≤10pt, 상위 10%) → governance
    """
    text_objs = [o for o in objects if o["obj_type"] == "text"]

    if not text_objs:
        return

    # 1단계: Y 위치 순서로 정렬 (위→아래)
    text_by_pos = sorted(text_objs, key=lambda o: o.get("y", 0))

    # 2단계: 가장 큰 폰트를 가진 오브젝트 찾기 (제목 후보)
    largest_obj = max(text_objs, key=lambda o: o.get("text_style", {}).get("font_size", 0))
    largest_font = largest_obj.get("text_style", {}).get("font_size", 16)

    # 3단계: 역할 할당
    title_obj = None
    subtitle_count = 0
    desc_count = 0

    # -- 제목 결정: 가장 큰 폰트 오브젝트 --
    if largest_font >= 14:
        title_obj = largest_obj
        title_obj["role"] = "title"
        title_obj["placeholder"] = "title_1"

    # -- 나머지 오브젝트에 역할 할당 (Y 위치 순서대로) --
    for obj in text_by_pos:
        if obj is title_obj:
            continue
        if obj.get("role") is not None:
            continue

        font_size = obj.get("text_style", {}).get("font_size", 16)
        y_pos = obj.get("y", 0)

        # governance: 상단 10% 이내, 매우 작은 폰트(≤10pt)
        if font_size <= 10 and y_pos < SLIDE_H_PX * 0.10:
            obj["role"] = "governance"
            obj["placeholder"] = "governance_1"
            continue

        # 슬라이드 유형별 부제목 vs 내용 분류
        if content_type in ("title_slide", "cover"):
            # 커버: 제목 외 → 부제목 (최대 2개), 나머지 → 내용
            if subtitle_count < 2:
                subtitle_count += 1
                obj["role"] = "subtitle"
                obj["placeholder"] = f"subtitle_{subtitle_count}"
            else:
                desc_count += 1
                obj["role"] = "description"
                obj["placeholder"] = f"desc_{desc_count}"

        elif content_type == "section_divider":
            # 간지: 제목 외 모두 → 부제목
            subtitle_count += 1
            obj["role"] = "subtitle"
            obj["placeholder"] = f"subtitle_{subtitle_count}"

        elif content_type == "closing":
            # 마무리: 제목 외 → 부제목 (최대 2개), 나머지 → 내용
            if subtitle_count < 2:
                subtitle_count += 1
                obj["role"] = "subtitle"
                obj["placeholder"] = f"subtitle_{subtitle_count}"
            else:
                desc_count += 1
                obj["role"] = "description"
                obj["placeholder"] = f"desc_{desc_count}"

        elif content_type == "toc":
            # 목차: 제목 외 모두 → 내용 (목차 항목)
            desc_count += 1
            obj["role"] = "description"
            obj["placeholder"] = f"desc_{desc_count}"

        else:
            # body (본문): 제목 다음 큰 폰트 1개 → 부제목, 나머지 → 내용
            if subtitle_count == 0 and font_size >= largest_font * 0.6 and font_size >= 14:
                subtitle_count += 1
                obj["role"] = "subtitle"
                obj["placeholder"] = f"subtitle_{subtitle_count}"
            else:
                desc_count += 1
                obj["role"] = "description"
                obj["placeholder"] = f"desc_{desc_count}"

    # 제목이 미할당인 경우 (모든 폰트가 14pt 미만)
    if title_obj is None and text_by_pos:
        # 맨 위 텍스트를 제목으로 강제 지정
        first = text_by_pos[0]
        if first.get("role") == "description":
            desc_count -= 1
        elif first.get("role") == "subtitle":
            subtitle_count -= 1
        first["role"] = "title"
        first["placeholder"] = "title_1"

    # role이 아직 None인 텍스트 오브젝트 처리
    for obj in objects:
        if obj["obj_type"] == "text" and obj.get("role") is None:
            desc_count += 1
            obj["role"] = "description"
            obj["placeholder"] = f"desc_{desc_count}"


# ============ 슬라이드 분류 ============


def _classify_slide(slide_index: int, total_slides: int, objects: list[dict]) -> dict:
    """슬라이드를 content_type과 layout으로 분류

    Returns: {"content_type": str, "layout": str}
    """
    # 모든 텍스트 합치기 (분류 키워드 탐지용)
    all_text = ""
    text_objs = []
    max_font_size = 0

    for obj in objects:
        if obj["obj_type"] == "text":
            text_objs.append(obj)
            text_content = obj.get("text_content", "")
            all_text += " " + text_content
            font_size = obj.get("text_style", {}).get("font_size", 16)
            if font_size > max_font_size:
                max_font_size = font_size

    all_text_lower = all_text.lower().strip()

    # 1. 첫 번째 슬라이드 → title_slide / cover
    if slide_index == 0:
        return {"content_type": "title_slide", "layout": "cover"}

    # 2. 마지막 슬라이드 (총 3개 이상) → closing
    if slide_index == total_slides - 1 and total_slides > 2:
        return {"content_type": "closing", "layout": "closing"}

    # 3. 목차 키워드 확인
    for kw in _TOC_KEYWORDS:
        if kw in all_text_lower:
            return {"content_type": "toc", "layout": "numbered_list"}

    # 4. 마감 키워드 확인 (감사, Thank, Q&A 등)
    for kw in _CLOSING_KEYWORDS:
        if kw in all_text_lower:
            return {"content_type": "closing", "layout": "closing"}

    # 5. 텍스트 박스 적고 (1-2개) 큰 폰트 (>24pt) → section_divider
    if 1 <= len(text_objs) <= 2 and max_font_size > 24:
        return {"content_type": "section_divider", "layout": "divider"}

    # 6. 기본값 → body / single_column
    return {"content_type": "body", "layout": "single_column"}


def _build_slide_meta(classification: dict, objects: list[dict]) -> dict:
    """슬라이드 메타 정보 생성"""
    has_title = False
    has_governance = False
    description_count = 0

    for obj in objects:
        role = obj.get("role")
        if role == "title":
            has_title = True
        elif role == "governance":
            has_governance = True
        elif role == "description":
            description_count += 1

    return {
        "content_type": classification["content_type"],
        "layout": classification["layout"],
        "has_title": has_title,
        "has_governance": has_governance,
        "description_count": description_count,
    }


# ============ 슬라이드 파싱 ============


def _parse_slide(pptx_slide, slide_index: int, total_slides: int) -> dict:
    """하나의 PPTX 슬라이드를 파싱하여 오브젝트 리스트 + 분류 정보 반환"""
    objects = []
    z_index = 10

    # 배경 이미지 추출
    bg_url = _extract_slide_background(pptx_slide)

    # 도형 목록 가져오기
    shapes_list = list(pptx_slide.shapes)

    # 첫 번째 도형이 전체 슬라이드를 덮는 배경 이미지인지 확인
    first_shape_bg, is_first_bg = _check_first_shape_as_background(shapes_list, len(shapes_list))
    if first_shape_bg and not bg_url:
        bg_url = first_shape_bg

    # 도형 순회 (첫 번째가 배경이면 건너뜀)
    start_idx = 1 if is_first_bg else 0

    for shape in shapes_list[start_idx:]:
        try:
            obj = _extract_shape(shape, z_index)
            if obj is not None:
                objects.append(obj)
                z_index += 1
        except Exception as e:
            logger.warning(f"슬라이드 {slide_index + 1} 도형 추출 실패: {e}")
            continue

    # 슬라이드 분류 (역할 할당보다 먼저 수행)
    classification = _classify_slide(slide_index, total_slides, objects)

    # 텍스트 오브젝트에 role 할당 (분류 결과 활용)
    _assign_roles(objects, content_type=classification["content_type"])

    # 슬라이드 메타 정보
    slide_meta = _build_slide_meta(classification, objects)

    # 텍스트 오브젝트의 실제 내용을 플레이스홀더로 교체 (분류/역할 할당 이후)
    for obj in objects:
        if obj["obj_type"] == "text":
            obj["text_content"] = "텍스트를 입력하세요"
            # 제목은 최소 높이 60px, 부제목은 최소 높이 50px 보장
            role = obj.get("role")
            if role == "title" and obj.get("height", 0) < 60:
                obj["height"] = 60
            elif role == "subtitle" and obj.get("height", 0) < 50:
                obj["height"] = 50
            elif role == "description" and obj.get("height", 0) < 45:
                obj["height"] = 45

    return {
        "objects": objects,
        "classification": classification,
        "slide_meta": slide_meta,
        "background_image": bg_url,
    }


def _extract_shape(shape, z_index: int) -> dict | None:
    """도형 타입을 판별하여 적절한 추출 함수 호출"""
    # 그룹 도형 → 건너뛰기
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            logger.debug("그룹 도형 건너뜀")
            return None
    except (AttributeError, TypeError):
        pass

    # 이미지 도형 확인 (MSO_SHAPE_TYPE.PICTURE 또는 .image 속성)
    is_picture = False
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            is_picture = True
    except (AttributeError, TypeError):
        pass
    if not is_picture:
        try:
            _ = shape.image
            is_picture = True
        except Exception:
            pass

    if is_picture:
        return _extract_image_shape(shape, z_index)

    # 테이블
    try:
        if shape.has_table:
            return _extract_table_as_text(shape, z_index)
    except (AttributeError, TypeError):
        pass

    # 텍스트 프레임이 있는 도형 → 텍스트 오브젝트로 우선 처리
    has_text = False
    try:
        if shape.has_text_frame:
            text_content = _extract_text_content(shape.text_frame)
            if text_content.strip():
                has_text = True
    except (AttributeError, TypeError):
        pass
    if not has_text:
        try:
            if hasattr(shape, "text_frame"):
                text_content = _extract_text_content(shape.text_frame)
                if text_content.strip():
                    has_text = True
        except (AttributeError, TypeError):
            pass

    if has_text:
        # AutoShape에 텍스트 포함 → _extract_auto_shape가 텍스트로 변환
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                return _extract_auto_shape(shape, z_index)
        except (AttributeError, TypeError):
            pass
        # 일반 텍스트 프레임
        return _extract_text_shape(shape, z_index)

    # 텍스트 없는 자동 도형 → shape 그대로 유지
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            return _extract_auto_shape(shape, z_index)
    except (AttributeError, TypeError):
        pass

    # 기타 도형 → None (건너뛰기)
    return None


# ============ 메인 함수 ============


async def import_pptx_as_template(file_path: str, template_name: str, user_key: str) -> dict:
    """PPTX 파일을 분석하여 템플릿 + 슬라이드를 자동 생성

    Args:
        file_path: PPTX 파일 경로
        template_name: 생성할 템플릿 이름
        user_key: 생성자 사용자 키

    Returns:
        {"template_id": str, "slides_count": int, "classification": [...]}
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"PPTX 파일을 찾을 수 없습니다: {file_path}")

    # PPTX 파일 열기
    prs = Presentation(file_path)
    slides = list(prs.slides)
    total_slides = len(slides)

    if total_slides == 0:
        raise ValueError("슬라이드가 없는 PPTX 파일입니다")

    db = get_db()
    now = datetime.utcnow()

    # 슬라이드 파싱
    parsed_slides = []
    classifications = []

    for idx, pptx_slide in enumerate(slides):
        try:
            parsed = _parse_slide(pptx_slide, idx, total_slides)
            parsed_slides.append(parsed)
            classifications.append({
                "slide_order": idx + 1,
                "content_type": parsed["classification"]["content_type"],
                "layout": parsed["classification"]["layout"],
                "objects_count": len(parsed["objects"]),
            })
        except Exception as e:
            logger.error(f"슬라이드 {idx + 1} 파싱 실패: {e}")
            # 실패한 슬라이드는 빈 body 슬라이드로 추가
            parsed_slides.append({
                "objects": [],
                "classification": {"content_type": "body", "layout": "single_column"},
                "slide_meta": {
                    "content_type": "body",
                    "layout": "single_column",
                    "has_title": False,
                    "has_governance": False,
                    "description_count": 0,
                },
                "background_image": None,
            })
            classifications.append({
                "slide_order": idx + 1,
                "content_type": "body",
                "layout": "single_column",
                "objects_count": 0,
            })

    # 첫 번째 슬라이드 배경을 템플릿 기본 배경으로 사용
    template_bg = parsed_slides[0]["background_image"] if parsed_slides else None

    # 1. 템플릿 생성 (MongoDB)
    template_doc = {
        "name": template_name,
        "description": f"PPTX 파일에서 자동 생성 ({total_slides}개 슬라이드)",
        "background_image": template_bg,
        "is_published": False,
        "created_by": user_key,
        "created_at": now,
        "updated_at": now,
    }
    result = await db.templates.insert_one(template_doc)
    template_id = str(result.inserted_id)

    # 2. 슬라이드 생성 (MongoDB - 일괄 삽입)
    slide_docs = []
    for idx, parsed in enumerate(parsed_slides):
        slide_doc = {
            "template_id": template_id,
            "order": idx + 1,
            "objects": parsed["objects"],
            "slide_meta": parsed["slide_meta"],
            "background_image": parsed["background_image"],
            "created_at": now,
            "updated_at": now,
        }
        slide_docs.append(slide_doc)

    if slide_docs:
        await db.slides.insert_many(slide_docs)

    logger.info(
        f"PPTX 임포트 완료: template_id={template_id}, "
        f"slides={total_slides}, user={user_key}"
    )

    return {
        "template_id": template_id,
        "slides_count": total_slides,
        "classification": classifications,
    }
