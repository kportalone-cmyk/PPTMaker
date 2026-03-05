"""PPTX 생성 서비스 - python-pptx 기반 슬라이드 렌더링"""

from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from services.mongo_service import get_db
from bson import ObjectId
import os
import uuid
from config import settings

# 캔버스 / PPTX 상수
SLIDE_W_PX = 960
SLIDE_H_PX = 540
SLIDE_W_EMU = 12192000
SLIDE_H_EMU = 6858000


def _hex_to_rgb(hex_color: str) -> RGBColor:
    """HEX 색상을 RGBColor로 변환"""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) < 6:
        hex_color = "000000"
    r, g, b = int(hex_color[:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    return RGBColor(r, g, b)


def _get_alignment(align: str):
    """텍스트 정렬 매핑"""
    return {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}.get(
        align, PP_ALIGN.LEFT
    )


def _px_to_emu(x, y, w, h):
    """캔버스 좌표(px) → PPTX EMU 변환"""
    return (
        int(x / SLIDE_W_PX * SLIDE_W_EMU),
        int(y / SLIDE_H_PX * SLIDE_H_EMU),
        int(w / SLIDE_W_PX * SLIDE_W_EMU),
        int(h / SLIDE_H_PX * SLIDE_H_EMU),
    )


async def generate_pptx(project_id: str) -> str:
    """생성된 슬라이드 데이터를 기반으로 PPTX 파일 생성"""
    db = get_db()

    project = await db.projects.find_one({"_id": ObjectId(project_id)})
    if not project:
        raise ValueError("프로젝트를 찾을 수 없습니다")

    cursor = db.generated_slides.find({"project_id": project_id}).sort("order", 1)
    gen_slides = []
    async for s in cursor:
        gen_slides.append(s)

    if not gen_slides:
        raise ValueError("생성된 슬라이드가 없습니다")

    template = None
    if project.get("template_id"):
        template = await db.templates.find_one({"_id": ObjectId(project["template_id"])})

    prs = Presentation()
    prs.slide_width = Emu(SLIDE_W_EMU)
    prs.slide_height = Emu(SLIDE_H_EMU)

    for gen_slide in gen_slides:
        slide_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)

        # 배경 이미지
        bg_image = gen_slide.get("background_image") or (
            template.get("background_image") if template else None
        )
        if bg_image:
            bg_path = os.path.join(".", bg_image.lstrip("/"))
            if os.path.exists(bg_path):
                try:
                    slide.shapes.add_picture(
                        bg_path, Emu(0), Emu(0), prs.slide_width, prs.slide_height,
                    )
                except Exception:
                    pass

        # 오브젝트 렌더링 (items 소진된 초과 subtitle/description 제거)
        items = gen_slide.get("items", [])
        sub_idx = 0
        desc_idx = 0
        for obj in gen_slide.get("objects", []):
            role = obj.get("role", "") or obj.get("_auto_role", "")
            if items and role == "subtitle":
                if sub_idx >= len(items):
                    continue
                sub_idx += 1
            elif items and role == "description":
                if desc_idx >= len(items):
                    continue
                desc_idx += 1

            xe, ye, we, he = _px_to_emu(obj["x"], obj["y"], obj["width"], obj["height"])

            if obj["obj_type"] == "image" and obj.get("image_url"):
                img_path = os.path.join(".", obj["image_url"].lstrip("/"))
                if os.path.exists(img_path):
                    try:
                        slide.shapes.add_picture(
                            img_path, Emu(xe), Emu(ye), Emu(we), Emu(he),
                        )
                    except Exception:
                        pass

            elif obj["obj_type"] == "text":
                _render_text(slide, obj, xe, ye, we, he)

            elif obj["obj_type"] == "shape":
                _render_shape(slide, obj, xe, ye, we, he)

    # 파일 저장
    output_dir = os.path.join(settings.UPLOAD_DIR, "generated")
    os.makedirs(output_dir, exist_ok=True)
    filename = f"{uuid.uuid4().hex}.pptx"
    output_path = os.path.join(output_dir, filename)
    prs.save(output_path)

    return f"/uploads/generated/{filename}"


# ============ 텍스트 렌더링 ============

def _render_text(slide, obj: dict, x: int, y: int, w: int, h: int):
    """텍스트 오브젝트를 PPTX 슬라이드에 렌더링"""
    txBox = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = txBox.text_frame
    tf.word_wrap = True

    role = obj.get("role", "") or obj.get("_auto_role", "")
    if role == "description":
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    text_content = obj.get("generated_text") or obj.get("text_content", "")
    style = obj.get("text_style", {})

    p = tf.paragraphs[0]
    p.alignment = _get_alignment(style.get("align", "left"))

    run = p.add_run()
    run.text = text_content
    run.font.size = Pt(style.get("font_size", 16))
    run.font.bold = style.get("bold", False)
    run.font.italic = style.get("italic", False)
    run.font.color.rgb = _hex_to_rgb(style.get("color", "#000000"))

    if style.get("font_family"):
        run.font.name = style["font_family"]


# ============ 도형 렌더링 ============

def _render_shape(slide, obj: dict, x: int, y: int, w: int, h: int):
    """도형 오브젝트를 PPTX 슬라이드에 렌더링"""
    style = obj.get("shape_style", {})
    shape_type = style.get("shape_type", "rectangle")

    if shape_type in ("line", "arrow"):
        _render_line_shape(slide, style, x, y, w, h)
    else:
        _render_block_shape(slide, style, shape_type, x, y, w, h)


def _render_block_shape(slide, style: dict, shape_type: str, x: int, y: int, w: int, h: int):
    """블록 도형 (사각형, 둥근 사각형, 타원) 렌더링"""
    shape_map = {
        "rectangle": MSO_SHAPE.RECTANGLE,
        "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
        "ellipse": MSO_SHAPE.OVAL,
    }
    shape_enum = shape_map.get(shape_type, MSO_SHAPE.RECTANGLE)
    shape = slide.shapes.add_shape(shape_enum, Emu(x), Emu(y), Emu(w), Emu(h))

    # 채우기
    fill_opacity = style.get("fill_opacity", 1.0)
    fill_color = style.get("fill_color", "#4A90D9")

    if fill_opacity == 0 or fill_color in ("transparent", "none"):
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _hex_to_rgb(fill_color)
        if fill_opacity < 1.0:
            _set_fill_opacity(shape, fill_opacity)

    # 둥근 사각형 커스텀 border_radius
    if shape_type == "rounded_rectangle" and style.get("border_radius"):
        _set_corner_radius(shape, style["border_radius"], w, h)

    # 테두리
    stroke_width = style.get("stroke_width", 2)
    if stroke_width > 0:
        shape.line.color.rgb = _hex_to_rgb(style.get("stroke_color", "#333333"))
        shape.line.width = Pt(stroke_width)
        _set_line_dash(shape, style.get("stroke_dash", "solid"))
    else:
        shape.line.fill.background()


def _render_line_shape(slide, style: dict, x: int, y: int, w: int, h: int):
    """라인/화살표 도형 렌더링"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(max(h, Pt(1).emu))
    )

    shape.fill.background()

    stroke_width = style.get("stroke_width", 2)
    shape.line.color.rgb = _hex_to_rgb(style.get("stroke_color", "#333333"))
    shape.line.width = Pt(stroke_width)
    _set_line_dash(shape, style.get("stroke_dash", "solid"))

    if style.get("shape_type") == "arrow":
        try:
            ln_elem = shape.line._ln
            tail_end = ln_elem.makeelement(qn("a:tailEnd"), {})
            tail_end.set("type", "triangle")
            tail_end.set("w", "med")
            tail_end.set("len", "med")
            ln_elem.append(tail_end)

            if style.get("arrow_head") == "both":
                head_end = ln_elem.makeelement(qn("a:headEnd"), {})
                head_end.set("type", "triangle")
                head_end.set("w", "med")
                head_end.set("len", "med")
                ln_elem.append(head_end)
        except Exception:
            pass


# ============ 유틸리티 ============

def _set_fill_opacity(shape, opacity: float):
    """도형 채우기 투명도 설정 (XML 직접 조작)"""
    try:
        solid_fill = shape.fill._fill.find(qn("a:solidFill"))
        if solid_fill is not None:
            srgb = solid_fill.find(qn("a:srgbClr"))
            if srgb is not None:
                alpha = srgb.makeelement(qn("a:alpha"), {})
                alpha.set("val", str(int(opacity * 100000)))
                srgb.append(alpha)
    except Exception:
        pass


def _set_corner_radius(shape, border_radius_px: int, w_emu: int, h_emu: int):
    """둥근 사각형 모서리 반경 커스텀 설정"""
    try:
        min_dim_emu = min(w_emu, h_emu)
        if min_dim_emu <= 0:
            return
        radius_emu = int(border_radius_px / SLIDE_W_PX * SLIDE_W_EMU)
        adj_val = int(radius_emu / (min_dim_emu / 2) * 50000)
        adj_val = max(0, min(adj_val, 50000))

        sp = shape._element
        prst_geom = sp.find(qn("a:prstGeom"))
        if prst_geom is not None:
            av_lst = prst_geom.find(qn("a:avLst"))
            if av_lst is None:
                av_lst = prst_geom.makeelement(qn("a:avLst"), {})
                prst_geom.append(av_lst)
            for existing in av_lst.findall(qn("a:gd")):
                av_lst.remove(existing)
            gd = av_lst.makeelement(qn("a:gd"), {"name": "adj", "fmla": f"val {adj_val}"})
            av_lst.append(gd)
    except Exception:
        pass


def _set_line_dash(shape, dash_style: str):
    """테두리 선 스타일 설정"""
    try:
        if dash_style == "dashed":
            shape.line.dash_style = 4  # DASH
        elif dash_style == "dotted":
            shape.line.dash_style = 3  # ROUND_DOT
    except Exception:
        pass
