from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from services.mongo_service import get_db
from bson import ObjectId
import os
import uuid
from config import settings


def hex_to_rgb(hex_color: str) -> RGBColor:
    """HEX 색상을 RGBColor로 변환"""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) < 6:
        hex_color = "000000"
    r, g, b = int(hex_color[:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    return RGBColor(r, g, b)


def get_alignment(align: str):
    """텍스트 정렬 매핑"""
    mapping = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
    }
    return mapping.get(align, PP_ALIGN.LEFT)


async def generate_pptx(project_id: str) -> str:
    """생성된 슬라이드 데이터를 기반으로 PPTX 파일 생성"""
    db = get_db()

    project = await db.projects.find_one({"_id": ObjectId(project_id)})
    if not project:
        raise ValueError("프로젝트를 찾을 수 없습니다")

    # 생성된 슬라이드 조회
    cursor = db.generated_slides.find({"project_id": project_id}).sort("order", 1)
    gen_slides = []
    async for s in cursor:
        gen_slides.append(s)

    if not gen_slides:
        raise ValueError("생성된 슬라이드가 없습니다")

    # 템플릿 배경 이미지
    template = None
    if project.get("template_id"):
        template = await db.templates.find_one({"_id": ObjectId(project["template_id"])})

    prs = Presentation()
    # 16:9 비율
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)

    slide_width_px = 960
    slide_height_px = 540

    for gen_slide in gen_slides:
        slide_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)

        # 배경 이미지 적용
        bg_image = gen_slide.get("background_image") or (template.get("background_image") if template else None)
        if bg_image:
            bg_path = os.path.join(".", bg_image.lstrip("/"))
            if os.path.exists(bg_path):
                slide.shapes.add_picture(
                    bg_path, Emu(0), Emu(0),
                    prs.slide_width, prs.slide_height
                )

        # 오브젝트 렌더링
        for obj in gen_slide.get("objects", []):
            # 캔버스 좌표 → PPTX EMU 변환
            x_emu = int(obj["x"] / slide_width_px * 12192000)
            y_emu = int(obj["y"] / slide_height_px * 6858000)
            w_emu = int(obj["width"] / slide_width_px * 12192000)
            h_emu = int(obj["height"] / slide_height_px * 6858000)

            if obj["obj_type"] == "image" and obj.get("image_url"):
                img_path = os.path.join(".", obj["image_url"].lstrip("/"))
                if os.path.exists(img_path):
                    slide.shapes.add_picture(
                        img_path, Emu(x_emu), Emu(y_emu),
                        Emu(w_emu), Emu(h_emu)
                    )

            elif obj["obj_type"] == "text":
                txBox = slide.shapes.add_textbox(
                    Emu(x_emu), Emu(y_emu),
                    Emu(w_emu), Emu(h_emu)
                )
                tf = txBox.text_frame
                tf.word_wrap = True

                # description 역할: 텍스트 넘침 시 높이 자동 확장
                role = obj.get("role", "") or obj.get("_auto_role", "")
                if role == "description":
                    tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

                text_content = obj.get("generated_text") or obj.get("text_content", "")
                style = obj.get("text_style", {})

                p = tf.paragraphs[0]
                p.text = text_content
                p.alignment = get_alignment(style.get("align", "left"))

                run = p.runs[0] if p.runs else p.add_run()
                run.text = text_content
                run.font.size = Pt(style.get("font_size", 16))
                run.font.bold = style.get("bold", False)
                run.font.italic = style.get("italic", False)

                color = style.get("color", "#000000")
                run.font.color.rgb = hex_to_rgb(color)

                if style.get("font_family"):
                    run.font.name = style["font_family"]

            elif obj["obj_type"] == "shape":
                _render_shape(slide, obj, x_emu, y_emu, w_emu, h_emu)

    # 파일 저장
    output_dir = os.path.join(settings.UPLOAD_DIR, "generated")
    os.makedirs(output_dir, exist_ok=True)
    filename = f"{uuid.uuid4().hex}.pptx"
    output_path = os.path.join(output_dir, filename)
    prs.save(output_path)

    return f"/uploads/generated/{filename}"


def _render_shape(slide, obj: dict, x_emu: int, y_emu: int, w_emu: int, h_emu: int):
    """도형 오브젝트를 PPTX 슬라이드에 렌더링"""
    shape_style = obj.get("shape_style", {})
    shape_type = shape_style.get("shape_type", "rectangle")

    if shape_type in ("line", "arrow"):
        _render_line_shape(slide, shape_style, x_emu, y_emu, w_emu, h_emu)
    else:
        _render_block_shape(slide, shape_style, shape_type, x_emu, y_emu, w_emu, h_emu)


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
        shape.fill.fore_color.rgb = hex_to_rgb(fill_color)
        if fill_opacity < 1.0:
            _set_fill_opacity(shape, fill_opacity)

    # 테두리
    stroke_width = style.get("stroke_width", 2)
    if stroke_width > 0:
        shape.line.color.rgb = hex_to_rgb(style.get("stroke_color", "#333333"))
        shape.line.width = Pt(stroke_width)
        _set_line_dash(shape, style.get("stroke_dash", "solid"))
    else:
        shape.line.fill.background()


def _render_line_shape(slide, style: dict, x: int, y: int, w: int, h: int):
    """라인/화살표 도형 렌더링 (얇은 사각형으로 구현)"""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(max(h, Pt(1).emu)))

    # 채우기 투명
    shape.fill.background()

    # 선 스타일
    stroke_width = style.get("stroke_width", 2)
    shape.line.color.rgb = hex_to_rgb(style.get("stroke_color", "#333333"))
    shape.line.width = Pt(stroke_width)
    _set_line_dash(shape, style.get("stroke_dash", "solid"))

    # 화살표 머리 (XML로 추가)
    shape_type = style.get("shape_type", "line")
    if shape_type == "arrow":
        try:
            ln_elem = shape.line._ln
            tail_end = ln_elem.makeelement(qn('a:tailEnd'), {})
            tail_end.set('type', 'triangle')
            tail_end.set('w', 'med')
            tail_end.set('len', 'med')
            ln_elem.append(tail_end)

            if style.get("arrow_head") == "both":
                head_end = ln_elem.makeelement(qn('a:headEnd'), {})
                head_end.set('type', 'triangle')
                head_end.set('w', 'med')
                head_end.set('len', 'med')
                ln_elem.append(head_end)
        except Exception:
            pass  # 화살표 렌더링 실패해도 기본 선으로 표시


def _set_fill_opacity(shape, opacity: float):
    """도형 채우기 투명도 설정 (XML 직접 조작)"""
    try:
        solid_fill = shape.fill._fill.find(qn('a:solidFill'))
        if solid_fill is not None:
            srgb = solid_fill.find(qn('a:srgbClr'))
            if srgb is not None:
                alpha = srgb.makeelement(qn('a:alpha'), {})
                alpha.set('val', str(int(opacity * 100000)))
                srgb.append(alpha)
    except Exception:
        pass


def _set_line_dash(shape, dash_style: str):
    """테두리 선 스타일 설정"""
    try:
        if dash_style == "dashed":
            shape.line.dash_style = 4  # MSO_LINE_DASH_STYLE.DASH
        elif dash_style == "dotted":
            shape.line.dash_style = 3  # MSO_LINE_DASH_STYLE.ROUND_DOT
    except Exception:
        pass
