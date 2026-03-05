from pydantic import BaseModel, Field
from typing import Optional
from datetime import datetime


class TextObjectStyle(BaseModel):
    font_family: str = "Arial"
    font_size: int = 16
    color: str = "#000000"
    bold: bool = False
    italic: bool = False
    align: str = "left"


class ShapeStyle(BaseModel):
    shape_type: str = "rectangle"  # "rectangle" | "rounded_rectangle" | "ellipse" | "line" | "arrow"
    fill_color: str = "#4A90D9"
    fill_opacity: float = 1.0  # 0.0 ~ 1.0
    stroke_color: str = "#333333"
    stroke_width: int = 2
    stroke_dash: str = "solid"  # "solid" | "dashed" | "dotted"
    border_radius: int = 12  # rounded_rectangle 전용
    arrow_head: str = "end"  # "none" | "end" | "both" (arrow 전용)


class SlideObject(BaseModel):
    obj_id: str
    obj_type: str  # "image" | "text" | "shape"
    x: float = 0
    y: float = 0
    width: float = 200
    height: float = 100
    # 이미지용
    image_url: Optional[str] = None
    # 텍스트용
    text_content: Optional[str] = None
    text_style: Optional[TextObjectStyle] = None
    # 도형용
    shape_style: Optional[ShapeStyle] = None
    # 메타 정보 (슬라이드 추천용)
    role: Optional[str] = None  # "title" | "governance" | "subtitle" | "body" | "description"
    placeholder: Optional[str] = None  # 사용자 데이터가 들어갈 플레이스홀더 이름


class SlideCreate(BaseModel):
    template_id: str
    order: int = 0
    objects: list[SlideObject] = []
    slide_meta: dict = {}  # 슬라이드 메타정보 (제목여부, 설명 텍스트 수 등)


class SlideUpdate(BaseModel):
    objects: Optional[list[SlideObject]] = None
    order: Optional[int] = None
    slide_meta: Optional[dict] = None


class TemplateCreate(BaseModel):
    name: str
    description: str = ""
    background_image: Optional[str] = None


class TemplateUpdate(BaseModel):
    name: Optional[str] = None
    description: Optional[str] = None
    background_image: Optional[str] = None
