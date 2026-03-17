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


class TableStyle(BaseModel):
    rows: int = 3
    cols: int = 3
    data: list = []  # [[cell_text, ...], ...] 2D array
    header_row: bool = True
    banded_rows: bool = False
    banded_cols: bool = False
    cell_styles: dict = {}  # "r_c": {bg_color, text_color, text_align, bold}
    border_color: str = "#BFBFBF"
    border_width: int = 1
    header_bg_color: str = "#4472C4"
    header_text_color: str = "#FFFFFF"
    font_family: str = "Arial"
    font_size: int = 11
    merged_cells: list = []  # [{start_row, start_col, end_row, end_col}]


class ChartStyle(BaseModel):
    chart_type: str = "bar"  # bar, line, pie, doughnut, area, radar
    chart_data: dict = {}  # {labels:[], datasets:[{label, data, backgroundColor}]}
    title: str = ""
    show_legend: bool = True
    show_grid: bool = True
    color_scheme: list = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"]


class SlideObject(BaseModel):
    obj_id: str
    obj_type: str  # "image" | "text" | "shape" | "image_area" | "table" | "chart"
    x: float = 0
    y: float = 0
    width: float = 200
    height: float = 100
    # 이미지용
    image_url: Optional[str] = None
    image_fit: Optional[str] = "contain"  # "contain" | "cover"
    # 텍스트용
    text_content: Optional[str] = None
    text_style: Optional[TextObjectStyle] = None
    # 도형용
    shape_style: Optional[ShapeStyle] = None
    # 테이블용
    table_style: Optional[TableStyle] = None
    # 차트용
    chart_style: Optional[ChartStyle] = None
    # 레이어 순서 (z-index)
    z_index: int = 10
    # 메타 정보 (슬라이드 추천용)
    role: Optional[str] = None  # "title" | "governance" | "subtitle" | "body" | "description" | "table" | "chart"
    placeholder: Optional[str] = None  # 사용자 데이터가 들어갈 플레이스홀더 이름


class SlideCreate(BaseModel):
    template_id: str
    order: int = 0
    objects: list[SlideObject] = []
    slide_meta: dict = {}  # 슬라이드 메타정보 (제목여부, 설명 텍스트 수 등)
    background_image: Optional[str] = None  # 슬라이드별 배경 이미지


class SlideUpdate(BaseModel):
    objects: Optional[list[SlideObject]] = None
    order: Optional[int] = None
    slide_meta: Optional[dict] = None
    background_image: Optional[str] = None  # 슬라이드별 배경 이미지 (빈 문자열이면 제거)


class TemplateCreate(BaseModel):
    name: str
    description: str = ""
    background_image: Optional[str] = None
    is_published: bool = False
    slide_size: str = "16:9"  # "16:9", "4:3", "A4"


class TemplateUpdate(BaseModel):
    name: Optional[str] = None
    description: Optional[str] = None
    background_image: Optional[str] = None
    is_published: Optional[bool] = None
    slide_size: Optional[str] = None


class BulkFontUpdate(BaseModel):
    from_font: Optional[str] = None  # None이면 전체 텍스트 대상
    to_font: str
