from pydantic import BaseModel
from typing import Optional


class ProjectCreate(BaseModel):
    name: str
    description: str = ""
    template_id: Optional[str] = None
    project_type: str = "slide"  # "slide", "excel", "onlyoffice_pptx", "onlyoffice_xlsx", "onlyoffice_docx"


class ProjectUpdate(BaseModel):
    name: Optional[str] = None
    description: Optional[str] = None
    template_id: Optional[str] = None
    instructions: Optional[str] = None


class GenerateRequest(BaseModel):
    project_id: str
    template_id: str
    instructions: str = ""
    lang: str = ""  # 출력 언어 (ko/en/ja/zh), 빈 값이면 서버 기본 언어
    slide_count: str = "auto"  # 슬라이드 수: "auto" 또는 "5","10","15","20","25","30"


class SlideUpdateRequest(BaseModel):
    objects: list
    items: Optional[list] = None


class SlideReorderRequest(BaseModel):
    slide_ids: list[str]


class ManualSlideRequest(BaseModel):
    project_id: str
    template_slide_id: str
    insert_after_order: Optional[int] = None  # None이면 맨 끝에 추가


class SlideTextRequest(BaseModel):
    project_id: str
    slide_id: str
    instruction: str
    template_slide_id: str = ""
    current_content: Optional[dict] = None


class ExcelGenerateRequest(BaseModel):
    project_id: str
    instructions: str = ""
    lang: str = ""
    sheet_count: str = "auto"


class ExcelModifyRequest(BaseModel):
    project_id: str
    instruction: str
    current_data: dict
    lang: str = ""
    target_sheet_index: Optional[int] = None  # None이면 전체 시트, 숫자면 해당 시트만 수정


class ExcelChartRequest(BaseModel):
    project_id: str
    sheet_index: int = 0
    chart_type: str = "bar"
    title: Optional[str] = None


class DocxGenerateRequest(BaseModel):
    project_id: str
    instructions: str = ""
    lang: str = ""
    section_count: str = "auto"


class DocxModifyRequest(BaseModel):
    project_id: str
    instruction: str
    current_data: dict
    lang: str = ""
