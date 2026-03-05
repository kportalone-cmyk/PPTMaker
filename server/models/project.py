from pydantic import BaseModel
from typing import Optional


class ProjectCreate(BaseModel):
    name: str
    description: str = ""
    template_id: Optional[str] = None


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
