"""파워포인트 스타일(PPT Style) Pydantic 모델

ppt_styles MongoDB 컬렉션 스키마 및 요청/응답 모델 정의.
관리자가 만든 스타일은 디자인 토큰(색상, 폰트), 폰트 참조, 샘플 이미지,
패턴 라이브러리, 구조화 프롬프트, 빌더 힌트 등을 포함한다.
"""

from pydantic import BaseModel, Field
from typing import Optional, Any
from datetime import datetime


# ============ 기본값 정의 ============

# 기본 컬러 팔레트 (CLAUDE.md 3.1절 - 기업 슬라이드 표준)
DEFAULT_COLORS: dict = {
    "primary": "#1C60EF",
    "light":   "#DBE8FE",
    "white":   "#FFFFFF",
    "ink":     "#0B1E3F",
    "grey":    "#6B7B9A",
    "line":    "#D5E2FB",
    "darker":  "#071949",
}

# 기본 폰트 사이즈
DEFAULT_FONT_SIZES: dict = {
    "h1":    26,
    "h2":    22,
    "body":  11,
    "label": 10,
    "stat":  88,
}

# 기본 패턴 라이브러리 (12종, 빌더 매핑 순서 준수)
DEFAULT_PATTERN_IDS: list[str] = [
    "cover",
    "toc",
    "chapter",
    "content_3col",
    "content_2col_hero",
    "content_2x2",
    "big_stat",
    "content_3col_icon_block",
    "content_2_numbered",
    "content_3col_sidebar",
    "content_2x2_top_line",
    "closing",
]


def default_design_tokens() -> dict:
    """design_tokens 기본 구조"""
    return {
        "colors": dict(DEFAULT_COLORS),
        "fonts": {
            "title_font_id": None,
            "body_font_id":  None,
            "sizes": dict(DEFAULT_FONT_SIZES),
        },
    }


def default_pattern_library() -> list[dict]:
    """12개 패턴 enabled=true 기본값"""
    return [
        {"id": pid, "enabled": True, "options": {}}
        for pid in DEFAULT_PATTERN_IDS
    ]


# ============ 요청/응답 모델 ============

class PPTStyleCreate(BaseModel):
    """PPT 스타일 신규 생성 - title만 필수, 나머지는 기본값 자동 설정"""
    title: str
    description: str = ""
    is_published: bool = False
    lang: str = "ko"
    design_tokens: Optional[dict] = None
    font_refs: Optional[list[str]] = None
    pattern_library: Optional[list[dict]] = None
    structurer_prompt: str = ""
    builder_hints: Optional[dict] = None


class PPTStyleUpdate(BaseModel):
    """PPT 스타일 부분 수정 (exclude_unset 사용)"""
    title: Optional[str] = None
    description: Optional[str] = None
    is_published: Optional[bool] = None
    lang: Optional[str] = None
    design_tokens: Optional[dict] = None
    font_refs: Optional[list[str]] = None
    sample_image_refs: Optional[list[dict]] = None
    vision_analysis: Optional[dict] = None
    pattern_library: Optional[list[dict]] = None
    extracted_patterns: Optional[list[dict]] = None
    structurer_prompt: Optional[str] = None
    builder_hints: Optional[dict] = None


class PPTStyleFontLink(BaseModel):
    """폰트 ID 연결 요청 본문"""
    font_ids: list[str] = []


class PPTStyleResponse(BaseModel):
    """PPT 스타일 응답 모델 (ObjectId → str)"""
    id: str = Field(alias="_id")
    title: str
    description: str = ""
    is_published: bool = False
    lang: str = "ko"
    design_tokens: dict = {}
    font_refs: list[str] = []
    sample_image_refs: list[dict] = []
    vision_analysis: dict = {}
    pattern_library: list[dict] = []
    extracted_patterns: list[dict] = []
    structurer_prompt: str = ""
    builder_hints: dict = {}
    created_by: Optional[str] = None
    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None

    class Config:
        populate_by_name = True


def build_new_style_doc(data: PPTStyleCreate, user_key: str) -> dict:
    """신규 PPT 스타일 도큐먼트 빌드 (DB 삽입 직전 사용)"""
    now = datetime.utcnow()
    design_tokens = data.design_tokens if data.design_tokens else default_design_tokens()
    pattern_library = data.pattern_library if data.pattern_library is not None else default_pattern_library()

    return {
        "title": data.title,
        "description": data.description or "",
        "is_published": bool(data.is_published),
        "lang": data.lang or "ko",
        "design_tokens": design_tokens,
        "font_refs": data.font_refs or [],
        "sample_image_refs": [],
        "vision_analysis": {},
        "pattern_library": pattern_library,
        "extracted_patterns": [],
        "structurer_prompt": data.structurer_prompt or "",
        "builder_hints": data.builder_hints or {},
        "created_by": user_key,
        "created_at": now,
        "updated_at": now,
    }


def serialize_style(doc: dict) -> dict:
    """DB 도큐먼트 → API 응답 (ObjectId → str, datetime → isoformat 호환)"""
    if not doc:
        return doc
    out = dict(doc)
    if "_id" in out and not isinstance(out["_id"], str):
        out["_id"] = str(out["_id"])
    return out
