from services.mongo_service import get_db
from bson import ObjectId


async def get_template_slides(template_id: str) -> list:
    """템플릿의 슬라이드 목록 조회"""
    db = get_db()
    cursor = db.slides.find({"template_id": template_id}).sort("order", 1)
    slides = []
    async for s in cursor:
        s["_id"] = str(s["_id"])
        slides.append(s)
    return slides


async def recommend_slide(template_id: str, content_meta: dict) -> dict | None:
    """콘텐츠 메타정보 기반 최적 슬라이드 추천

    content_meta 예시:
    {
        "has_title": True,
        "has_governance": False,
        "description_count": 3,  # 서브 텍스트 설명 수
        "content_type": "body",  # "title_slide" | "toc" | "section_divider" | "body" | "closing"
        "layout": "two_column"   # "cover" | "numbered_list" | "divider" | "single_column" | "two_column" | "grid" | "closing"
    }
    """
    db = get_db()
    slides = await get_template_slides(template_id)

    best_match = None
    best_score = -1

    for slide in slides:
        meta = slide.get("slide_meta", {})
        score = 0

        # 콘텐츠 타입 매칭 (가장 중요)
        if meta.get("content_type") == content_meta.get("content_type"):
            score += 10

        # 레이아웃 매칭
        if content_meta.get("layout") and meta.get("layout") == content_meta.get("layout"):
            score += 7

        # 제목 존재 여부
        if meta.get("has_title") == content_meta.get("has_title"):
            score += 5

        # 거버넌스 존재 여부
        if meta.get("has_governance") == content_meta.get("has_governance"):
            score += 3

        # 설명 텍스트 수 매칭
        slide_desc_count = meta.get("description_count", 0)
        content_desc_count = content_meta.get("description_count", 0)
        if slide_desc_count == content_desc_count:
            score += 8
        elif abs(slide_desc_count - content_desc_count) <= 1:
            score += 4

        if score > best_score:
            best_score = score
            best_match = slide

    return best_match


async def get_all_templates_summary() -> list:
    """모든 템플릿의 요약 정보 (사용자 선택용) - 첫 슬라이드 포함"""
    db = get_db()
    cursor = db.templates.find().sort("name", 1)
    templates = []
    async for t in cursor:
        t["_id"] = str(t["_id"])
        # 슬라이드 수 포함
        slide_count = await db.slides.count_documents({"template_id": str(t["_id"])})
        t["slide_count"] = slide_count
        # 첫 번째 슬라이드 (썸네일 미리보기용)
        first_slide = await db.slides.find_one(
            {"template_id": str(t["_id"])},
            sort=[("order", 1)]
        )
        if first_slide:
            first_slide["_id"] = str(first_slide["_id"])
            t["first_slide"] = first_slide
        else:
            t["first_slide"] = None
        templates.append(t)
    return templates
