from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from services.mongo_service import get_db
from services import redis_service
from bson import ObjectId

router = APIRouter(tags=["fonts"])

_FONTS_CACHE_KEY = "fonts:all"
_FONTS_PUBLIC_CACHE_KEY = "fonts:public"
_FONTS_TTL = 604800  # 7일 (폰트는 거의 변경되지 않음)


class FontCreate(BaseModel):
    name: str
    family: str
    url: str = ""


async def _invalidate_fonts_cache():
    """폰트 캐시 무효화"""
    await redis_service.cache_delete(_FONTS_CACHE_KEY, _FONTS_PUBLIC_CACHE_KEY)


@router.get("/{jwt_token}/api/fonts")
async def list_fonts(jwt_token: str):
    """등록된 폰트 목록 조회 (Redis 캐시)"""
    cached = await redis_service.cache_get(_FONTS_CACHE_KEY)
    if cached is not None:
        return {"fonts": cached}

    db = get_db()
    cursor = db.fonts.find().sort("name", 1)
    fonts = []
    async for f in cursor:
        f["_id"] = str(f["_id"])
        fonts.append(f)

    await redis_service.cache_set(_FONTS_CACHE_KEY, fonts, ttl=_FONTS_TTL)
    return {"fonts": fonts}


@router.get("/api/fonts/public")
async def list_fonts_public():
    """공개 폰트 목록 (공유 페이지용, 인증 불필요, Redis 캐시)"""
    cached = await redis_service.cache_get(_FONTS_PUBLIC_CACHE_KEY)
    if cached is not None:
        return {"fonts": cached}

    db = get_db()
    cursor = db.fonts.find({}, {"name": 1, "family": 1, "url": 1}).sort("name", 1)
    fonts = []
    async for f in cursor:
        f["_id"] = str(f["_id"])
        fonts.append(f)

    await redis_service.cache_set(_FONTS_PUBLIC_CACHE_KEY, fonts, ttl=_FONTS_TTL)
    return {"fonts": fonts}


@router.post("/{jwt_token}/api/admin/fonts")
async def add_font(jwt_token: str, body: FontCreate):
    """폰트 등록"""
    db = get_db()
    doc = {"name": body.name, "family": body.family, "url": body.url}
    try:
        result = await db.fonts.insert_one(doc)
        doc["_id"] = str(result.inserted_id)
        await _invalidate_fonts_cache()
        return {"font": doc}
    except Exception:
        raise HTTPException(status_code=409, detail="이미 등록된 폰트입니다")


@router.delete("/{jwt_token}/api/admin/fonts/{font_id}")
async def delete_font(jwt_token: str, font_id: str):
    """폰트 삭제"""
    db = get_db()
    await db.fonts.delete_one({"_id": ObjectId(font_id)})
    await _invalidate_fonts_cache()
    return {"success": True}
