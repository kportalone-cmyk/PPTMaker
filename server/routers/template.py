from fastapi import APIRouter, HTTPException, UploadFile, File, Request
from bson import ObjectId
from datetime import datetime
from models.template import (
    TemplateCreate, TemplateUpdate,
    SlideCreate, SlideUpdate
)
from services.mongo_service import get_db
from services.auth_service import decode_jwt_token, get_user_flexible, is_admin
from config import settings
import aiofiles
import os
import uuid

router = APIRouter(tags=["templates"])


async def verify_admin(jwt_token: str):
    """관리자 권한 검증"""
    payload = decode_jwt_token(jwt_token)
    if not payload:
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")
    user = await get_user_flexible(payload)
    if not user or not is_admin(user):
        raise HTTPException(status_code=403, detail="관리자 권한이 필요합니다")
    return user


# ============ 템플릿 CRUD ============

@router.get("/{jwt_token}/api/admin/templates")
async def list_templates(jwt_token: str):
    """템플릿 목록 조회"""
    await verify_admin(jwt_token)
    db = get_db()
    cursor = db.templates.find().sort("created_at", -1)
    templates = []
    async for t in cursor:
        t["_id"] = str(t["_id"])
        templates.append(t)
    return {"templates": templates}


@router.post("/{jwt_token}/api/admin/templates")
async def create_template(jwt_token: str, data: TemplateCreate):
    """템플릿 생성"""
    user = await verify_admin(jwt_token)
    db = get_db()
    doc = {
        "name": data.name,
        "description": data.description,
        "background_image": data.background_image,
        "created_by": user.get("ky"),
        "created_at": datetime.utcnow(),
        "updated_at": datetime.utcnow(),
    }
    result = await db.templates.insert_one(doc)
    doc["_id"] = str(result.inserted_id)
    return {"template": doc}


@router.get("/{jwt_token}/api/admin/templates/{template_id}")
async def get_template(jwt_token: str, template_id: str):
    """템플릿 상세 조회"""
    await verify_admin(jwt_token)
    db = get_db()
    template = await db.templates.find_one({"_id": ObjectId(template_id)})
    if not template:
        raise HTTPException(status_code=404, detail="템플릿을 찾을 수 없습니다")
    template["_id"] = str(template["_id"])

    # 해당 템플릿의 슬라이드 목록
    cursor = db.slides.find({"template_id": template_id}).sort("order", 1)
    slides = []
    async for s in cursor:
        s["_id"] = str(s["_id"])
        slides.append(s)

    return {"template": template, "slides": slides}


@router.put("/{jwt_token}/api/admin/templates/{template_id}")
async def update_template(jwt_token: str, template_id: str, data: TemplateUpdate):
    """템플릿 수정"""
    await verify_admin(jwt_token)
    db = get_db()
    # exclude_unset: 클라이언트가 실제로 보낸 필드만 포함 (null 포함)
    update_data = data.dict(exclude_unset=True)
    update_data["updated_at"] = datetime.utcnow()
    result = await db.templates.update_one(
        {"_id": ObjectId(template_id)},
        {"$set": update_data}
    )
    if result.matched_count == 0:
        raise HTTPException(status_code=404, detail="템플릿을 찾을 수 없습니다")
    return {"success": True}


@router.delete("/{jwt_token}/api/admin/templates/{template_id}")
async def delete_template(jwt_token: str, template_id: str):
    """템플릿 삭제 (관련 슬라이드도 삭제)"""
    await verify_admin(jwt_token)
    db = get_db()
    await db.templates.delete_one({"_id": ObjectId(template_id)})
    await db.slides.delete_many({"template_id": template_id})
    return {"success": True}


# ============ 배경이미지 업로드 ============

@router.post("/{jwt_token}/api/admin/templates/{template_id}/background")
async def upload_background(jwt_token: str, template_id: str, file: UploadFile = File(...)):
    """배경 이미지 업로드"""
    await verify_admin(jwt_token)
    ext = os.path.splitext(file.filename)[1]
    filename = f"{uuid.uuid4().hex}{ext}"
    save_path = os.path.join(settings.UPLOAD_DIR, "backgrounds", filename)

    os.makedirs(os.path.dirname(save_path), exist_ok=True)
    async with aiofiles.open(save_path, "wb") as f:
        content = await file.read()
        await f.write(content)

    image_url = f"/uploads/backgrounds/{filename}"
    db = get_db()
    await db.templates.update_one(
        {"_id": ObjectId(template_id)},
        {"$set": {"background_image": image_url, "updated_at": datetime.utcnow()}}
    )
    return {"image_url": image_url}


# ============ 슬라이드 CRUD ============

@router.post("/{jwt_token}/api/admin/slides")
async def create_slide(jwt_token: str, data: SlideCreate):
    """슬라이드 생성"""
    await verify_admin(jwt_token)
    db = get_db()

    # 순서 자동 배정
    if data.order == 0:
        last = await db.slides.find_one(
            {"template_id": data.template_id},
            sort=[("order", -1)]
        )
        data.order = (last["order"] + 1) if last else 1

    doc = {
        "template_id": data.template_id,
        "order": data.order,
        "objects": [obj.dict() for obj in data.objects],
        "slide_meta": data.slide_meta,
        "created_at": datetime.utcnow(),
        "updated_at": datetime.utcnow(),
    }
    if data.background_image:
        doc["background_image"] = data.background_image
    result = await db.slides.insert_one(doc)
    doc["_id"] = str(result.inserted_id)
    return {"slide": doc}


@router.put("/{jwt_token}/api/admin/slides/{slide_id}")
async def update_slide(jwt_token: str, slide_id: str, data: SlideUpdate):
    """슬라이드 수정"""
    await verify_admin(jwt_token)
    db = get_db()
    update_data = {}
    if data.objects is not None:
        update_data["objects"] = [obj.dict() for obj in data.objects]
    if data.order is not None:
        update_data["order"] = data.order
    if data.slide_meta is not None:
        update_data["slide_meta"] = data.slide_meta
    if data.background_image is not None:
        update_data["background_image"] = data.background_image if data.background_image else None
    update_data["updated_at"] = datetime.utcnow()

    result = await db.slides.update_one(
        {"_id": ObjectId(slide_id)},
        {"$set": update_data}
    )
    if result.matched_count == 0:
        raise HTTPException(status_code=404, detail="슬라이드를 찾을 수 없습니다")
    return {"success": True}


@router.delete("/{jwt_token}/api/admin/slides/{slide_id}")
async def delete_slide(jwt_token: str, slide_id: str):
    """슬라이드 삭제"""
    await verify_admin(jwt_token)
    db = get_db()
    await db.slides.delete_one({"_id": ObjectId(slide_id)})
    return {"success": True}


# ============ 슬라이드 이미지 업로드 ============

@router.post("/{jwt_token}/api/admin/slides/upload-image")
async def upload_slide_image(jwt_token: str, file: UploadFile = File(...)):
    """슬라이드 이미지 오브젝트용 이미지 업로드"""
    await verify_admin(jwt_token)
    ext = os.path.splitext(file.filename)[1]
    filename = f"{uuid.uuid4().hex}{ext}"
    save_path = os.path.join(settings.UPLOAD_DIR, "images", filename)

    os.makedirs(os.path.dirname(save_path), exist_ok=True)
    async with aiofiles.open(save_path, "wb") as f:
        content = await file.read()
        await f.write(content)

    image_url = f"/uploads/images/{filename}"
    return {"image_url": image_url}
