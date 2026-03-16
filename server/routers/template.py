from fastapi import APIRouter, HTTPException, UploadFile, File, Form, Request
from bson import ObjectId
from datetime import datetime, timedelta
from models.template import (
    TemplateCreate, TemplateUpdate,
    SlideCreate, SlideUpdate, BulkFontUpdate
)
from services.mongo_service import get_db, get_org_db
from services.auth_service import decode_jwt_token, get_user_flexible, is_admin
from services.template_service import invalidate_template_cache
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
        "is_published": data.is_published,
        "slide_size": data.slide_size,
        "created_by": user.get("ky"),
        "created_at": datetime.utcnow(),
        "updated_at": datetime.utcnow(),
    }
    result = await db.templates.insert_one(doc)
    doc["_id"] = str(result.inserted_id)
    await invalidate_template_cache()
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
    await invalidate_template_cache(template_id)
    return {"success": True}


@router.put("/{jwt_token}/api/admin/templates/{template_id}/publish")
async def toggle_publish(jwt_token: str, template_id: str):
    """템플릿 게시/비게시 토글"""
    await verify_admin(jwt_token)
    db = get_db()
    template = await db.templates.find_one({"_id": ObjectId(template_id)})
    if not template:
        raise HTTPException(status_code=404, detail="템플릿을 찾을 수 없습니다")

    new_status = not template.get("is_published", False)
    await db.templates.update_one(
        {"_id": ObjectId(template_id)},
        {"$set": {"is_published": new_status, "updated_at": datetime.utcnow()}}
    )
    await invalidate_template_cache(template_id)
    return {"success": True, "is_published": new_status}


@router.delete("/{jwt_token}/api/admin/templates/{template_id}")
async def delete_template(jwt_token: str, template_id: str):
    """템플릿 삭제 (관련 슬라이드도 삭제)"""
    await verify_admin(jwt_token)
    db = get_db()
    await db.templates.delete_one({"_id": ObjectId(template_id)})
    await db.slides.delete_many({"template_id": template_id})
    await invalidate_template_cache(template_id)
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
    await invalidate_template_cache(template_id)
    return {"image_url": image_url}


# ============ PPTX 가져오기 ============

@router.post("/{jwt_token}/api/admin/templates/import-pptx")
async def import_pptx_template(jwt_token: str, file: UploadFile = File(...), template_name: str = Form(...)):
    """PPTX 파일을 분석하여 자동 템플릿 생성"""
    user = await verify_admin(jwt_token)

    if not file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail=".pptx 파일만 지원합니다")

    # 임시 파일 저장
    temp_filename = f"import_{uuid.uuid4().hex}.pptx"
    temp_path = os.path.join(settings.UPLOAD_DIR, temp_filename)
    os.makedirs(os.path.dirname(temp_path), exist_ok=True)

    try:
        async with aiofiles.open(temp_path, "wb") as f:
            content = await file.read()
            await f.write(content)

        from services.pptx_import_service import import_pptx_as_template
        result = await import_pptx_as_template(
            file_path=temp_path,
            template_name=template_name.strip(),
            user_key=user.get("ky", ""),
        )

        await invalidate_template_cache()
        return result

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PPTX 가져오기 실패: {str(e)}")
    finally:
        # 임시 파일 삭제
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


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
    await invalidate_template_cache(data.template_id)
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

    # 캐시 무효화를 위해 template_id 조회
    slide_doc = await db.slides.find_one({"_id": ObjectId(slide_id)}, {"template_id": 1})
    result = await db.slides.update_one(
        {"_id": ObjectId(slide_id)},
        {"$set": update_data}
    )
    if result.matched_count == 0:
        raise HTTPException(status_code=404, detail="슬라이드를 찾을 수 없습니다")
    if slide_doc:
        await invalidate_template_cache(slide_doc.get("template_id"))
    return {"success": True}


@router.delete("/{jwt_token}/api/admin/slides/{slide_id}")
async def delete_slide(jwt_token: str, slide_id: str):
    """슬라이드 삭제"""
    await verify_admin(jwt_token)
    db = get_db()
    slide_doc = await db.slides.find_one({"_id": ObjectId(slide_id)}, {"template_id": 1})
    await db.slides.delete_one({"_id": ObjectId(slide_id)})
    if slide_doc:
        await invalidate_template_cache(slide_doc.get("template_id"))
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


# ============ 폰트 일괄 변경 ============

@router.put("/{jwt_token}/api/admin/templates/{template_id}/bulk-font")
async def bulk_font_update(jwt_token: str, template_id: str, data: BulkFontUpdate):
    """템플릿 전체 슬라이드의 폰트 일괄 변경"""
    await verify_admin(jwt_token)
    db = get_db()

    template = await db.templates.find_one({"_id": ObjectId(template_id)})
    if not template:
        raise HTTPException(status_code=404, detail="템플릿을 찾을 수 없습니다")

    cursor = db.slides.find({"template_id": template_id})
    updated_count = 0

    async for slide in cursor:
        objects = slide.get("objects", [])
        modified = False

        for obj in objects:
            if obj.get("obj_type") != "text":
                continue
            text_style = obj.get("text_style")
            if not text_style:
                continue
            if data.from_font is None or text_style.get("font_family") == data.from_font:
                text_style["font_family"] = data.to_font
                updated_count += 1
                modified = True

        if modified:
            await db.slides.update_one(
                {"_id": slide["_id"]},
                {"$set": {"objects": objects, "updated_at": datetime.utcnow()}}
            )

    await invalidate_template_cache(template_id)
    return {"success": True, "updated_count": updated_count}


# ============ 대시보드 API ============

@router.get("/{jwt_token}/api/admin/dashboard/overview")
async def dashboard_overview(jwt_token: str):
    """대시보드 전체 현황"""
    await verify_admin(jwt_token)
    db = get_db()
    org_db = get_org_db()
    now = datetime.utcnow()
    month_start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)

    # 전체 사용자
    total_users = await org_db[settings.ORG_COLLECTION].count_documents({})

    # 활성 사용자 (프로젝트 생성한 고유 사용자)
    active_pipeline = [{"$group": {"_id": "$user_key"}}, {"$count": "count"}]
    active_result = await db.projects.aggregate(active_pipeline).to_list(1)
    active_users = active_result[0]["count"] if active_result else 0

    # 프로젝트 통계
    total_projects = await db.projects.count_documents({})
    month_projects = await db.projects.count_documents({"created_at": {"$gte": month_start}})
    today_projects = await db.projects.count_documents({"created_at": {"$gte": today_start}})
    generated_projects = await db.projects.count_documents({"status": "generated"})

    # 생성 건수
    total_gen_slides = await db.generated_slides.count_documents({})
    total_gen_excel = await db.generated_excel.count_documents({})
    total_gen_docx = await db.generated_docx.count_documents({})

    # 템플릿
    total_templates = await db.templates.count_documents({})
    published_templates = await db.templates.count_documents({"is_published": True})

    # 리소스
    total_resources = await db.resources.count_documents({})

    # 프로젝트 유형 분포
    type_pipeline = [{"$group": {"_id": {"$ifNull": ["$project_type", "slide"]}, "count": {"$sum": 1}}}]
    type_dist = await db.projects.aggregate(type_pipeline).to_list(100)

    # 일별 프로젝트 추이 (30일)
    thirty_days_ago = now - timedelta(days=30)
    daily_pipeline = [
        {"$match": {"created_at": {"$gte": thirty_days_ago}}},
        {"$group": {
            "_id": {"$dateToString": {"format": "%Y-%m-%d", "date": "$created_at"}},
            "count": {"$sum": 1}
        }},
        {"$sort": {"_id": 1}}
    ]
    daily_trend = await db.projects.aggregate(daily_pipeline).to_list(100)

    # 리소스 유형 분포
    res_pipeline = [{"$group": {"_id": "$resource_type", "count": {"$sum": 1}}}]
    resource_dist = await db.resources.aggregate(res_pipeline).to_list(100)

    return {
        "total_users": total_users,
        "active_users": active_users,
        "total_projects": total_projects,
        "month_projects": month_projects,
        "today_projects": today_projects,
        "generated_projects": generated_projects,
        "total_generations": total_gen_slides + total_gen_excel + total_gen_docx,
        "generation_breakdown": {
            "slides": total_gen_slides,
            "excel": total_gen_excel,
            "docx": total_gen_docx,
        },
        "total_templates": total_templates,
        "published_templates": published_templates,
        "total_resources": total_resources,
        "project_type_distribution": type_dist,
        "daily_trend": daily_trend,
        "resource_type_distribution": resource_dist,
    }


@router.get("/{jwt_token}/api/admin/dashboard/users")
async def dashboard_users(jwt_token: str, search: str = "", page: int = 1, limit: int = 20):
    """대시보드 사용자 현황"""
    await verify_admin(jwt_token)
    db = get_db()
    org_db = get_org_db()
    col = org_db[settings.ORG_COLLECTION]

    query = {}
    if search:
        query = {"$or": [
            {"nm": {"$regex": search, "$options": "i"}},
            {"dp": {"$regex": search, "$options": "i"}},
            {"em": {"$regex": search, "$options": "i"}},
        ]}

    total = await col.count_documents(query)
    skip = (page - 1) * limit

    users = []
    async for user in col.find(query).skip(skip).limit(limit).sort("nm", 1):
        user_key = user.get("ky", "")
        project_count = await db.projects.count_documents({"user_key": user_key})
        last_project = await db.projects.find_one(
            {"user_key": user_key},
            sort=[("updated_at", -1)],
            projection={"updated_at": 1}
        )
        users.append({
            "nm": user.get("nm", ""),
            "dp": user.get("dp", ""),
            "em": user.get("em", ""),
            "ky": user_key,
            "role": user.get("role", ""),
            "project_count": project_count,
            "last_activity": last_project.get("updated_at").isoformat() if last_project and last_project.get("updated_at") else None,
        })

    return {"users": users, "total": total, "page": page, "limit": limit}


@router.get("/{jwt_token}/api/admin/dashboard/projects")
async def dashboard_projects(jwt_token: str, search: str = "", project_type: str = "",
                              status: str = "", page: int = 1, limit: int = 20):
    """대시보드 프로젝트 현황"""
    await verify_admin(jwt_token)
    db = get_db()
    org_db = get_org_db()

    query = {}
    if search:
        query["name"] = {"$regex": search, "$options": "i"}
    if project_type:
        query["project_type"] = project_type
    if status:
        query["status"] = status

    total = await db.projects.count_documents(query)
    skip = (page - 1) * limit

    projects = []
    async for p in db.projects.find(query).skip(skip).limit(limit).sort("created_at", -1):
        pid = str(p["_id"])
        p["_id"] = pid
        user_key = p.get("user_key", "")
        user = await org_db[settings.ORG_COLLECTION].find_one(
            {"ky": user_key}, {"nm": 1, "dp": 1}
        )
        p["_user_name"] = user.get("nm", "") if user else user_key
        p["_user_dept"] = user.get("dp", "") if user else ""
        p["_slide_count"] = await db.generated_slides.count_documents({"project_id": pid})
        p["_resource_count"] = await db.resources.count_documents({"project_id": pid})
        # datetime serialization
        if p.get("created_at"):
            p["created_at"] = p["created_at"].isoformat()
        if p.get("updated_at"):
            p["updated_at"] = p["updated_at"].isoformat()
        projects.append(p)

    return {"projects": projects, "total": total, "page": page, "limit": limit}


@router.get("/{jwt_token}/api/admin/dashboard/activity")
async def dashboard_activity(jwt_token: str, page: int = 1, limit: int = 30,
                               user_search: str = "", date_from: str = "", date_to: str = ""):
    """대시보드 활동 로그"""
    await verify_admin(jwt_token)
    db = get_db()
    org_db = get_org_db()

    query = {"status": "generated"}
    if date_from:
        query.setdefault("updated_at", {})["$gte"] = datetime.fromisoformat(date_from)
    if date_to:
        query.setdefault("updated_at", {})["$lte"] = datetime.fromisoformat(date_to + "T23:59:59")

    if user_search:
        user_cursor = org_db[settings.ORG_COLLECTION].find(
            {"$or": [
                {"nm": {"$regex": user_search, "$options": "i"}},
                {"dp": {"$regex": user_search, "$options": "i"}},
            ]},
            {"ky": 1}
        )
        user_keys = [u.get("ky") async for u in user_cursor]
        if user_keys:
            query["user_key"] = {"$in": user_keys}
        else:
            return {"activities": [], "total": 0, "page": page, "limit": limit}

    total = await db.projects.count_documents(query)
    skip = (page - 1) * limit

    activities = []
    async for p in db.projects.find(query).skip(skip).limit(limit).sort("updated_at", -1):
        pid = str(p["_id"])
        user_key = p.get("user_key", "")
        user = await org_db[settings.ORG_COLLECTION].find_one(
            {"ky": user_key}, {"nm": 1, "dp": 1}
        )
        activities.append({
            "project_id": pid,
            "project_name": p.get("name", ""),
            "project_type": p.get("project_type", "slide"),
            "user_key": user_key,
            "user_name": user.get("nm", "") if user else user_key,
            "user_dept": user.get("dp", "") if user else "",
            "template_id": p.get("template_id"),
            "status": p.get("status"),
            "created_at": p.get("created_at").isoformat() if p.get("created_at") else None,
            "updated_at": p.get("updated_at").isoformat() if p.get("updated_at") else None,
        })

    return {"activities": activities, "total": total, "page": page, "limit": limit}


@router.get("/{jwt_token}/api/admin/dashboard/user/{user_key}")
async def dashboard_user_detail(jwt_token: str, user_key: str):
    """특정 사용자 상세 사용 현황"""
    await verify_admin(jwt_token)
    db = get_db()
    org_db = get_org_db()
    now = datetime.utcnow()
    month_start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)

    # 사용자 정보
    user = await org_db[settings.ORG_COLLECTION].find_one({"ky": user_key})
    if not user:
        raise HTTPException(status_code=404, detail="사용자를 찾을 수 없습니다")

    user_info = {
        "nm": user.get("nm", ""),
        "dp": user.get("dp", ""),
        "em": user.get("em", ""),
        "ky": user_key,
        "role": user.get("role", ""),
    }

    # 프로젝트 통계
    total_projects = await db.projects.count_documents({"user_key": user_key})
    month_projects = await db.projects.count_documents({"user_key": user_key, "created_at": {"$gte": month_start}})
    generated_projects = await db.projects.count_documents({"user_key": user_key, "status": "generated"})

    # 프로젝트 유형 분포
    type_pipeline = [
        {"$match": {"user_key": user_key}},
        {"$group": {"_id": {"$ifNull": ["$project_type", "slide"]}, "count": {"$sum": 1}}}
    ]
    type_dist = await db.projects.aggregate(type_pipeline).to_list(100)

    # 프로젝트 상태 분포
    status_pipeline = [
        {"$match": {"user_key": user_key}},
        {"$group": {"_id": "$status", "count": {"$sum": 1}}}
    ]
    status_dist = await db.projects.aggregate(status_pipeline).to_list(100)

    # 생성 건수 (프로젝트 ID 목록으로 조회)
    project_ids = []
    async for p in db.projects.find({"user_key": user_key}, {"_id": 1}):
        project_ids.append(str(p["_id"]))

    gen_slides = 0
    gen_excel = 0
    gen_docx = 0
    if project_ids:
        gen_slides = await db.generated_slides.count_documents({"project_id": {"$in": project_ids}})
        gen_excel = await db.generated_excel.count_documents({"project_id": {"$in": project_ids}})
        gen_docx = await db.generated_docx.count_documents({"project_id": {"$in": project_ids}})

    # 리소스 통계
    total_resources = 0
    res_dist = []
    if project_ids:
        total_resources = await db.resources.count_documents({"project_id": {"$in": project_ids}})
        res_pipeline = [
            {"$match": {"project_id": {"$in": project_ids}}},
            {"$group": {"_id": "$resource_type", "count": {"$sum": 1}}}
        ]
        res_dist = await db.resources.aggregate(res_pipeline).to_list(100)

    # 일별 프로젝트 추이 (30일)
    thirty_days_ago = now - timedelta(days=30)
    daily_pipeline = [
        {"$match": {"user_key": user_key, "created_at": {"$gte": thirty_days_ago}}},
        {"$group": {
            "_id": {"$dateToString": {"format": "%Y-%m-%d", "date": "$created_at"}},
            "count": {"$sum": 1}
        }},
        {"$sort": {"_id": 1}}
    ]
    daily_trend = await db.projects.aggregate(daily_pipeline).to_list(100)

    # 최근 프로젝트 목록 (최근 10개)
    recent_projects = []
    async for p in db.projects.find({"user_key": user_key}).sort("updated_at", -1).limit(10):
        pid = str(p["_id"])
        slide_count = await db.generated_slides.count_documents({"project_id": pid})
        if p.get("created_at"):
            p["created_at"] = p["created_at"].isoformat()
        if p.get("updated_at"):
            p["updated_at"] = p["updated_at"].isoformat()
        recent_projects.append({
            "project_id": pid,
            "name": p.get("name", ""),
            "project_type": p.get("project_type", "slide"),
            "status": p.get("status", "draft"),
            "slide_count": slide_count,
            "created_at": p.get("created_at"),
            "updated_at": p.get("updated_at"),
        })

    return {
        "user": user_info,
        "total_projects": total_projects,
        "month_projects": month_projects,
        "generated_projects": generated_projects,
        "total_generations": gen_slides + gen_excel + gen_docx,
        "generation_breakdown": {"slides": gen_slides, "excel": gen_excel, "docx": gen_docx},
        "total_resources": total_resources,
        "project_type_distribution": type_dist,
        "project_status_distribution": status_dist,
        "resource_type_distribution": res_dist,
        "daily_trend": daily_trend,
        "recent_projects": recent_projects,
    }
