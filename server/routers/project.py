import os
from fastapi import APIRouter, HTTPException
from bson import ObjectId
from datetime import datetime
from models.project import ProjectCreate, ProjectUpdate
from services.mongo_service import get_db, get_org_db
from services.auth_service import decode_jwt_token, extract_user_key, get_user_flexible
from services import redis_service
from routers.collaboration import check_project_access

router = APIRouter(tags=["projects"])


async def get_user_key(jwt_token: str) -> str:
    """JWT에서 user_key 추출 (내부/외부 JWT 모두 지원)"""
    payload = decode_jwt_token(jwt_token)
    if not payload:
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")
    # extract_user_key로 user_key 또는 email 추출
    user_key = extract_user_key(payload)
    if user_key:
        return user_key
    # userid → em 매칭으로 사용자 조회 후 ky 반환
    user = await get_user_flexible(payload)
    if user:
        return user.get("ky", "")
    raise HTTPException(status_code=401, detail="사용자를 확인할 수 없습니다")


@router.get("/{jwt_token}/api/projects")
async def list_projects(jwt_token: str):
    """사용자 프로젝트 목록 조회 (소유 + 공유)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()

    # 내 프로젝트
    cursor = db.projects.find({"user_key": user_key}).sort("created_at", -1)
    projects = []
    async for p in cursor:
        pid = str(p["_id"])
        p["_id"] = pid
        p["_collab_role"] = "owner"
        p["_collab_count"] = await db.collaborators.count_documents({"project_id": pid})
        projects.append(p)

    # 공유된 프로젝트
    org_db = get_org_db()
    shared_projects = []
    async for collab in db.collaborators.find({"user_key": user_key}):
        project_id = collab.get("project_id")
        p = await db.projects.find_one({"_id": ObjectId(project_id)})
        if p:
            p["_id"] = str(p["_id"])
            p["_collab_role"] = collab.get("role", "viewer")
            p["_collab_count"] = await db.collaborators.count_documents({"project_id": project_id})
            # 소유자 이름 조회
            owner_key = p.get("user_key", "")
            if owner_key:
                owner = await org_db.user_info.find_one({"ky": owner_key}, {"nm": 1, "dp": 1})
                if owner:
                    p["_owner_name"] = owner.get("nm", "")
                    p["_owner_dept"] = owner.get("dp", "")
            shared_projects.append(p)

    return {"projects": projects, "shared_projects": shared_projects}


@router.post("/{jwt_token}/api/projects")
async def create_project(jwt_token: str, data: ProjectCreate):
    """프로젝트 생성"""
    user_key = await get_user_key(jwt_token)
    db = get_db()
    doc = {
        "name": data.name,
        "description": data.description,
        "template_id": data.template_id if data.project_type in ("slide", "onlyoffice_pptx") else None,
        "project_type": data.project_type,
        "instructions": "",
        "user_key": user_key,
        "status": "draft",
        "created_at": datetime.utcnow(),
        "updated_at": datetime.utcnow(),
    }
    result = await db.projects.insert_one(doc)
    doc["_id"] = str(result.inserted_id)
    return {"project": doc}


@router.get("/{jwt_token}/api/projects/{project_id}")
async def get_project(jwt_token: str, project_id: str):
    """프로젝트 상세 조회 (소유자 또는 협업자)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()

    # 소유자 또는 협업자 접근 권한 확인
    access = await check_project_access(db, project_id, user_key, "viewer")
    project = access["project"]
    project["_collab_role"] = access["role"]

    # 리소스 목록
    cursor = db.resources.find({"project_id": project_id})
    resources = []
    async for r in cursor:
        r["_id"] = str(r["_id"])
        resources.append(r)

    # 생성된 슬라이드
    slide_cursor = db.generated_slides.find({"project_id": project_id}).sort("order", 1)
    generated_slides = []
    async for s in slide_cursor:
        s["_id"] = str(s["_id"])
        generated_slides.append(s)

    # 엑셀 데이터 조회 (엑셀 프로젝트인 경우)
    generated_excel = None
    if project.get("project_type") in ("excel", "onlyoffice_xlsx"):
        generated_excel = await db.generated_excel.find_one({"project_id": project_id})
        if generated_excel:
            generated_excel["_id"] = str(generated_excel["_id"])

    # OnlyOffice 문서 정보 조회
    onlyoffice_doc = None
    project_type = project.get("project_type", "")
    if project_type.startswith("onlyoffice_"):
        onlyoffice_doc = await db.onlyoffice_documents.find_one({"project_id": project_id})
        if onlyoffice_doc:
            onlyoffice_doc["_id"] = str(onlyoffice_doc["_id"])

    # Word 문서 데이터 조회
    generated_docx = None
    if project_type in ("word", "onlyoffice_docx"):
        generated_docx = await db.generated_docx.find_one({"project_id": project_id})
        if generated_docx:
            generated_docx["_id"] = str(generated_docx["_id"])

    return {
        "project": project,
        "resources": resources,
        "generated_slides": generated_slides,
        "generated_excel": generated_excel,
        "onlyoffice_doc": onlyoffice_doc,
        "generated_docx": generated_docx,
    }


@router.put("/{jwt_token}/api/projects/{project_id}")
async def update_project(jwt_token: str, project_id: str, data: ProjectUpdate):
    """프로젝트 수정"""
    user_key = await get_user_key(jwt_token)
    db = get_db()
    update_data = {k: v for k, v in data.dict().items() if v is not None}
    update_data["updated_at"] = datetime.utcnow()
    result = await db.projects.update_one(
        {"_id": ObjectId(project_id), "user_key": user_key},
        {"$set": update_data}
    )
    if result.matched_count == 0:
        raise HTTPException(status_code=404, detail="프로젝트를 찾을 수 없습니다")
    return {"success": True}


@router.post("/{jwt_token}/api/projects/{project_id}/reset")
async def reset_project(jwt_token: str, project_id: str):
    """프로젝트 초기화 (생성 슬라이드 삭제 + 상태 초기화, 리소스 유지)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()

    project = await db.projects.find_one({
        "_id": ObjectId(project_id),
        "user_key": user_key
    })
    if not project:
        raise HTTPException(status_code=404, detail="프로젝트를 찾을 수 없습니다")

    # 생성된 슬라이드/엑셀/문서 삭제 (리소스는 유지)
    await db.generated_slides.delete_many({"project_id": project_id})
    await db.generated_excel.delete_many({"project_id": project_id})
    await db.generated_docx.delete_many({"project_id": project_id})
    # OnlyOffice 문서 정리 (파일 포함)
    from services.onlyoffice_service import delete_onlyoffice_document
    await delete_onlyoffice_document(project_id)
    # 협업 데이터 정리 (Lock/히스토리 리셋)
    await db.slide_locks.delete_many({"project_id": project_id})
    await db.slide_history.delete_many({"project_id": project_id})
    # Redis 잠금/상태/취소 플래그 정리
    await redis_service.delete_project_locks(project_id)
    await redis_service.clear_generation_cancel(project_id)
    await redis_service.cache_delete_pattern(f"presence:{project_id}")

    # 프로젝트 상태 초기화 (지침과 입력 내용은 유지)
    await db.projects.update_one(
        {"_id": ObjectId(project_id)},
        {"$set": {
            "status": "draft",
            "template_id": None,
            "updated_at": datetime.utcnow()
        }}
    )

    return {"success": True}


@router.delete("/{jwt_token}/api/projects/{project_id}")
async def delete_project(jwt_token: str, project_id: str):
    """프로젝트 삭제 (리소스, 생성된 슬라이드도 삭제)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()
    result = await db.projects.delete_one({
        "_id": ObjectId(project_id),
        "user_key": user_key
    })
    if result.deleted_count == 0:
        raise HTTPException(status_code=404, detail="프로젝트를 찾을 수 없습니다")
    await db.resources.delete_many({"project_id": project_id})
    await db.generated_slides.delete_many({"project_id": project_id})
    await db.generated_excel.delete_many({"project_id": project_id})
    await db.generated_docx.delete_many({"project_id": project_id})
    # OnlyOffice 문서 정리 (파일 포함)
    from services.onlyoffice_service import delete_onlyoffice_document
    await delete_onlyoffice_document(project_id)
    # 협업 데이터 정리
    await db.collaborators.delete_many({"project_id": project_id})
    await db.slide_locks.delete_many({"project_id": project_id})
    await db.slide_history.delete_many({"project_id": project_id})
    await db.online_presence.delete_many({"project_id": project_id})
    # Redis 잠금/상태/취소 플래그 정리
    await redis_service.delete_project_locks(project_id)
    await redis_service.clear_generation_cancel(project_id)
    await redis_service.remove_presence(project_id, user_key)
    return {"success": True}
