from fastapi import APIRouter, HTTPException, UploadFile, File, Form
from bson import ObjectId
from datetime import datetime
from models.resource import ResourceCreate, WebSearchRequest
from services.mongo_service import get_db
from services.auth_service import decode_jwt_token
from services.search_service import search_web
from services.file_service import extract_text_from_file
from config import settings
import aiofiles
import os
import uuid

router = APIRouter(tags=["resources"])


def get_user_key(jwt_token: str) -> str:
    payload = decode_jwt_token(jwt_token)
    if not payload:
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")
    return payload.get("user_key", "")


@router.get("/{jwt_token}/api/resources/content/{resource_id}")
async def get_resource_content(jwt_token: str, resource_id: str):
    """리소스 내용 조회"""
    get_user_key(jwt_token)
    db = get_db()
    resource = await db.resources.find_one({"_id": ObjectId(resource_id)})
    if not resource:
        raise HTTPException(status_code=404, detail="리소스를 찾을 수 없습니다")
    return {
        "resource_id": str(resource["_id"]),
        "title": resource.get("title", ""),
        "resource_type": resource.get("resource_type", ""),
        "content": resource.get("content", ""),
        "sources": resource.get("sources", []),
    }


@router.get("/{jwt_token}/api/resources/{project_id}")
async def list_resources(jwt_token: str, project_id: str):
    """프로젝트 리소스 목록"""
    get_user_key(jwt_token)
    db = get_db()
    cursor = db.resources.find({"project_id": project_id}).sort("created_at", -1)
    resources = []
    async for r in cursor:
        r["_id"] = str(r["_id"])
        resources.append(r)
    return {"resources": resources}


@router.post("/{jwt_token}/api/resources/text")
async def add_text_resource(jwt_token: str, data: ResourceCreate):
    """텍스트 리소스 추가 (복사 붙여넣기) - 원본 그대로 저장"""
    get_user_key(jwt_token)
    db = get_db()
    doc = {
        "project_id": data.project_id,
        "resource_type": "text",
        "title": data.title or "텍스트 리소스",
        "content": data.content,
        "file_path": None,
        "source_url": None,
        "created_at": datetime.utcnow(),
    }
    result = await db.resources.insert_one(doc)
    doc["_id"] = str(result.inserted_id)
    return {"resource": doc}


@router.post("/{jwt_token}/api/resources/file")
async def upload_file_resource(
    jwt_token: str,
    project_id: str = Form(...),
    title: str = Form(""),
    file: UploadFile = File(...)
):
    """파일 리소스 업로드 - 텍스트 추출 후 마크다운 형식으로 저장"""
    get_user_key(jwt_token)

    # 허용 확장자 확인
    allowed_ext = {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".pdf", ".txt", ".csv"}
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in allowed_ext:
        raise HTTPException(status_code=400, detail=f"허용되지 않은 파일 형식입니다: {ext}")

    filename = f"{uuid.uuid4().hex}{ext}"
    save_path = os.path.join(settings.UPLOAD_DIR, "resources", filename)
    os.makedirs(os.path.dirname(save_path), exist_ok=True)

    async with aiofiles.open(save_path, "wb") as f:
        content = await file.read()
        await f.write(content)

    # 파일에서 텍스트 추출 (마크다운 형식)
    extracted_content = extract_text_from_file(save_path, ext)

    db = get_db()
    doc = {
        "project_id": project_id,
        "resource_type": "file",
        "title": title or file.filename,
        "content": extracted_content,
        "file_path": f"/uploads/resources/{filename}",
        "original_filename": file.filename,
        "source_url": None,
        "created_at": datetime.utcnow(),
    }
    result = await db.resources.insert_one(doc)
    doc["_id"] = str(result.inserted_id)
    return {"resource": doc}


@router.post("/{jwt_token}/api/resources/web-search")
async def add_web_search_resource(jwt_token: str, data: WebSearchRequest):
    """웹 검색 후 각 출처 페이지의 원본 콘텐츠를 개별 리소스로 저장"""
    get_user_key(jwt_token)

    search_result = await search_web(data.query)

    error = search_result.get("error")
    pages = search_result.get("pages", [])

    if error and not pages:
        raise HTTPException(status_code=500, detail=error)

    db = get_db()
    resources = []
    now = datetime.utcnow()

    for i, page in enumerate(pages):
        doc = {
            "project_id": data.project_id,
            "resource_type": "web",
            "title": page.get("title") or f"웹 검색 {i + 1}: {data.query}",
            "content": page.get("content", ""),
            "file_path": None,
            "source_url": page.get("url", ""),
            "sources": [page.get("url", "")],
            "created_at": now,
        }
        result = await db.resources.insert_one(doc)
        doc["_id"] = str(result.inserted_id)
        resources.append(doc)

    return {"resources": resources, "count": len(resources)}


@router.delete("/{jwt_token}/api/resources/{resource_id}")
async def delete_resource(jwt_token: str, resource_id: str):
    """리소스 삭제"""
    get_user_key(jwt_token)
    db = get_db()
    resource = await db.resources.find_one({"_id": ObjectId(resource_id)})
    if resource and resource.get("file_path"):
        file_path = os.path.join(".", resource["file_path"].lstrip("/"))
        if os.path.exists(file_path):
            os.remove(file_path)
    await db.resources.delete_one({"_id": ObjectId(resource_id)})
    return {"success": True}
