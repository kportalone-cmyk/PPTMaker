from fastapi import APIRouter, HTTPException, UploadFile, File, Form
from bson import ObjectId
from datetime import datetime
from models.resource import ResourceCreate, WebSearchRequest, URLResourceRequest
from services.mongo_service import get_db
from services.auth_service import decode_jwt_token, extract_user_key, get_user_flexible
from services.search_service import search_web
from services.file_service import extract_text_from_file
from config import settings
import aiofiles
import os
import uuid

router = APIRouter(tags=["resources"])


async def get_user_key(jwt_token: str) -> str:
    """JWT에서 user_key 추출 (내부/외부 JWT 모두 지원)"""
    payload = decode_jwt_token(jwt_token)
    if not payload:
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")
    user_key = extract_user_key(payload)
    if user_key:
        return user_key
    user = await get_user_flexible(payload)
    if user:
        return user.get("ky", "")
    raise HTTPException(status_code=401, detail="사용자를 확인할 수 없습니다")


@router.get("/{jwt_token}/api/resources/content/{resource_id}")
async def get_resource_content(jwt_token: str, resource_id: str):
    """리소스 내용 조회"""
    await get_user_key(jwt_token)
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
    await get_user_key(jwt_token)
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
    await get_user_key(jwt_token)
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
    await get_user_key(jwt_token)

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
    await get_user_key(jwt_token)

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


@router.post("/{jwt_token}/api/resources/urls")
async def add_url_resources(jwt_token: str, data: URLResourceRequest):
    """URL 멀티 리소스 추가 - 각 URL의 텍스트 수집 (YouTube는 자막 추출)"""
    from services.url_service import process_url

    await get_user_key(jwt_token)
    db = get_db()
    resources = []
    errors = []
    now = datetime.utcnow()

    for url in data.urls:
        url = url.strip()
        if not url:
            continue
        try:
            result = await process_url(url)
            if not result:
                continue

            if result.get("error") and not result.get("content"):
                errors.append({"url": url, "error": result["error"]})
                continue

            doc = {
                "project_id": data.project_id,
                "resource_type": result.get("resource_type", "url"),
                "title": result.get("title", url),
                "content": result.get("content", ""),
                "file_path": None,
                "source_url": result.get("source_url", url),
                "sources": [result.get("source_url", url)],
                "created_at": now,
            }
            insert_result = await db.resources.insert_one(doc)
            doc["_id"] = str(insert_result.inserted_id)
            resources.append(doc)
        except Exception as e:
            errors.append({"url": url, "error": str(e)})

    return {"resources": resources, "count": len(resources), "errors": errors}


@router.delete("/{jwt_token}/api/resources/all/{project_id}")
async def delete_all_resources(jwt_token: str, project_id: str):
    """프로젝트의 모든 리소스 삭제"""
    await get_user_key(jwt_token)
    db = get_db()
    # 파일 리소스의 실제 파일 삭제
    cursor = db.resources.find({"project_id": project_id, "file_path": {"$ne": None}})
    async for r in cursor:
        file_path = os.path.join(".", r["file_path"].lstrip("/"))
        if os.path.exists(file_path):
            os.remove(file_path)
    result = await db.resources.delete_many({"project_id": project_id})
    return {"success": True, "deleted_count": result.deleted_count}


@router.delete("/{jwt_token}/api/resources/{resource_id}")
async def delete_resource(jwt_token: str, resource_id: str):
    """리소스 삭제"""
    await get_user_key(jwt_token)
    db = get_db()
    resource = await db.resources.find_one({"_id": ObjectId(resource_id)})
    if resource and resource.get("file_path"):
        file_path = os.path.join(".", resource["file_path"].lstrip("/"))
        if os.path.exists(file_path):
            os.remove(file_path)
    await db.resources.delete_one({"_id": ObjectId(resource_id)})
    return {"success": True}
