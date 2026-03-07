"""
OnlyOffice Document Server 통합 서비스

문서 파일 생성, 등록, 콜백 처리를 담당합니다.
"""

import os
import uuid
import shutil
from datetime import datetime

import httpx
from bson import ObjectId
from services.mongo_service import get_db
from config import settings


async def create_onlyoffice_document(project_id: str, file_type: str, source_path: str) -> dict:
    """파일을 OnlyOffice용으로 등록

    Args:
        project_id: 프로젝트 ID
        file_type: "pptx", "xlsx", "docx"
        source_path: 원본 파일 경로 (uploads/generated/xxx.ext 등)

    Returns:
        onlyoffice_documents 문서
    """
    db = get_db()

    # 파일을 documents 디렉토리로 복사
    doc_dir = os.path.join(settings.UPLOAD_DIR, "documents")
    os.makedirs(doc_dir, exist_ok=True)

    filename = f"{uuid.uuid4().hex}.{file_type}"
    dest_path = os.path.join(doc_dir, filename)

    # source_path가 상대경로인 경우 절대경로로 변환
    if source_path.startswith("/uploads/"):
        abs_source = os.path.join(settings.UPLOAD_DIR, source_path.lstrip("/uploads/"))
    else:
        abs_source = source_path

    if os.path.exists(abs_source):
        shutil.copy2(abs_source, dest_path)
    else:
        raise FileNotFoundError(f"원본 파일을 찾을 수 없습니다: {abs_source}")

    document_key = uuid.uuid4().hex
    now = datetime.utcnow()

    doc = {
        "project_id": project_id,
        "file_path": f"/uploads/documents/{filename}",
        "file_type": file_type,
        "document_key": document_key,
        "version": 1,
        "created_at": now,
        "updated_at": now,
    }

    # upsert: 기존 문서가 있으면 업데이트
    await db.onlyoffice_documents.update_one(
        {"project_id": project_id},
        {"$set": doc},
        upsert=True,
    )

    result = await db.onlyoffice_documents.find_one({"project_id": project_id})
    result["_id"] = str(result["_id"])
    return result


async def update_document_from_callback(document_key: str, download_url: str) -> dict:
    """OnlyOffice 콜백 처리 - 편집된 파일 다운로드 및 업데이트

    Args:
        document_key: OnlyOffice document key
        download_url: OnlyOffice가 제공하는 편집된 파일 다운로드 URL

    Returns:
        업데이트된 onlyoffice_documents 문서
    """
    db = get_db()
    doc = await db.onlyoffice_documents.find_one({"document_key": document_key})
    if not doc:
        raise ValueError(f"문서를 찾을 수 없습니다: key={document_key}")

    # 편집된 파일 다운로드
    file_path = doc["file_path"]
    if file_path.startswith("/uploads/"):
        abs_path = os.path.join(settings.UPLOAD_DIR, file_path[len("/uploads/"):])
    else:
        abs_path = file_path

    async with httpx.AsyncClient(timeout=60.0, verify=False) as client:
        response = await client.get(download_url)
        if response.status_code == 200:
            os.makedirs(os.path.dirname(abs_path), exist_ok=True)
            with open(abs_path, "wb") as f:
                f.write(response.content)
            print(f"[OnlyOffice] 파일 저장 완료: {abs_path} ({len(response.content)} bytes)")
        else:
            print(f"[OnlyOffice] 파일 다운로드 실패: status={response.status_code}")
            raise Exception(f"파일 다운로드 실패: {response.status_code}")

    # 새 document_key 생성 (OnlyOffice 캐시 무효화)
    new_key = uuid.uuid4().hex
    await db.onlyoffice_documents.update_one(
        {"_id": doc["_id"]},
        {
            "$set": {
                "document_key": new_key,
                "updated_at": datetime.utcnow(),
            },
            "$inc": {"version": 1},
        },
    )

    updated = await db.onlyoffice_documents.find_one({"_id": doc["_id"]})
    updated["_id"] = str(updated["_id"])
    return updated


async def get_onlyoffice_document(project_id: str) -> dict | None:
    """프로젝트의 OnlyOffice 문서 정보 조회"""
    db = get_db()
    doc = await db.onlyoffice_documents.find_one({"project_id": project_id})
    if doc:
        doc["_id"] = str(doc["_id"])
    return doc


async def delete_onlyoffice_document(project_id: str):
    """프로젝트의 OnlyOffice 문서 삭제 (파일 포함)"""
    db = get_db()
    doc = await db.onlyoffice_documents.find_one({"project_id": project_id})
    if doc:
        # 물리 파일 삭제
        file_path = doc.get("file_path", "")
        if file_path.startswith("/uploads/"):
            abs_path = os.path.join(settings.UPLOAD_DIR, file_path[len("/uploads/"):])
            if os.path.exists(abs_path):
                os.remove(abs_path)
                print(f"[OnlyOffice] 파일 삭제: {abs_path}")
        await db.onlyoffice_documents.delete_one({"_id": doc["_id"]})
