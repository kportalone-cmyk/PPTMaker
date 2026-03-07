"""
OnlyOffice Document Server 연동 라우터

에디터 설정 생성, 저장 콜백, 파일 서빙 엔드포인트를 제공합니다.
"""

import os
from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import FileResponse
from jose import jwt as jose_jwt
from bson import ObjectId

from config import settings
from services.mongo_service import get_db
from services.auth_service import get_user_flexible
from services.onlyoffice_service import (
    get_onlyoffice_document,
    update_document_from_callback,
)

router = APIRouter(tags=["onlyoffice"])

# file_type → OnlyOffice documentType 매핑
FILE_TYPE_TO_DOC_TYPE = {
    "pptx": "slide",
    "xlsx": "cell",
    "docx": "word",
}

# file_type → MIME type 매핑
FILE_TYPE_TO_MIME = {
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}


@router.get("/{jwt_token}/api/onlyoffice/{project_id}/config")
async def get_editor_config(jwt_token: str, project_id: str, request: Request):
    """OnlyOffice 에디터 설정 생성 (JWT 서명 포함)"""
    from services.auth_service import decode_jwt_token

    if not settings.ONLYOFFICE_URL:
        raise HTTPException(status_code=500, detail="OnlyOffice 서버가 설정되지 않았습니다")

    payload = decode_jwt_token(jwt_token)
    if not payload:
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")

    user = await get_user_flexible(payload)
    if not user:
        raise HTTPException(status_code=401, detail="사용자를 찾을 수 없습니다")

    # OnlyOffice 문서 정보 조회
    oo_doc = await get_onlyoffice_document(project_id)
    if not oo_doc:
        raise HTTPException(status_code=404, detail="OnlyOffice 문서가 없습니다")

    # 프로젝트 이름 조회
    db = get_db()
    project = await db.projects.find_one({"_id": ObjectId(project_id)}, {"name": 1})
    project_name = project["name"] if project else "문서"

    file_type = oo_doc["file_type"]
    document_type = FILE_TYPE_TO_DOC_TYPE.get(file_type, "word")

    # 서버 공개 URL 구성 (SERVER_BASE_URL 우선, 없으면 request에서 추출)
    if settings.SERVER_BASE_URL:
        base_url = settings.SERVER_BASE_URL.rstrip("/")
    else:
        scheme = request.headers.get("x-forwarded-proto", request.url.scheme)
        host = request.headers.get("x-forwarded-host", request.headers.get("host", ""))
        base_url = f"{scheme}://{host}"

    # 파일 다운로드 URL (OnlyOffice 서버가 접근)
    file_url = f"{base_url}/api/onlyoffice/file/{project_id}/{os.path.basename(oo_doc['file_path'])}"

    # 콜백 URL (OnlyOffice 서버가 저장 시 호출)
    callback_url = f"{base_url}/api/onlyoffice/callback"

    print(f"[OnlyOffice] file_url={file_url}, callback_url={callback_url}")

    config = {
        "documentType": document_type,
        "document": {
            "fileType": file_type,
            "key": oo_doc["document_key"],
            "title": f"{project_name}.{file_type}",
            "url": file_url,
            "permissions": {
                "comment": True,
                "copy": True,
                "download": True,
                "edit": True,
                "print": True,
                "review": False,
            },
        },
        "editorConfig": {
            "callbackUrl": callback_url,
            "mode": "edit",
            "lang": "ko",
            "user": {
                "id": user.get("ky", ""),
                "name": user.get("nm", "사용자"),
            },
            "customization": {
                "autosave": True,
                "forcesave": True,
                "compactHeader": True,
            },
        },
    }

    # JWT 서명 (OnlyOffice JWT Secret이 설정된 경우)
    if settings.ONLYOFFICE_JWT_SECRET:
        token = jose_jwt.encode(
            config,
            settings.ONLYOFFICE_JWT_SECRET,
            algorithm="HS256",
        )
        config["token"] = token

    return {
        "config": config,
        "onlyoffice_url": settings.ONLYOFFICE_URL,
    }


@router.post("/api/onlyoffice/callback")
async def onlyoffice_callback(request: Request):
    """OnlyOffice 저장 콜백 처리

    OnlyOffice Document Server가 문서 편집 상태 변경 시 호출합니다.
    status 2: 문서 저장 완료 (모든 사용자가 에디터를 닫은 후)
    status 6: 강제 저장 (편집 중 저장)
    """
    try:
        body = await request.json()
    except Exception:
        return {"error": 0}

    status = body.get("status")
    key = body.get("key")
    download_url = body.get("url")

    print(f"[OnlyOffice Callback] status={status}, key={key}")

    if status in (2, 6) and download_url:
        # 문서 저장: 편집된 파일 다운로드 및 업데이트
        try:
            await update_document_from_callback(key, download_url)
            print(f"[OnlyOffice Callback] 파일 저장 완료: key={key}")
        except Exception as e:
            print(f"[OnlyOffice Callback] 저장 실패: {e}")
            import traceback
            traceback.print_exc()
    elif status == 4:
        # 변경 없이 닫힘
        print(f"[OnlyOffice Callback] 변경 없이 닫힘: key={key}")

    # OnlyOffice는 반드시 {"error": 0}을 받아야 함
    return {"error": 0}


@router.get("/api/onlyoffice/diag/{project_id}")
async def onlyoffice_diagnostic(project_id: str):
    """OnlyOffice 파일 URL 접근성 진단 (디버그용)"""
    import httpx
    db = get_db()
    oo_doc = await db.onlyoffice_documents.find_one({"project_id": project_id})
    if not oo_doc:
        return {"error": "onlyoffice_documents에 문서 없음"}

    file_path = oo_doc.get("file_path", "")
    if file_path.startswith("/uploads/"):
        abs_path = os.path.join(settings.UPLOAD_DIR, file_path[len("/uploads/"):])
    else:
        abs_path = file_path

    base_url = settings.SERVER_BASE_URL.rstrip("/") if settings.SERVER_BASE_URL else "(미설정)"
    file_url = f"{base_url}/api/onlyoffice/file/{project_id}/{os.path.basename(file_path)}" if settings.SERVER_BASE_URL else "(SERVER_BASE_URL 미설정)"

    result = {
        "project_id": project_id,
        "file_path_db": file_path,
        "abs_path": abs_path,
        "file_exists": os.path.exists(abs_path),
        "file_size": os.path.getsize(abs_path) if os.path.exists(abs_path) else 0,
        "file_url": file_url,
        "server_base_url": settings.SERVER_BASE_URL or "(미설정)",
        "onlyoffice_url": settings.ONLYOFFICE_URL,
        "document_key": oo_doc.get("document_key", ""),
    }

    # file_url 셀프 테스트
    if settings.SERVER_BASE_URL:
        try:
            async with httpx.AsyncClient(timeout=10.0, verify=False) as client:
                resp = await client.head(file_url)
                result["self_test_status"] = resp.status_code
                result["self_test_ok"] = resp.status_code == 200
        except Exception as e:
            result["self_test_status"] = str(e)
            result["self_test_ok"] = False

    return result


@router.get("/api/onlyoffice/file/{project_id}/{filename}")
async def serve_onlyoffice_file(project_id: str, filename: str, request: Request):
    """OnlyOffice 서버가 파일을 다운로드하는 엔드포인트

    인증 없음 (OnlyOffice 서버가 직접 호출)
    """
    print(f"[OnlyOffice File] 요청: project={project_id}, file={filename}, from={request.client.host if request.client else 'unknown'}")
    db = get_db()
    oo_doc = await db.onlyoffice_documents.find_one({"project_id": project_id})
    if not oo_doc:
        print(f"[OnlyOffice File] 404: 문서 없음 project={project_id}")
        raise HTTPException(status_code=404, detail="문서를 찾을 수 없습니다")

    file_path = oo_doc.get("file_path", "")
    if file_path.startswith("/uploads/"):
        abs_path = os.path.join(settings.UPLOAD_DIR, file_path[len("/uploads/"):])
    else:
        abs_path = file_path

    if not os.path.exists(abs_path):
        print(f"[OnlyOffice File] 404: 파일 없음 path={abs_path}")
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다")

    print(f"[OnlyOffice File] 서빙: {abs_path} ({os.path.getsize(abs_path)} bytes)")

    file_type = oo_doc.get("file_type", "docx")
    media_type = FILE_TYPE_TO_MIME.get(file_type, "application/octet-stream")

    return FileResponse(
        abs_path,
        media_type=media_type,
        filename=filename,
    )
