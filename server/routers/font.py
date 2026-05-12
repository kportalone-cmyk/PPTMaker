import os
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

import aiofiles
from bson import ObjectId
from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from pydantic import BaseModel

from config import settings
from services import redis_service
from services.font_service import generate_font_preview
from services.mongo_service import get_db

router = APIRouter(tags=["fonts"])

_FONTS_CACHE_KEY = "fonts:all"
_FONTS_PUBLIC_CACHE_KEY = "fonts:public"
_FONTS_TTL = 604800  # 7일 (폰트는 거의 변경되지 않음)

# 폰트 업로드 허용 확장자
_ALLOWED_FONT_EXT = {".ttf", ".otf", ".ttc", ".woff", ".woff2"}


class FontCreate(BaseModel):
    name: str
    family: str
    url: str = ""


async def _invalidate_fonts_cache():
    """폰트 캐시 무효화"""
    await redis_service.cache_delete(_FONTS_CACHE_KEY, _FONTS_PUBLIC_CACHE_KEY)


def _get_fonts_dir() -> Path:
    """업로드된 폰트 파일을 저장할 디렉토리. 없으면 생성."""
    fonts_dir = Path(settings.UPLOAD_DIR) / "fonts"
    fonts_dir.mkdir(parents=True, exist_ok=True)
    return fonts_dir


def _safe_unlink(path_str: Optional[str]):
    """파일이 존재하면 삭제. 실패해도 예외 던지지 않음."""
    if not path_str:
        return
    try:
        p = Path(path_str)
        if p.exists() and p.is_file():
            p.unlink()
    except Exception as e:
        print(f"[Fonts] 파일 삭제 실패 {path_str}: {e}")


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
    """폰트 등록 (URL 기반 - 외부 호스팅 폰트)

    기존 호환 유지용. 폰트 파일을 직접 업로드하려면
    `POST /{jwt}/api/admin/fonts/upload` 를 사용하세요.
    """
    db = get_db()
    doc = {"name": body.name, "family": body.family, "url": body.url}
    try:
        result = await db.fonts.insert_one(doc)
        doc["_id"] = str(result.inserted_id)
        await _invalidate_fonts_cache()
        return {"font": doc}
    except Exception:
        raise HTTPException(status_code=409, detail="이미 등록된 폰트입니다")


async def _save_one_font(file: UploadFile) -> dict:
    """단일 폰트 파일을 디스크에 저장하고 글리프 미리보기를 생성한 뒤 DB에 등록.

    실패 시 HTTPException 발생. 호출자가 try/except 로 한 건씩 처리할 수 있도록 설계.
    이름 / family 는 파일명에서 자동 추출 (업로드 후 사용자가 수정 가능).
    """
    original_filename = file.filename or "upload"
    ext = os.path.splitext(original_filename)[1].lower()
    if ext not in _ALLOWED_FONT_EXT:
        raise HTTPException(
            status_code=400,
            detail=f"지원하지 않는 폰트 형식입니다 (허용: {', '.join(sorted(_ALLOWED_FONT_EXT))})",
        )

    base_name = os.path.splitext(original_filename)[0] or "Font"
    name = base_name
    family = base_name

    fonts_dir = _get_fonts_dir()
    file_id = uuid.uuid4().hex
    filename = f"{file_id}{ext}"
    save_path = fonts_dir / filename
    preview_filename = f"{file_id}_preview.png"
    preview_path = fonts_dir / preview_filename

    max_size = settings.MAX_UPLOAD_SIZE
    bytes_written = 0
    try:
        async with aiofiles.open(save_path, "wb") as out:
            while True:
                chunk = await file.read(1024 * 1024)  # 1MB
                if not chunk:
                    break
                bytes_written += len(chunk)
                if bytes_written > max_size:
                    await out.close()
                    _safe_unlink(str(save_path))
                    raise HTTPException(
                        status_code=413,
                        detail=f"파일 크기가 제한({max_size} bytes)을 초과합니다",
                    )
                await out.write(chunk)
    except HTTPException:
        raise
    except Exception as e:
        _safe_unlink(str(save_path))
        raise HTTPException(status_code=500, detail=f"파일 저장 실패: {e}")

    # 글리프 미리보기 PNG 생성 (실패해도 등록은 계속 진행)
    preview_url: Optional[str] = None
    preview_local_path: Optional[str] = None
    try:
        await generate_font_preview(str(save_path), str(preview_path), family=family)
        if preview_path.exists():
            preview_url = f"/uploads/fonts/{preview_filename}"
            preview_local_path = str(preview_path)
    except Exception as e:
        print(f"[Fonts] 미리보기 생성 실패 ({original_filename}): {e}")

    doc = {
        "name": name,
        "family": family,
        "url": f"/uploads/fonts/{filename}",
        "local_path": str(save_path),
        "preview_url": preview_url,
        "preview_local_path": preview_local_path,
        "original_filename": original_filename,
        "file_size": bytes_written,
        "uploaded_at": datetime.utcnow(),
    }
    try:
        db = get_db()
        result = await db.fonts.insert_one(doc)
        doc["_id"] = str(result.inserted_id)
    except Exception as e:
        _safe_unlink(str(save_path))
        _safe_unlink(str(preview_path))
        raise HTTPException(status_code=500, detail=f"폰트 도큐먼트 저장 실패: {e}")

    return doc


@router.post("/{jwt_token}/api/admin/fonts/upload")
async def upload_fonts(
    jwt_token: str,
    files: list[UploadFile] = File(...),
):
    """폰트 파일 직접 업로드 (다중 파일 지원).

    - 확장자: .ttf / .otf / .ttc / .woff / .woff2
    - 파일 크기 제한: settings.MAX_UPLOAD_SIZE (파일 별 적용)
    - 저장 위치: {UPLOAD_DIR}/fonts/<uuid>.<ext>
    - 글리프 미리보기 PNG: <UPLOAD_DIR>/fonts/<uuid>_preview.png (PIL 자동 생성)
    - 각 파일의 이름/패밀리는 파일명에서 자동 추출. 업로드 후 별도 PUT 으로 수정 가능
    - 응답:
        {
          "results":  [{"success": bool, "filename": str, "font?": {...}, "error?": str}, ...],
          "uploaded": [{...font doc...}, ...],   # 성공한 도큐먼트만
          "success_count": int,
          "failed_count":  int
        }
    """
    if not files:
        raise HTTPException(status_code=400, detail="업로드할 파일이 없습니다")

    results: list[dict] = []
    uploaded: list[dict] = []

    for f in files:
        original_filename = f.filename or "upload"
        try:
            doc = await _save_one_font(f)
            results.append({"success": True, "filename": original_filename, "font": doc})
            uploaded.append(doc)
        except HTTPException as e:
            results.append({
                "success": False,
                "filename": original_filename,
                "error": e.detail if hasattr(e, "detail") else str(e),
                "status": e.status_code,
            })
        except Exception as e:
            results.append({
                "success": False,
                "filename": original_filename,
                "error": str(e),
            })

    # 단 한 개라도 성공했다면 캐시 무효화 (목록 새로고침)
    if uploaded:
        await _invalidate_fonts_cache()

    return {
        "results": results,
        "uploaded": uploaded,
        "success_count": len(uploaded),
        "failed_count": len(results) - len(uploaded),
    }


@router.delete("/{jwt_token}/api/admin/fonts/{font_id}")
async def delete_font(jwt_token: str, font_id: str):
    """폰트 삭제 (도큐먼트 + 디스크 파일)"""
    db = get_db()
    try:
        oid = ObjectId(font_id)
    except Exception:
        raise HTTPException(status_code=400, detail="유효하지 않은 폰트 ID 입니다")

    doc = await db.fonts.find_one({"_id": oid})
    if doc:
        # 업로드 파일이면 디스크에서도 제거
        _safe_unlink(doc.get("local_path"))
        _safe_unlink(doc.get("preview_local_path"))

    await db.fonts.delete_one({"_id": oid})
    await _invalidate_fonts_cache()
    return {"success": True}
