"""파워포인트 스타일(PPT Style) 관리 API

관리자 전용 — 모든 엔드포인트는 JWT 경로 prefix를 포함하며
admin 권한 검증을 수행한다.

기존 HTML 리포트(html_skills) 백엔드는 건드리지 않는다.
"""

import os
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

import aiofiles
from bson import ObjectId
from fastapi import APIRouter, HTTPException, UploadFile, File

from config import settings
from models.ppt_style import (
    PPTStyleCreate,
    PPTStyleUpdate,
    PPTStyleFontLink,
    build_new_style_doc,
)
from services.mongo_service import get_db
from services.auth_service import decode_jwt_token, get_user_flexible, is_admin
from services import ppt_style_service


router = APIRouter(tags=["ppt-styles"])


# ============ 관리자 권한 검증 ============

async def verify_admin(jwt_token: str):
    """관리자 권한 검증 (template.py 패턴과 동일)"""
    payload = decode_jwt_token(jwt_token)
    if not payload:
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")
    user = await get_user_flexible(payload)
    if not user or not is_admin(user):
        raise HTTPException(status_code=403, detail="관리자 권한이 필요합니다")
    return user


async def verify_user(jwt_token: str):
    """일반 사용자 JWT 검증 (admin 권한 불필요).

    공개 라우트(/api/ppt-styles)용 — JWT 디코딩만 검증하면 충분.
    """
    payload = decode_jwt_token(jwt_token)
    if not payload:
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")
    return payload


def _serialize(doc: dict) -> dict:
    if not doc:
        return doc
    out = dict(doc)
    if "_id" in out and not isinstance(out["_id"], str):
        out["_id"] = str(out["_id"])
    return out


def _to_oid(style_id: str) -> ObjectId:
    try:
        return ObjectId(style_id)
    except Exception:
        raise HTTPException(status_code=400, detail="유효하지 않은 style_id 입니다")


# ============ CRUD ============

@router.get("/{jwt_token}/api/admin/ppt-styles")
async def list_ppt_styles(jwt_token: str):
    """PPT 스타일 목록 조회 (최신순)"""
    await verify_admin(jwt_token)
    db = get_db()
    cursor = db.ppt_styles.find().sort("created_at", -1)
    styles = []
    async for s in cursor:
        styles.append(_serialize(s))
    return {"styles": styles}


@router.post("/{jwt_token}/api/admin/ppt-styles")
async def create_ppt_style(jwt_token: str, data: PPTStyleCreate):
    """PPT 스타일 생성 (title만 필수, 나머지는 기본값 자동 설정)"""
    user = await verify_admin(jwt_token)
    if not data.title or not data.title.strip():
        raise HTTPException(status_code=400, detail="title은 필수 입력입니다")

    db = get_db()
    doc = build_new_style_doc(data, user_key=user.get("ky", ""))
    result = await db.ppt_styles.insert_one(doc)
    doc["_id"] = str(result.inserted_id)
    return {"style": doc}


@router.get("/{jwt_token}/api/admin/ppt-styles/{style_id}")
async def get_ppt_style(jwt_token: str, style_id: str):
    """PPT 스타일 단일 조회"""
    await verify_admin(jwt_token)
    db = get_db()
    doc = await db.ppt_styles.find_one({"_id": _to_oid(style_id)})
    if not doc:
        raise HTTPException(status_code=404, detail="스타일을 찾을 수 없습니다")
    return {"style": _serialize(doc)}


@router.put("/{jwt_token}/api/admin/ppt-styles/{style_id}")
async def update_ppt_style(jwt_token: str, style_id: str, data: PPTStyleUpdate):
    """PPT 스타일 전체 수정 (exclude_unset)"""
    await verify_admin(jwt_token)
    db = get_db()
    update_data = data.dict(exclude_unset=True)
    if not update_data:
        raise HTTPException(status_code=400, detail="수정할 필드가 없습니다")
    update_data["updated_at"] = datetime.utcnow()

    result = await db.ppt_styles.update_one(
        {"_id": _to_oid(style_id)},
        {"$set": update_data},
    )
    if result.matched_count == 0:
        raise HTTPException(status_code=404, detail="스타일을 찾을 수 없습니다")
    doc = await db.ppt_styles.find_one({"_id": _to_oid(style_id)})
    return {"style": _serialize(doc)}


@router.delete("/{jwt_token}/api/admin/ppt-styles/{style_id}")
async def delete_ppt_style(jwt_token: str, style_id: str):
    """PPT 스타일 삭제 (샘플 이미지 디렉토리 포함)"""
    await verify_admin(jwt_token)
    db = get_db()
    oid = _to_oid(style_id)
    existing = await db.ppt_styles.find_one({"_id": oid})
    if not existing:
        raise HTTPException(status_code=404, detail="스타일을 찾을 수 없습니다")

    # 샘플 파일 디렉토리 통째로 정리
    removed_files = ppt_style_service.delete_all_samples(style_id)

    await db.ppt_styles.delete_one({"_id": oid})
    return {"success": True, "removed_files": removed_files}


@router.put("/{jwt_token}/api/admin/ppt-styles/{style_id}/publish")
async def toggle_ppt_style_publish(jwt_token: str, style_id: str):
    """게시/비게시 토글 (body 없이)"""
    await verify_admin(jwt_token)
    db = get_db()
    oid = _to_oid(style_id)
    existing = await db.ppt_styles.find_one({"_id": oid})
    if not existing:
        raise HTTPException(status_code=404, detail="스타일을 찾을 수 없습니다")

    new_status = not bool(existing.get("is_published", False))
    await db.ppt_styles.update_one(
        {"_id": oid},
        {"$set": {"is_published": new_status, "updated_at": datetime.utcnow()}},
    )
    return {"success": True, "is_published": new_status}


# ============ 샘플 이미지 ============

_SUPPORTED_IMAGE_EXT = {".jpg", ".jpeg", ".png", ".gif", ".webp"}


@router.post("/{jwt_token}/api/admin/ppt-styles/{style_id}/samples")
async def upload_ppt_style_samples(
    jwt_token: str,
    style_id: str,
    files: list[UploadFile] = File(...),
):
    """샘플 이미지 다중 업로드 (multipart, field name: files)

    저장 위치: {UPLOAD_DIR}/ppt_style_samples/{style_id}/<uuid>.<ext>
    URL: /uploads/ppt_style_samples/{style_id}/<filename>
    """
    await verify_admin(jwt_token)
    db = get_db()
    oid = _to_oid(style_id)
    existing = await db.ppt_styles.find_one({"_id": oid})
    if not existing:
        raise HTTPException(status_code=404, detail="스타일을 찾을 수 없습니다")

    if not files:
        raise HTTPException(status_code=400, detail="업로드된 파일이 없습니다")

    sample_dir = ppt_style_service.get_sample_dir(style_id)
    sample_dir.mkdir(parents=True, exist_ok=True)

    added: list[dict] = []
    for f in files:
        original_filename = f.filename or "upload"
        ext = os.path.splitext(original_filename)[1].lower()
        if ext not in _SUPPORTED_IMAGE_EXT:
            # 지원하지 않는 형식은 스킵 (전체 실패시키지 않음)
            continue

        file_id = uuid.uuid4().hex
        filename = f"{file_id}{ext}"
        save_path = sample_dir / filename

        async with aiofiles.open(save_path, "wb") as out:
            content = await f.read()
            await out.write(content)

        added.append({
            "file_id": file_id,
            "url": ppt_style_service.sample_url(style_id, filename),
            "filename": filename,
            "original_filename": original_filename,
            "uploaded_at": datetime.utcnow(),
        })

    if not added:
        raise HTTPException(status_code=400, detail="지원하는 이미지 형식이 없습니다 (jpg/png/gif/webp)")

    await db.ppt_styles.update_one(
        {"_id": oid},
        {
            "$push": {"sample_image_refs": {"$each": added}},
            "$set": {"updated_at": datetime.utcnow()},
        },
    )
    return {"added": added}


@router.delete("/{jwt_token}/api/admin/ppt-styles/{style_id}/samples/{file_id}")
async def delete_ppt_style_sample(jwt_token: str, style_id: str, file_id: str):
    """개별 샘플 이미지 삭제 (DB 레퍼런스 + 디스크 파일)"""
    await verify_admin(jwt_token)
    db = get_db()
    oid = _to_oid(style_id)
    existing = await db.ppt_styles.find_one({"_id": oid})
    if not existing:
        raise HTTPException(status_code=404, detail="스타일을 찾을 수 없습니다")

    refs: list[dict] = existing.get("sample_image_refs", []) or []
    target = next((r for r in refs if r.get("file_id") == file_id), None)
    if not target:
        raise HTTPException(status_code=404, detail="샘플을 찾을 수 없습니다")

    filename = target.get("filename")
    if not filename:
        # URL 끝의 basename 으로 폴백
        url = target.get("url", "")
        filename = os.path.basename(url) if url else None
    if filename:
        ppt_style_service.delete_sample_file(style_id, filename)

    await db.ppt_styles.update_one(
        {"_id": oid},
        {
            "$pull": {"sample_image_refs": {"file_id": file_id}},
            "$set": {"updated_at": datetime.utcnow()},
        },
    )
    return {"success": True}


# ============ 폰트 연결 ============

@router.post("/{jwt_token}/api/admin/ppt-styles/{style_id}/fonts")
async def link_ppt_style_fonts(
    jwt_token: str,
    style_id: str,
    body: PPTStyleFontLink,
):
    """폰트 ID 목록 연결 (body: {font_ids: [...]})

    fonts 컬렉션에 실재하는 ObjectId만 유지한다.
    """
    await verify_admin(jwt_token)
    db = get_db()
    oid = _to_oid(style_id)
    existing = await db.ppt_styles.find_one({"_id": oid})
    if not existing:
        raise HTTPException(status_code=404, detail="스타일을 찾을 수 없습니다")

    # 유효한 fonts._id 만 추리기
    valid_ids: list[str] = []
    for fid in body.font_ids or []:
        try:
            oid_f = ObjectId(fid)
        except Exception:
            continue
        if await db.fonts.find_one({"_id": oid_f}):
            valid_ids.append(fid)

    await db.ppt_styles.update_one(
        {"_id": oid},
        {"$set": {"font_refs": valid_ids, "updated_at": datetime.utcnow()}},
    )
    return {"success": True, "font_refs": valid_ids}


# ============ Vision 분석 ============

@router.post("/{jwt_token}/api/admin/ppt-styles/{style_id}/analyze")
async def analyze_ppt_style_samples(jwt_token: str, style_id: str):
    """샘플 이미지 자동 Vision 분석 + 패턴 추출 (M8 통합).

    1) 색상/폰트 자동 분석 → `vision_analysis` 필드 갱신.
       (빈 design_tokens.colors 슬롯은 추출된 컬러로 자동 채워진다 — 수동 입력값 보존.)
    2) 슬라이드 레이아웃 패턴 자동 추출 → `extracted_patterns` 필드 덮어쓰기.

    응답:
    {
      "style": {...최종 도큐먼트...},
      "vision_analysis": {...},
      "extracted_patterns": [...],
      "total_patterns_extracted": N
    }
    """
    await verify_admin(jwt_token)
    db = get_db()
    oid = _to_oid(style_id)
    existing = await db.ppt_styles.find_one({"_id": oid})
    if not existing:
        raise HTTPException(status_code=404, detail="스타일을 찾을 수 없습니다")

    if not (existing.get("sample_image_refs") or []):
        raise HTTPException(status_code=400, detail="분석할 샘플 이미지가 없습니다. 먼저 업로드하세요.")

    # 1) 색상/폰트 Vision 분석 (기존 동작)
    updated = await ppt_style_service.analyze_samples_with_vision(style_id)
    if not updated:
        raise HTTPException(status_code=500, detail="분석에 실패했습니다")

    # 2) 패턴 추출 (M8 신규)
    try:
        extract_result = await ppt_style_service.extract_patterns_from_samples(style_id)
    except Exception as e:
        # 패턴 추출이 실패해도 색상/폰트 분석은 보존
        print(f"[PPTStyle] 패턴 추출 실패 (style_id={style_id}): {e}")
        extract_result = {
            "extracted_patterns": [],
            "by_sample": [],
            "total_patterns": 0,
        }

    # 최종 도큐먼트 다시 로드 (extracted_patterns 반영)
    final_doc = await db.ppt_styles.find_one({"_id": oid}) or updated

    return {
        "style": _serialize(final_doc),
        "vision_analysis": final_doc.get("vision_analysis", {}) if isinstance(final_doc, dict) else {},
        "extracted_patterns": extract_result.get("extracted_patterns", []),
        "total_patterns_extracted": extract_result.get("total_patterns", 0),
    }


# ============ 공개(사용자) 라우트 — M4.1 ============
# admin 권한 검증 없이 JWT만 검증. 게시(is_published=True) 스타일만 노출.

@router.get("/{jwt_token}/api/ppt-styles")
async def list_published_ppt_styles(
    jwt_token: str,
    lang: Optional[str] = None,
    limit: int = 50,
):
    """게시된 PPT 스타일 목록 (사용자 공개).

    - 정렬: updated_at 내림차순
    - 필터: lang (선택), is_published=True
    - 응답 카드용 최소 필드만 포함
    """
    await verify_user(jwt_token)
    db = get_db()

    # limit 보정 (음수/0/과도값 방지)
    try:
        n = int(limit)
    except Exception:
        n = 50
    if n <= 0:
        n = 50
    if n > 200:
        n = 200

    query: dict = {"is_published": True}
    if lang:
        query["lang"] = lang

    projection = {
        "title": 1,
        "description": 1,
        "lang": 1,
        "sample_image_refs": 1,
        "pattern_library": 1,
        "design_tokens.colors.primary": 1,
        "updated_at": 1,
    }

    cursor = (
        db.ppt_styles
        .find(query, projection)
        .sort("updated_at", -1)
        .limit(n)
    )

    styles: list[dict] = []
    async for s in cursor:
        styles.append(_serialize(s))
    return {"styles": styles}


@router.get("/{jwt_token}/api/ppt-styles/{style_id}")
async def get_published_ppt_style(jwt_token: str, style_id: str):
    """게시된 PPT 스타일 단일 상세 (사용자 공개).

    is_published=False 인 경우 404 (존재 자체를 노출하지 않음).
    """
    await verify_user(jwt_token)
    db = get_db()
    doc = await db.ppt_styles.find_one({
        "_id": _to_oid(style_id),
        "is_published": True,
    })
    if not doc:
        raise HTTPException(status_code=404, detail="스타일을 찾을 수 없습니다")
    return {"style": _serialize(doc)}
