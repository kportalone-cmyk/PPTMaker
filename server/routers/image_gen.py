"""
이미지 생성 워크스페이스(image_gen project_type) 라우터.

엔드포인트:
- POST   /{jwt}/api/image-gen/{project_id}/generate
- GET    /{jwt}/api/image-gen/{project_id}
- DELETE /{jwt}/api/image-gen/{project_id}/{generation_id}
"""

from fastapi import APIRouter, HTTPException, UploadFile, File

from models.image_gen import ImageGenRequest
from services.mongo_service import get_db
from services.auth_service import decode_jwt_token, extract_user_key, get_user_flexible
from services import image_gen_service
from routers.collaboration import check_project_access


router = APIRouter(tags=["image-gen"])


async def _get_user_key(jwt_token: str) -> str:
    """JWT 에서 user_key 추출 (내부/외부 JWT 모두 지원). 다른 라우터와 동일 패턴."""
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


@router.post("/{jwt_token}/api/image-gen/{project_id}/generate")
async def generate_image(jwt_token: str, project_id: str, data: ImageGenRequest):
    """프롬프트로 이미지 N장 생성 (편집자 이상)."""
    user_key = await _get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "editor")

    prompt = (data.prompt or "").strip()
    if not prompt:
        raise HTTPException(status_code=400, detail="prompt 가 비어 있습니다")

    if data.model not in image_gen_service.ALLOWED_MODELS:
        raise HTTPException(status_code=400, detail=f"지원하지 않는 model: {data.model}")
    if data.aspect_ratio not in image_gen_service.ALLOWED_ASPECT_RATIOS:
        raise HTTPException(status_code=400, detail=f"지원하지 않는 aspect_ratio: {data.aspect_ratio}")
    if not (1 <= data.count <= 4):
        raise HTTPException(status_code=400, detail="count 는 1~4 사이여야 합니다")

    try:
        generation = await image_gen_service.generate_images(
            project_id=project_id,
            user_key=user_key,
            prompt=prompt,
            model=data.model,
            aspect_ratio=data.aspect_ratio,
            count=data.count,
            reference_file_ids=data.reference_file_ids,
        )
    except RuntimeError as e:
        raise HTTPException(status_code=500, detail=str(e))

    return {"generation": generation}


@router.post("/{jwt_token}/api/image-gen/{project_id}/upload")
async def upload_images(
    jwt_token: str,
    project_id: str,
    files: list[UploadFile] = File(...),
):
    """로컬 이미지 파일 다중 업로드 (드래그&드롭). 한 호출 = 한 image_generations 도큐먼트."""
    user_key = await _get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "editor")

    if not files:
        raise HTTPException(status_code=400, detail="업로드할 파일이 없습니다")

    try:
        generation = await image_gen_service.save_uploaded_images(
            project_id=project_id,
            user_key=user_key,
            files=files,
        )
    except RuntimeError as e:
        raise HTTPException(status_code=400, detail=str(e))

    return {"generation": generation}


@router.get("/{jwt_token}/api/image-gen/{project_id}")
async def list_image_generations(jwt_token: str, project_id: str):
    """프로젝트의 이미지 생성 이력 조회 (뷰어 이상)."""
    user_key = await _get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "viewer")

    generations = await image_gen_service.list_generations(project_id)
    return {"generations": generations}


@router.delete("/{jwt_token}/api/image-gen/{project_id}/{generation_id}")
async def delete_image_generation(jwt_token: str, project_id: str, generation_id: str):
    """이미지 생성 도큐먼트 + 디스크 파일 삭제 (편집자 이상). 묶음 전체."""
    user_key = await _get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "editor")

    ok = await image_gen_service.delete_generation(project_id, generation_id)
    if not ok:
        raise HTTPException(status_code=404, detail="이미지 생성 기록을 찾을 수 없습니다")
    return {"success": True}


@router.delete("/{jwt_token}/api/image-gen/{project_id}/{generation_id}/images/{file_id}")
async def delete_image_generation_image(
    jwt_token: str, project_id: str, generation_id: str, file_id: str
):
    """generation 안의 이미지 한 장만 삭제 (편집자 이상).

    응답:
      - { "success": true, "generation": {...갱신 도큐먼트...} } — 남은 이미지가 있음
      - { "success": true, "deleted_generation": true } — 마지막 이미지였고 도큐먼트도 삭제됨
    """
    user_key = await _get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "editor")

    result = await image_gen_service.delete_generation_image(project_id, generation_id, file_id)
    if result is None:
        raise HTTPException(status_code=404, detail="이미지를 찾을 수 없습니다")
    if result.get("deleted_generation"):
        return {"success": True, "deleted_generation": True}
    return {"success": True, "generation": result}
