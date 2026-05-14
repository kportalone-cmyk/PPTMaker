"""
이미지 생성 워크스페이스 서비스 (Google Flow 스타일).

프론트 model 키 → Google AI 모델 ID 매핑 + 이미지 호출/저장/DB 영속화.
- nano-banana-pro, nano-banana-2: Gemini image-preview 계열, 1장씩 N번 호출 (asyncio.gather)
- imagen-4: Imagen 4 단일 호출로 N장 배치 생성
"""

import uuid
import asyncio
import mimetypes
from pathlib import Path
from datetime import datetime
from typing import Optional

from bson import ObjectId
from google import genai
from google.genai import types

from config import settings, google_key_rotator
from services.mongo_service import get_db


# ── 상수 / 매핑 ──────────────────────────────────────────
ALLOWED_MODELS = ("nano-banana-pro", "nano-banana-2", "imagen-4")
ALLOWED_ASPECT_RATIOS = ("16:9", "4:3", "1:1", "3:4", "9:16")
GEMINI_MODELS = ("nano-banana-pro", "nano-banana-2")
IMAGEN_MODELS = ("imagen-4",)

IMAGE_GEN_DIR = Path(settings.UPLOAD_DIR).resolve() / "image_gen"

# Gemini 클라이언트 (키별 캐시 — infographic_service 와 동일 패턴)
_clients: dict[str, genai.Client] = {}


def _resolve_model_id(model_key: str) -> str:
    """프론트 model 키를 실제 Google AI 모델 ID 로 변환"""
    if model_key == "nano-banana-pro":
        return settings.NANO_BANANA_PRO_MODEL
    if model_key == "nano-banana-2":
        return settings.NANO_BANANA_2_MODEL
    if model_key == "imagen-4":
        return settings.IMAGEN_4_MODEL
    raise ValueError(f"Unknown model key: {model_key}")


def _get_client() -> genai.Client:
    """라운드 로빈으로 Google API 키 선택 후 클라이언트 반환"""
    api_key = google_key_rotator.next()
    if not api_key:
        raise RuntimeError("GOOGLE_API_KEY 가 설정되지 않았습니다.")
    if api_key not in _clients:
        _clients[api_key] = genai.Client(api_key=api_key)
    return _clients[api_key]


def _project_dir(project_id: str) -> Path:
    """프로젝트별 image_gen 디렉토리 (없으면 mkdir)"""
    d = IMAGE_GEN_DIR / project_id
    d.mkdir(parents=True, exist_ok=True)
    return d


def _save_image_bytes(project_id: str, image_bytes: bytes, mime_type: Optional[str]) -> dict:
    """이미지 바이트를 디스크에 저장하고 {file_id, url} 반환.
    저장 실패 시 {"error": ...} 반환.
    """
    try:
        ext = mimetypes.guess_extension(mime_type) if mime_type else None
        if not ext or ext == ".jpe":
            ext = ".png"
        file_id = uuid.uuid4().hex
        filename = f"{file_id}{ext}"
        path = _project_dir(project_id) / filename
        path.write_bytes(image_bytes)
        return {
            "file_id": file_id,
            "url": f"/uploads/image_gen/{project_id}/{filename}",
        }
    except Exception as e:
        print(f"[ImageGen] 파일 저장 실패: {e}")
        return {"error": f"파일 저장 실패: {e}"}


def _load_reference_image_bytes(project_id: str, file_id: str) -> Optional[tuple[bytes, str]]:
    """동일 프로젝트의 image_generations 도큐먼트 안에서 file_id 와 일치하는 파일을 디스크에서 읽어 반환.

    Returns (bytes, mime_type) or None.
    """
    pdir = _project_dir(project_id)
    if not pdir.exists():
        return None
    # 파일 시스템상 file_id 가 파일명 prefix
    for path in pdir.iterdir():
        try:
            if path.is_file() and path.stem == file_id:
                data = path.read_bytes()
                mime = mimetypes.guess_type(path.name)[0] or "image/png"
                return data, mime
        except Exception:
            continue
    return None


async def _resolve_reference_parts(project_id: str, file_ids: list[str]) -> list:
    """참조 file_id 들을 Gemini `types.Part` 리스트로 변환.

    스레드풀에서 디스크 I/O 수행. 못 찾은 file_id 는 조용히 스킵 (호출 측에서 경고 처리).
    """
    if not file_ids:
        return []

    def _sync_load() -> list:
        parts: list = []
        for fid in file_ids:
            loaded = _load_reference_image_bytes(project_id, fid)
            if not loaded:
                print(f"[ImageGen] 참조 이미지 누락 (file_id={fid})")
                continue
            data, mime = loaded
            parts.append(types.Part.from_bytes(data=data, mime_type=mime))
        return parts

    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _sync_load)


# ── Gemini (nano-banana) 호출 ────────────────────────────
def _gen_gemini_sync(
    model_id: str,
    prompt: str,
    aspect_ratio: str,
    reference_parts: Optional[list] = None,
) -> tuple[Optional[bytes], Optional[str], Optional[str]]:
    """Gemini image-preview 호출 (동기, 1장). (bytes, mime_type, error_msg) 반환.

    reference_parts: types.Part.from_bytes(...) 리스트. 있으면 텍스트 part 앞에 prepend.
    """
    try:
        client = _get_client()
        parts: list = []
        if reference_parts:
            parts.extend(reference_parts)
        parts.append(types.Part.from_text(text=prompt))
        contents = [types.Content(role="user", parts=parts)]
        config = types.GenerateContentConfig(
            image_config=types.ImageConfig(
                aspect_ratio=aspect_ratio,
                image_size="1K",
            ),
            response_modalities=["IMAGE", "TEXT"],
        )
        data_buffer: Optional[bytes] = None
        mime_type: Optional[str] = None
        for chunk in client.models.generate_content_stream(
            model=model_id,
            contents=contents,
            config=config,
        ):
            if chunk.parts is None:
                continue
            for part in chunk.parts:
                inline = getattr(part, "inline_data", None)
                if inline and inline.data:
                    data_buffer = inline.data
                    mime_type = inline.mime_type
                elif getattr(part, "text", None):
                    # 텍스트 응답은 무시 (디버그용 로그만)
                    print(f"[ImageGen] Gemini 텍스트 응답: {part.text[:200]}")
        if not data_buffer:
            return None, None, "Gemini 응답에 이미지 데이터가 없습니다."
        return data_buffer, mime_type, None
    except Exception as e:
        return None, None, f"Gemini API 호출 실패: {e}"


async def _gen_gemini_one(
    model_id: str,
    prompt: str,
    aspect_ratio: str,
    project_id: str,
    reference_parts: Optional[list] = None,
) -> dict:
    """Gemini 1장 생성 + 디스크 저장. 성공 {file_id,url} / 실패 {error}.

    reference_parts: 참조 이미지 Part 리스트 (모든 병렬 호출이 동일한 리스트 공유 가능).
    """
    loop = asyncio.get_event_loop()
    image_bytes, mime_type, err = await loop.run_in_executor(
        None, _gen_gemini_sync, model_id, prompt, aspect_ratio, reference_parts
    )
    if err or not image_bytes:
        return {"error": err or "이미지 데이터 없음"}
    return _save_image_bytes(project_id, image_bytes, mime_type)


# ── Imagen 4 호출 ─────────────────────────────────────────
def _gen_imagen_sync(model_id: str, prompt: str, aspect_ratio: str, count: int):
    """Imagen 4 호출 (동기, N장 배치). GeneratedImage 리스트 또는 ("error", msg)."""
    try:
        client = _get_client()
        response = client.models.generate_images(
            model=model_id,
            prompt=prompt,
            config=types.GenerateImagesConfig(
                number_of_images=count,
                aspect_ratio=aspect_ratio,
            ),
        )
        generated = getattr(response, "generated_images", None) or []
        return list(generated)
    except Exception as e:
        return ("error", f"Imagen API 호출 실패: {e}")


async def _gen_imagen_batch(model_id: str, prompt: str, aspect_ratio: str, count: int, project_id: str) -> list[dict]:
    """Imagen 4 배치 생성 + 디스크 저장. 결과 리스트 (각 항목은 {file_id,url} 또는 {error})."""
    loop = asyncio.get_event_loop()
    result = await loop.run_in_executor(
        None, _gen_imagen_sync, model_id, prompt, aspect_ratio, count
    )
    if isinstance(result, tuple) and result and result[0] == "error":
        return [{"error": result[1]} for _ in range(count)]

    out: list[dict] = []
    for gi in result:
        img = getattr(gi, "image", None)
        if img is None or not getattr(img, "image_bytes", None):
            rai = getattr(gi, "rai_filtered_reason", None)
            out.append({"error": f"이미지 누락{(': ' + rai) if rai else ''}"})
            continue
        out.append(_save_image_bytes(project_id, img.image_bytes, getattr(img, "mime_type", None)))

    # SDK 가 count 보다 적게 돌려준 경우 보정
    while len(out) < count:
        out.append({"error": "응답 이미지 개수 부족"})
    return out[:count]


# ── 메인 진입점 ───────────────────────────────────────────
async def generate_images(
    project_id: str,
    user_key: str,
    prompt: str,
    model: str,
    aspect_ratio: str,
    count: int,
    reference_file_ids: Optional[list[str]] = None,
) -> dict:
    """
    이미지 생성 → 디스크 저장 → image_generations 컬렉션 insert.
    반환: 저장된 도큐먼트 (_id 는 str).
    완전 실패(이미지 0개) 시 RuntimeError.

    reference_file_ids: 동일 프로젝트 image_generations 안의 file_id 들. Gemini 계열만 지원.
    """
    ref_ids = [fid for fid in (reference_file_ids or []) if fid]
    if ref_ids and model in IMAGEN_MODELS:
        raise RuntimeError("Imagen 4 는 참조 이미지를 지원하지 않습니다. 나노 바나나 모델을 선택해주세요.")

    if model in GEMINI_MODELS:
        model_id = _resolve_model_id(model)
        # 참조 이미지 Part 들은 한 번만 만들어 모든 병렬 호출이 재사용
        reference_parts = await _resolve_reference_parts(project_id, ref_ids) if ref_ids else None
        tasks = [
            _gen_gemini_one(model_id, prompt, aspect_ratio, project_id, reference_parts)
            for _ in range(count)
        ]
        images = await asyncio.gather(*tasks, return_exceptions=False)
    elif model in IMAGEN_MODELS:
        model_id = _resolve_model_id(model)
        images = await _gen_imagen_batch(model_id, prompt, aspect_ratio, count, project_id)
    else:
        raise ValueError(f"Unknown model: {model}")

    # 전부 실패면 예외
    if not any("url" in img for img in images):
        # 첫 에러 메시지를 대표로 사용
        first_err = next((img.get("error") for img in images if img.get("error")), "이미지 생성 실패")
        raise RuntimeError(first_err)

    db = get_db()
    doc = {
        "project_id": project_id,
        "user_key": user_key,
        "prompt": prompt,
        "model": model,
        "aspect_ratio": aspect_ratio,
        "count": count,
        "images": images,
        "created_at": datetime.utcnow(),
    }
    if ref_ids:
        doc["reference_file_ids"] = ref_ids
    result = await db.image_generations.insert_one(doc)
    doc["_id"] = str(result.inserted_id)
    return doc


# ── 업로드 (로컬 파일 → image_gen 도큐먼트) ──────────────
ALLOWED_UPLOAD_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"}
_MAX_UPLOAD_SIZE = 20 * 1024 * 1024  # 20 MB / 파일


def _read_image_size(path: Path) -> Optional[tuple[int, int]]:
    """저장된 이미지 파일에서 (width, height) 추출. 실패 시 None.
    의존성을 최소화하기 위해 PIL 이 있으면 사용, 없으면 None (프론트가 자체 자연 크기로 폴백).
    """
    try:
        from PIL import Image  # type: ignore
        with Image.open(path) as im:
            return im.size  # (w, h)
    except Exception:
        return None


async def save_uploaded_images(
    project_id: str,
    user_key: str,
    files: list,
) -> dict:
    """드래그&드롭으로 업로드된 이미지 파일들을 image_gen 도큐먼트로 영속화.

    files: FastAPI UploadFile 리스트. 각 파일은 ALLOWED_UPLOAD_EXTS 확장자만 허용.
    한 번의 업로드 액션 = 한 개의 image_generations 도큐먼트 (model="upload").
    각 image 에 width/height 가 들어가 프론트가 비율을 그대로 표시할 수 있다.
    """
    if not files:
        raise RuntimeError("업로드할 파일이 없습니다")

    saved_images: list[dict] = []
    original_names: list[str] = []
    for f in files:
        original_name = getattr(f, "filename", "") or ""
        original_names.append(original_name)
        ext = ""
        if "." in original_name:
            ext = "." + original_name.rsplit(".", 1)[-1].lower()
        if ext not in ALLOWED_UPLOAD_EXTS:
            saved_images.append({"error": f"지원하지 않는 확장자: {ext or '(unknown)'}"})
            continue

        try:
            data = await f.read()
        except Exception as e:
            saved_images.append({"error": f"파일 읽기 실패: {e}"})
            continue
        if not data:
            saved_images.append({"error": "빈 파일"})
            continue
        if len(data) > _MAX_UPLOAD_SIZE:
            saved_images.append({"error": f"파일 크기 초과 (최대 {_MAX_UPLOAD_SIZE // (1024 * 1024)} MB)"})
            continue

        try:
            file_id = uuid.uuid4().hex
            filename = f"{file_id}{ext}"
            path = _project_dir(project_id) / filename
            path.write_bytes(data)
        except Exception as e:
            saved_images.append({"error": f"저장 실패: {e}"})
            continue

        img_record: dict = {
            "file_id": file_id,
            "url": f"/uploads/image_gen/{project_id}/{filename}",
            "original_filename": original_name,
        }
        size = _read_image_size(path)
        if size:
            img_record["width"], img_record["height"] = size
        saved_images.append(img_record)

    if not any("url" in img for img in saved_images):
        # 전부 실패
        first_err = next((img.get("error") for img in saved_images if img.get("error")), "업로드 실패")
        raise RuntimeError(first_err)

    n_ok = sum(1 for img in saved_images if "url" in img)
    label = original_names[0] if len(original_names) == 1 else f"{n_ok}장 업로드"

    db = get_db()
    doc = {
        "project_id": project_id,
        "user_key": user_key,
        "prompt": label,
        "model": "upload",
        "aspect_ratio": "auto",
        "count": n_ok,
        "images": saved_images,
        "source": "upload",
        "created_at": datetime.utcnow(),
    }
    result = await db.image_generations.insert_one(doc)
    doc["_id"] = str(result.inserted_id)
    return doc


async def list_generations(project_id: str) -> list[dict]:
    """프로젝트의 이미지 생성 이력 (created_at desc)"""
    db = get_db()
    cursor = db.image_generations.find({"project_id": project_id}).sort("created_at", -1)
    out: list[dict] = []
    async for d in cursor:
        d["_id"] = str(d["_id"])
        out.append(d)
    return out


async def delete_generation(project_id: str, generation_id: str) -> bool:
    """도큐먼트 + 디스크 파일 삭제. 도큐먼트 없으면 False."""
    db = get_db()
    try:
        oid = ObjectId(generation_id)
    except Exception:
        return False
    doc = await db.image_generations.find_one({"_id": oid, "project_id": project_id})
    if not doc:
        return False

    # 디스크 파일 정리 (실패해도 로그만)
    for img in doc.get("images", []):
        url = img.get("url") or ""
        if not url.startswith(f"/uploads/image_gen/{project_id}/"):
            continue
        filename = url.rsplit("/", 1)[-1]
        path = _project_dir(project_id) / filename
        try:
            if path.exists():
                path.unlink()
        except Exception as e:
            print(f"[ImageGen] 파일 삭제 실패 ({path}): {e}")

    await db.image_generations.delete_one({"_id": oid})
    return True


async def delete_generation_image(project_id: str, generation_id: str, file_id: str) -> Optional[dict]:
    """generation 안의 이미지 한 장만 삭제.

    반환:
      - 갱신된 generation 도큐먼트 (남은 이미지 ≥ 1)
      - 마지막 이미지였으면 빈 dict + key "deleted_generation": True
      - 도큐먼트/이미지를 찾지 못하면 None
    """
    db = get_db()
    try:
        oid = ObjectId(generation_id)
    except Exception:
        return None
    doc = await db.image_generations.find_one({"_id": oid, "project_id": project_id})
    if not doc:
        return None

    images = doc.get("images", []) or []
    target = None
    remaining: list[dict] = []
    for img in images:
        if (img.get("file_id") == file_id) and target is None:
            target = img
        else:
            remaining.append(img)
    if target is None:
        return None

    # 디스크 파일 정리 (실패해도 로그만)
    url = (target.get("url") or "")
    if url.startswith(f"/uploads/image_gen/{project_id}/"):
        filename = url.rsplit("/", 1)[-1]
        path = _project_dir(project_id) / filename
        try:
            if path.exists():
                path.unlink()
        except Exception as e:
            print(f"[ImageGen] 파일 삭제 실패 ({path}): {e}")

    # 남은 이미지가 0 이면 도큐먼트 자체 삭제
    successful_remaining = [img for img in remaining if img.get("url")]
    if not successful_remaining:
        await db.image_generations.delete_one({"_id": oid})
        return {"deleted_generation": True}

    # 도큐먼트 업데이트 (images / count)
    new_count = len(successful_remaining)
    await db.image_generations.update_one(
        {"_id": oid},
        {"$set": {"images": remaining, "count": new_count}},
    )
    updated = await db.image_generations.find_one({"_id": oid})
    if updated:
        updated["_id"] = str(updated["_id"])
    return updated
