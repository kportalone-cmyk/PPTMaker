"""
외부 REST API — 파일 업로드 + 지침으로 슬라이드 자동 생성 후 공유 URL 반환

사용법:
  POST /api/external/generate
  - Content-Type: multipart/form-data
  - 필드: instructions (지침), file (파일), lang (언어, 선택), slide_count (슬라이드 수, 선택)
  - 헤더: X-API-Key (인증키)

응답:
  {
    "success": true,
    "project_id": "...",
    "share_url": "https://your-domain/shared/abc123...",
    "slide_count": 12
  }
"""

import os
import uuid
import hashlib
import time
import json
from pathlib import Path
from datetime import datetime

from fastapi import APIRouter, HTTPException, UploadFile, File, Form, Header
from bson import ObjectId

from config import settings
from services.mongo_service import get_db
from services import redis_service
from services.file_service import extract_text_from_file
from services.search_service import search_web
from services.llm_service import generate_slide_content_stream, _call_claude_api
from services.infographic_service import generate_infographic_batch
from routers.prompt import get_prompt_content, get_prompt_model
import aiofiles

router = APIRouter(prefix="/api/external", tags=["external"])

# 업로드 디렉토리
UPLOAD_DIR = Path(settings.UPLOAD_DIR).resolve()
RESOURCE_DIR = UPLOAD_DIR / "resources"
RESOURCE_DIR.mkdir(parents=True, exist_ok=True)

# 허용 파일 확장자
ALLOWED_EXT = {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".pdf", ".txt", ".csv", ".md"}

# 외부 API 인증키 (.env의 EXTERNAL_API_KEY)
EXTERNAL_API_KEY = os.getenv("EXTERNAL_API_KEY", "")


def _verify_api_key(api_key: str):
    """외부 API 키 검증"""
    if not EXTERNAL_API_KEY:
        # API 키 미설정 시 모든 요청 허용 (개발 모드)
        return
    if api_key != EXTERNAL_API_KEY:
        raise HTTPException(status_code=401, detail="Invalid API key")


@router.post("/generate")
async def external_generate(
    instructions: str = Form(""),
    file: UploadFile = File(None),
    lang: str = Form("ko"),
    slide_count: str = Form("auto"),
    user_key: str = Form("external_api"),
    x_api_key: str = Header("", alias="X-API-Key"),
):
    """
    외부 시스템에서 파일 + 지침을 전달받아 AI 자동 디자인 슬라이드를 생성하고 공유 URL을 반환합니다.

    - file: 분석할 파일 (docx, xlsx, pptx, pdf, txt 등) — 선택
    - instructions: 슬라이드 생성 지침 — 필수 (파일 없을 시)
    - lang: 출력 언어 (기본: ko)
    - slide_count: 슬라이드 수 (기본: auto)
    - user_key: 사용자 식별키 (기본: external_api)
    - X-API-Key 헤더: 인증키
    """
    _verify_api_key(x_api_key)

    if not instructions and not file:
        raise HTTPException(status_code=400, detail="instructions 또는 file을 제공해야 합니다.")

    db = get_db()

    # ── 1. 프로젝트 자동 생성 ──
    project_name = instructions[:50] if instructions else (file.filename if file else "External API")
    project_doc = {
        "user_key": user_key,
        "name": project_name,
        "description": "외부 API를 통해 생성된 프로젝트",
        "project_type": "slide",
        "status": "generating",
        "infographic_mode": True,
        "auto_template": True,
        "created_at": datetime.utcnow(),
        "updated_at": datetime.utcnow(),
    }
    result = await db.projects.insert_one(project_doc)
    project_id = str(result.inserted_id)

    # ── 2. 파일 업로드 + 텍스트 추출 (SHA-256 해시 기반 캐시) ──
    combined_text = ""
    if file:
        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in ALLOWED_EXT:
            raise HTTPException(status_code=400, detail=f"허용되지 않은 파일 형식: {ext}")

        file_bytes = await file.read()

        # SHA-256 해시 계산
        file_hash = hashlib.sha256(file_bytes).hexdigest()

        # DB에서 동일 해시의 캐시된 텍스트 조회
        cached = await db.file_text_cache.find_one({"file_hash": file_hash})
        if cached:
            extracted_content = cached.get("content", "")
            print(f"[ExternalAPI] 파일 캐시 히트: {file.filename} (hash: {file_hash[:16]}...)")
        else:
            # 파일 저장 후 텍스트 추출
            filename = f"{uuid.uuid4().hex}{ext}"
            save_path = str(RESOURCE_DIR / filename)

            async with aiofiles.open(save_path, "wb") as f:
                await f.write(file_bytes)

            extracted_content = extract_text_from_file(save_path, ext) or ""

            # 텍스트 추출 완료 → 업로드 파일 삭제
            try:
                os.remove(save_path)
                print(f"[ExternalAPI] 업로드 파일 삭제: {save_path}")
            except OSError:
                pass

            # 해시 기반 캐시 DB 저장
            await db.file_text_cache.update_one(
                {"file_hash": file_hash},
                {"$set": {
                    "file_hash": file_hash,
                    "content": extracted_content,
                    "original_filename": file.filename,
                    "file_size": len(file_bytes),
                    "created_at": datetime.utcnow(),
                }},
                upsert=True,
            )
            print(f"[ExternalAPI] 파일 텍스트 추출 및 캐시 저장: {file.filename} (hash: {file_hash[:16]}...)")

        combined_text = extracted_content

        # 리소스 DB 등록
        await db.resources.insert_one({
            "project_id": project_id,
            "resource_type": "file",
            "title": file.filename,
            "content": extracted_content,
            "file_path": "",
            "original_filename": file.filename,
            "file_hash": file_hash,
            "created_at": datetime.utcnow(),
        })

    # ── 3. 리소스가 없으면 지침으로 웹 검색 ──
    if not combined_text.strip() and instructions.strip():
        search_result = await search_web(instructions)
        pages = search_result.get("pages", [])
        if pages:
            search_contents = []
            for page in pages:
                title = page.get("title", "")
                content = page.get("content", "")
                if content:
                    search_contents.append(f"[{title}]\n{content}")
            combined_text = "\n\n".join(search_contents)
            if combined_text.strip():
                await db.resources.insert_one({
                    "project_id": project_id,
                    "resource_type": "web",
                    "title": f"자동 웹 검색: {instructions[:50]}",
                    "content": combined_text,
                    "created_at": datetime.utcnow(),
                })

    if not combined_text.strip():
        # 지침만 있는 경우 지침 자체를 리소스로 사용
        if instructions.strip():
            combined_text = instructions
            await db.resources.insert_one({
                "project_id": project_id,
                "resource_type": "text",
                "title": "사용자 지침",
                "content": instructions,
                "created_at": datetime.utcnow(),
            })
        else:
            raise HTTPException(status_code=400, detail="분석할 내용이 없습니다.")

    try:
        # ── 4. 아웃라인 생성 ──
        dummy_meta = [{
            "slide_index": 0,
            "slide_meta": {"content_type": "body", "has_title": True},
            "placeholders": [
                {"placeholder": "auto_title_0", "role": "title"},
                {"placeholder": "auto_description_0", "role": "description"},
            ],
        }]

        infographic_instruction = await get_prompt_content("infographic_outline_instruction")
        full_instructions = (instructions or "") + infographic_instruction

        llm_result = None
        async for event_type, event_data in generate_slide_content_stream(
            combined_text, full_instructions, dummy_meta, lang,
            slide_count=slide_count,
        ):
            if event_type == "result":
                llm_result = event_data

        if not llm_result:
            await _update_project_error(db, project_id)
            raise HTTPException(status_code=500, detail="아웃라인 생성에 실패했습니다.")

        raw_slides = llm_result.get("raw_slides", [])
        if not raw_slides:
            raw_slides = llm_result.get("slides", [])

        # ── 5. AI 자동 디자인 스타일 생성 ──
        style_hint = ""
        try:
            auto_system = await get_prompt_content("auto_template_system")
            auto_user_tpl = await get_prompt_content("auto_template_user")
            resource_excerpt = combined_text[:3000]
            auto_user = auto_user_tpl.format(
                resources_text=resource_excerpt,
                instructions=instructions or "없음",
            )
            auto_model = await get_prompt_model("auto_template_system")
            effective_model = auto_model or "claude-sonnet-4-6"
            auto_result = await _call_claude_api(auto_system, auto_user, model=effective_model)

            import re
            json_match = re.search(r'```json\s*(.*?)\s*```', auto_result, re.DOTALL)
            if json_match:
                design_style = json.loads(json_match.group(1))
            else:
                design_style = json.loads(auto_result)

            style_hint = design_style.get("style_prompt", "")
        except Exception as e:
            print(f"[ExternalAPI] 디자인 스타일 생성 실패: {e}")

        # ── 6. 인포그래픽 이미지 생성 ──
        generated_count = 0
        async for img_result in generate_infographic_batch(
            raw_slides,
            style_hint=style_hint,
            aspect_ratio="16:9",
            infographic_ratio=40,
        ):
            idx = img_result["index"]
            image_url = img_result["image_url"]
            title = img_result["title"]
            subtitle = img_result.get("subtitle", "")

            gen_objects = []
            if image_url:
                gen_objects.append({
                    "obj_type": "image",
                    "x": 0, "y": 0,
                    "width": 960, "height": 540,
                    "image_url": image_url,
                    "image_fit": "cover",
                    "z_index": 1,
                })

            # 커버 슬라이드: 제목 텍스트 오버레이
            if idx == 0:
                gen_objects.append({
                    "obj_type": "text",
                    "x": 80, "y": 220,
                    "width": 800, "height": 80,
                    "text_content": title,
                    "generated_text": title,
                    "role": "title",
                    "z_index": 10,
                    "text_style": {
                        "font_size": 40,
                        "bold": True,
                        "color": "#FFFFFF",
                        "align": "center",
                        "font_family": "Pretendard",
                    }
                })

            gen_slide = {
                "project_id": project_id,
                "template_slide_id": "infographic",
                "order": idx + 1,
                "objects": gen_objects,
                "items": [],
                "background_image": None,
                "infographic": True,
                "infographic_title": title,
                "created_at": datetime.utcnow(),
            }
            if idx == 0 and subtitle:
                gen_slide["infographic_subtitle"] = subtitle

            await db.generated_slides.insert_one(gen_slide)
            generated_count += 1

        # ── 7. 공유 링크 생성 ──
        share_token = hashlib.md5(f"{project_id}{time.time()}".encode()).hexdigest()
        await db.projects.update_one(
            {"_id": ObjectId(project_id)},
            {"$set": {
                "status": "generated",
                "share_token": share_token,
                "updated_at": datetime.utcnow(),
            }}
        )

        # 공유 URL 생성
        base_url = settings.SERVER_BASE_URL.rstrip("/") if settings.SERVER_BASE_URL else ""
        share_url = f"{base_url}/shared/{share_token}"

        return {
            "success": True,
            "project_id": project_id,
            "share_url": share_url,
            "share_token": share_token,
            "slide_count": generated_count,
        }

    except HTTPException:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        await _update_project_error(db, project_id)
        raise HTTPException(status_code=500, detail=f"슬라이드 생성 중 오류: {str(e)}")


async def _update_project_error(db, project_id: str):
    """프로젝트 상태를 에러로 업데이트"""
    await db.projects.update_one(
        {"_id": ObjectId(project_id)},
        {"$set": {"status": "error", "updated_at": datetime.utcnow()}}
    )
