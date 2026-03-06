from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import FileResponse, StreamingResponse
from bson import ObjectId
from datetime import datetime
from models.project import GenerateRequest, SlideReorderRequest, SlideUpdateRequest
from services.mongo_service import get_db
from services.auth_service import decode_jwt_token, extract_user_key, get_user_flexible
from services.template_service import recommend_slide, get_template_slides
from services.ppt_service import generate_pptx
from services.llm_service import generate_slide_content, generate_slide_content_stream
from config import settings
import os
import uuid
import json
import aiofiles

router = APIRouter(tags=["generate"])


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


@router.post("/{jwt_token}/api/generate/prepare")
async def prepare_generation(jwt_token: str, data: GenerateRequest):
    """PPT 생성 준비 - 리소스 수집 및 슬라이드 매칭"""
    await get_user_key(jwt_token)
    db = get_db()

    # 프로젝트 리소스 조회
    cursor = db.resources.find({"project_id": data.project_id})
    resources = []
    async for r in cursor:
        r["_id"] = str(r["_id"])
        resources.append(r)

    if not resources:
        raise HTTPException(status_code=400, detail="리소스가 없습니다. 먼저 자료를 등록하세요.")

    # 템플릿 슬라이드 조회
    slides = await get_template_slides(data.template_id)
    if not slides:
        raise HTTPException(status_code=400, detail="템플릿에 슬라이드가 없습니다.")

    # 프로젝트에 템플릿/지침 저장
    await db.projects.update_one(
        {"_id": ObjectId(data.project_id)},
        {"$set": {
            "template_id": data.template_id,
            "instructions": data.instructions,
            "status": "preparing",
            "updated_at": datetime.utcnow(),
        }}
    )

    return {
        "project_id": data.project_id,
        "resource_count": len(resources),
        "template_slide_count": len(slides),
        "status": "ready"
    }


@router.post("/{jwt_token}/api/generate/slides")
async def generate_slides(jwt_token: str, data: GenerateRequest):
    """Claude Opus를 활용하여 리소스 내용을 리치 스키마 기반 슬라이드 콘텐츠로 생성

    핵심 흐름:
    1. 템플릿 슬라이드 분석 (오브젝트 역할/위치 자동 추론)
    2. LLM에 리소스 + 템플릿 메타 전달
    3. LLM이 리치 스키마(타입/레이아웃/구조화된 콘텐츠) 생성
    4. 스키마를 템플릿 오브젝트에 역할 기반으로 매핑
    """
    await get_user_key(jwt_token)
    db = get_db()

    # 기존 생성 슬라이드 삭제
    await db.generated_slides.delete_many({"project_id": data.project_id})

    # 리소스 텍스트 수집
    cursor = db.resources.find({"project_id": data.project_id})
    all_content = []
    async for r in cursor:
        if r.get("content"):
            all_content.append(r["content"])
        elif r.get("title"):
            all_content.append(f"[{r['title']}]")

    combined_text = "\n\n".join(all_content)

    if not combined_text.strip():
        raise HTTPException(status_code=400, detail="리소스에 텍스트 내용이 없습니다.")

    # 템플릿 슬라이드 조회
    slides = await get_template_slides(data.template_id)
    if not slides:
        raise HTTPException(status_code=400, detail="템플릿에 슬라이드가 없습니다.")

    # 배경 이미지
    template = await db.templates.find_one({"_id": ObjectId(data.template_id)})
    bg_image = template.get("background_image") if template else None

    # ── 핵심: 템플릿 슬라이드 분석 및 자동 placeholder 할당 ──
    slides_meta = _analyze_template_slides(slides)

    # Claude Opus API 호출하여 리치 스키마 콘텐츠 생성
    llm_result = await generate_slide_content(
        resources_text=combined_text,
        instructions=data.instructions,
        slides_meta=slides_meta,
        lang=data.lang,
        slide_count=data.slide_count,
    )

    # 리치 스키마 결과에서 슬라이드/메타/출처 추출
    if isinstance(llm_result, dict):
        llm_slides = llm_result.get("slides", [])
        presentation_meta = llm_result.get("meta", {})
        presentation_sources = llm_result.get("sources", [])
    else:
        llm_slides = llm_result
        presentation_meta = {}
        presentation_sources = []

    # 템플릿 슬라이드를 인덱스로 조회할 수 있는 lookup 생성
    template_lookup = {idx: slide for idx, slide in enumerate(slides)}

    # AI가 설계한 순서대로 슬라이드 생성 (template_index로 템플릿 선택)
    generated = []
    for output_order, item in enumerate(llm_slides):
        template_idx = item.get("template_index", 0)

        # 유효하지 않은 template_index 처리
        if template_idx not in template_lookup:
            template_idx = _find_fallback_template(slides)

        template_slide = template_lookup[template_idx]
        contents = item.get("contents", {})
        gen_objects = _build_gen_objects(template_slide, contents)

        # 배경이미지: 슬라이드별 배경 > 템플릿 전체 배경
        slide_bg = template_slide.get("background_image") or bg_image

        gen_slide = {
            "project_id": data.project_id,
            "template_slide_id": str(template_slide["_id"]),
            "order": output_order + 1,
            "objects": gen_objects,
            "items": item.get("items", []),
            "background_image": slide_bg,
            "created_at": datetime.utcnow(),
        }
        result = await db.generated_slides.insert_one(gen_slide)
        gen_slide["_id"] = str(result.inserted_id)
        generated.append(gen_slide)

    # 프로젝트에 프레젠테이션 메타/출처/상태 업데이트
    update_fields = {
        "status": "generated",
        "updated_at": datetime.utcnow(),
    }
    if presentation_meta:
        update_fields["presentation_meta"] = presentation_meta
    if presentation_sources:
        update_fields["presentation_sources"] = presentation_sources

    await db.projects.update_one(
        {"_id": ObjectId(data.project_id)},
        {"$set": update_fields}
    )

    return {"slides": generated, "total": len(generated)}


@router.post("/{jwt_token}/api/generate/stream")
async def generate_slides_stream(jwt_token: str, data: GenerateRequest):
    """SSE 스트리밍으로 슬라이드 생성 - 실시간 진행 표시"""
    await get_user_key(jwt_token)
    db = get_db()

    # 기존 생성 슬라이드 삭제
    await db.generated_slides.delete_many({"project_id": data.project_id})

    # 리소스 텍스트 수집
    cursor = db.resources.find({"project_id": data.project_id})
    all_content = []
    async for r in cursor:
        if r.get("content"):
            all_content.append(r["content"])
        elif r.get("title"):
            all_content.append(f"[{r['title']}]")
    combined_text = "\n\n".join(all_content)

    if not combined_text.strip():
        raise HTTPException(status_code=400, detail="리소스에 텍스트 내용이 없습니다.")

    # 템플릿 슬라이드 조회
    slides = await get_template_slides(data.template_id)
    if not slides:
        raise HTTPException(status_code=400, detail="템플릿에 슬라이드가 없습니다.")

    template = await db.templates.find_one({"_id": ObjectId(data.template_id)})
    bg_image = template.get("background_image") if template else None

    # 템플릿 슬라이드 분석
    slides_meta = _analyze_template_slides(slides)

    # 생성 ID (중복 생성 방지 및 취소 체크용)
    generation_id = str(uuid.uuid4())

    # 프로젝트 상태 업데이트
    await db.projects.update_one(
        {"_id": ObjectId(data.project_id)},
        {"$set": {
            "template_id": data.template_id,
            "instructions": data.instructions,
            "status": "generating",
            "generation_id": generation_id,
            "updated_at": datetime.utcnow(),
        }}
    )

    def _sse(event: str, payload: dict) -> str:
        return f"data: {json.dumps({'event': event, **payload}, ensure_ascii=False, default=str)}\n\n"

    async def event_stream():
        try:
            # 템플릿 타입 분석 결과 전송
            type_availability = _get_type_availability(slides_meta)
            yield _sse("template_analysis", {"types": type_availability})

            yield _sse("start", {"message": "AI가 슬라이드를 설계하고 있습니다..."})

            # 취소 체크
            if await _check_cancelled(db, data.project_id, generation_id):
                yield _sse("stopped", {"message": "생성이 중단되었습니다."})
                await _set_stopped(db, data.project_id)
                return

            llm_result = None
            chunk_count = 0
            cancelled = False
            async for event_type, event_data in generate_slide_content_stream(
                combined_text, data.instructions, slides_meta, data.lang,
                slide_count=data.slide_count,
            ):
                if event_type == "delta":
                    yield _sse("delta", {"text": event_data})
                    chunk_count += 1
                    # 50 chunk마다 취소 체크 (~2-3초 간격)
                    if chunk_count % 50 == 0:
                        if await _check_cancelled(db, data.project_id, generation_id):
                            cancelled = True
                            break
                elif event_type == "result":
                    llm_result = event_data

            if cancelled:
                yield _sse("stopped", {"message": "생성이 중단되었습니다."})
                await _set_stopped(db, data.project_id)
                return

            if not llm_result:
                yield _sse("error", {"message": "콘텐츠 생성 실패"})
                return

            yield _sse("parsing", {"message": "슬라이드를 구성하고 있습니다..."})

            # 아웃라인 데이터 전송 (raw_slides에서 추출)
            raw_slides = llm_result.get("raw_slides", [])
            if raw_slides:
                outline_items = []
                for rs in raw_slides:
                    slide_type = rs.get("type", "content")
                    title = rs.get("title", "") or rs.get("section_title", "")
                    items_count = len(rs.get("items", []))
                    outline_items.append({
                        "type": slide_type,
                        "title": title,
                        "items_count": items_count,
                    })
                yield _sse("outline", {"slides": outline_items})

            # 리치 스키마 결과 처리
            llm_slides = llm_result.get("slides", [])
            presentation_meta = llm_result.get("meta", {})
            presentation_sources = llm_result.get("sources", [])

            template_lookup = {idx: slide for idx, slide in enumerate(slides)}
            generated = []

            for output_order, item in enumerate(llm_slides):
                # 각 슬라이드 저장 전 취소 체크
                if await _check_cancelled(db, data.project_id, generation_id):
                    yield _sse("stopped", {"message": "생성이 중단되었습니다."})
                    await _set_stopped(db, data.project_id)
                    return

                template_idx = item.get("template_index", 0)
                if template_idx not in template_lookup:
                    template_idx = _find_fallback_template(slides)

                template_slide = template_lookup[template_idx]
                contents = item.get("contents", {})
                gen_objects = _build_gen_objects(template_slide, contents)

                # 배경이미지: 슬라이드별 배경 > 템플릿 전체 배경
                slide_bg = template_slide.get("background_image") or bg_image

                gen_slide = {
                    "project_id": data.project_id,
                    "template_slide_id": str(template_slide["_id"]),
                    "order": output_order + 1,
                    "objects": gen_objects,
                    "items": item.get("items", []),
                    "background_image": slide_bg,
                    "created_at": datetime.utcnow(),
                }
                result = await db.generated_slides.insert_one(gen_slide)
                gen_slide["_id"] = str(result.inserted_id)
                generated.append(gen_slide)

                yield _sse("slide", {"slide": gen_slide})

            # 프로젝트 업데이트
            update_fields = {
                "status": "generated",
                "updated_at": datetime.utcnow(),
            }
            if presentation_meta:
                update_fields["presentation_meta"] = presentation_meta
            if presentation_sources:
                update_fields["presentation_sources"] = presentation_sources

            await db.projects.update_one(
                {"_id": ObjectId(data.project_id)},
                {"$set": update_fields}
            )

            yield _sse("complete", {"total": len(generated)})

        except Exception as e:
            print(f"[SSE] Stream error: {e}")
            # 중단 요청으로 인한 에러인지 확인
            project = await db.projects.find_one(
                {"_id": ObjectId(data.project_id)}, {"status": 1}
            )
            if project and project.get("status") in ("stop_requested", "stopped"):
                yield _sse("stopped", {"message": "생성이 중단되었습니다."})
            else:
                yield _sse("error", {"message": str(e)})

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        }
    )


@router.post("/{jwt_token}/api/generate/stop/{project_id}")
async def stop_generation(jwt_token: str, project_id: str):
    """생성 중단 요청"""
    await get_user_key(jwt_token)
    db = get_db()

    project = await db.projects.find_one({"_id": ObjectId(project_id)})
    if not project:
        raise HTTPException(status_code=404, detail="프로젝트를 찾을 수 없습니다")

    if project.get("status") != "generating":
        return {"success": True, "message": "생성 중이 아닙니다"}

    await db.projects.update_one(
        {"_id": ObjectId(project_id)},
        {"$set": {
            "status": "stop_requested",
            "updated_at": datetime.utcnow(),
        }}
    )

    return {"success": True}


@router.get("/{jwt_token}/api/generate/{project_id}/slides")
async def get_generated_slides(jwt_token: str, project_id: str):
    """생성된 슬라이드 조회"""
    await get_user_key(jwt_token)
    db = get_db()
    cursor = db.generated_slides.find({"project_id": project_id}).sort("order", 1)
    slides = []
    async for s in cursor:
        s["_id"] = str(s["_id"])
        slides.append(s)
    return {"slides": slides}


@router.put("/{jwt_token}/api/generate/slides/{slide_id}")
async def update_generated_slide(jwt_token: str, slide_id: str, data: SlideUpdateRequest):
    """생성된 슬라이드 수정"""
    await get_user_key(jwt_token)
    db = get_db()
    update_fields = {
        "objects": data.objects,
        "updated_at": datetime.utcnow(),
    }
    if data.items is not None:
        update_fields["items"] = data.items
    await db.generated_slides.update_one(
        {"_id": ObjectId(slide_id)},
        {"$set": update_fields}
    )
    return {"success": True}


@router.post("/{jwt_token}/api/generate/upload-image")
async def upload_edit_image(jwt_token: str, file: UploadFile = File(...)):
    """사용자 슬라이드 편집용 이미지 업로드"""
    await get_user_key(jwt_token)
    ext = os.path.splitext(file.filename)[1]
    filename = f"{uuid.uuid4().hex}{ext}"
    save_path = os.path.join(settings.UPLOAD_DIR, "images", filename)
    os.makedirs(os.path.dirname(save_path), exist_ok=True)
    async with aiofiles.open(save_path, "wb") as f:
        content = await file.read()
        await f.write(content)
    return {"image_url": f"/uploads/images/{filename}"}


@router.put("/{jwt_token}/api/generate/{project_id}/reorder")
async def reorder_slides(jwt_token: str, project_id: str, data: SlideReorderRequest):
    """슬라이드 순서 변경"""
    await get_user_key(jwt_token)
    db = get_db()
    now = datetime.utcnow()

    for idx, slide_id in enumerate(data.slide_ids):
        await db.generated_slides.update_one(
            {"_id": ObjectId(slide_id), "project_id": project_id},
            {"$set": {"order": idx + 1, "updated_at": now}}
        )

    return {"success": True}


@router.get("/{jwt_token}/api/generate/{project_id}/download/pptx")
async def download_pptx(jwt_token: str, project_id: str):
    """PPTX 파일 다운로드"""
    await get_user_key(jwt_token)
    try:
        file_url = await generate_pptx(project_id)
        file_path = os.path.join(".", file_url.lstrip("/"))
        if not os.path.exists(file_path):
            raise HTTPException(status_code=500, detail="파일 생성 실패")

        db = get_db()
        project = await db.projects.find_one({"_id": ObjectId(project_id)})
        filename = f"{project.get('name', 'presentation')}.pptx"

        return FileResponse(
            path=file_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.get("/{jwt_token}/api/generate/{project_id}/download/pdf")
async def download_pdf(jwt_token: str, project_id: str):
    """PDF 다운로드 (미구현)"""
    await get_user_key(jwt_token)
    raise HTTPException(status_code=501, detail="PDF 변환 기능은 준비 중입니다")


# ============ 프레젠테이션 공유 ============

@router.get("/{jwt_token}/api/generate/{project_id}/share-link")
async def get_share_link(jwt_token: str, project_id: str):
    """미리보기 공유 링크 생성"""
    await get_user_key(jwt_token)
    db = get_db()

    import hashlib
    import time
    share_token = hashlib.md5(f"{project_id}{time.time()}".encode()).hexdigest()

    await db.projects.update_one(
        {"_id": ObjectId(project_id)},
        {"$set": {"share_token": share_token, "updated_at": datetime.utcnow()}}
    )

    return {"share_token": share_token}


@router.get("/api/shared/{share_token}/slides")
async def get_shared_slides(share_token: str):
    """공유 링크로 슬라이드 조회 (인증 불필요)"""
    db = get_db()
    project = await db.projects.find_one({"share_token": share_token})
    if not project:
        raise HTTPException(status_code=404, detail="공유 링크를 찾을 수 없습니다")

    project_id = str(project["_id"])
    cursor = db.generated_slides.find({"project_id": project_id}).sort("order", 1)
    slides = []
    async for s in cursor:
        s["_id"] = str(s["_id"])
        slides.append(s)

    template = None
    if project.get("template_id"):
        template = await db.templates.find_one({"_id": ObjectId(project["template_id"])})
        if template:
            template["_id"] = str(template["_id"])

    return {
        "project_name": project.get("name"),
        "slides": slides,
        "template": template
    }


# ============ 유틸: 생성 취소 체크 ============

async def _check_cancelled(db, project_id: str, generation_id: str) -> bool:
    """MongoDB에서 생성 중단 요청 여부 확인"""
    project = await db.projects.find_one(
        {"_id": ObjectId(project_id)},
        {"status": 1, "generation_id": 1}
    )
    if not project:
        return True
    if project.get("status") == "stop_requested":
        return True
    # generation_id가 다르면 새 생성이 시작된 것 (재생성 케이스)
    if project.get("generation_id") != generation_id:
        return True
    return False


async def _set_stopped(db, project_id: str):
    """프로젝트 상태를 stopped로 변경"""
    await db.projects.update_one(
        {"_id": ObjectId(project_id)},
        {"$set": {"status": "stopped", "updated_at": datetime.utcnow()}}
    )


# ============ 유틸: 공통 헬퍼 ============

def _analyze_template_slides(slides: list) -> list[dict]:
    """템플릿 슬라이드 분석 및 자동 placeholder 할당"""
    slides_meta = []
    for idx, slide in enumerate(slides):
        placeholders = []
        role_counters = {}

        # 텍스트 오브젝트를 좌측 상단 → 우측 하단 순으로 정렬하여 처리
        # (y 좌표 우선, 같은 y면 x 좌표 순)
        text_objects = sorted(
            [o for o in slide.get("objects", []) if o.get("obj_type") == "text"],
            key=lambda o: (o.get("y", 0), o.get("x", 0))
        )

        for obj in text_objects:
            role = obj.get("role", "")
            placeholder = obj.get("placeholder", "")

            if not role:
                role = _infer_text_role(obj, slide.get("objects", []))
                obj["_auto_role"] = role

            if not placeholder:
                counter = role_counters.get(role, 0)
                placeholder = f"auto_{role}_{counter}"
                role_counters[role] = counter + 1
                obj["_auto_placeholder"] = placeholder
            else:
                role_counters[role] = role_counters.get(role, 0) + 1

            placeholders.append({"placeholder": placeholder, "role": role})

        enriched_meta = _enrich_slide_meta(slide, placeholders)
        slides_meta.append({
            "slide_index": idx,
            "slide_meta": enriched_meta,
            "placeholders": placeholders,
        })
    return slides_meta


def _build_gen_objects(template_slide: dict, contents: dict) -> list:
    """템플릿 슬라이드에서 generated objects 생성"""
    gen_objects = []
    for obj in template_slide.get("objects", []):
        gen_obj = obj.copy()
        if obj.get("obj_type") == "text":
            placeholder_name = obj.get("placeholder") or obj.get("_auto_placeholder", "")
            if placeholder_name:
                generated_text = contents.get(placeholder_name)
                if generated_text is None:
                    # 내용이 매핑되지 않은 subtitle/description/number 오브젝트 제거
                    role = obj.get("role") or obj.get("_auto_role", "")
                    if role in ("subtitle", "description", "number"):
                        continue
                    generated_text = obj.get("text_content", "")
                elif not generated_text:
                    generated_text = ""
                gen_obj["generated_text"] = generated_text
            # _auto_role을 role에 보존 (관리자가 명시적 role 미설정 시)
            if not gen_obj.get("role") and gen_obj.get("_auto_role"):
                gen_obj["role"] = gen_obj["_auto_role"]
            gen_obj.pop("_auto_placeholder", None)
            gen_obj.pop("_auto_role", None)
        gen_objects.append(gen_obj)
    return gen_objects


# ============ 유틸: 텍스트 역할 자동 추론 ============

def _infer_text_role(obj: dict, all_objects: list) -> str:
    """텍스트 오브젝트의 위치/크기/폰트에서 역할을 자동 추론

    추론 기준:
    - 큰 폰트(24pt+) → title
    - 상단(y<80)의 작은 텍스트(14pt 이하) → governance
    - 중간 폰트(18pt+)이고 상단(y<200) → subtitle
    - 나머지 → description
    """
    style = obj.get("text_style", {})
    font_size = style.get("font_size", 16)
    y = obj.get("y", 0)
    width = obj.get("width", 200)

    if font_size >= 24:
        return "title"
    if y < 80 and font_size <= 14:
        return "governance"
    if font_size >= 18 and y < 200:
        return "subtitle"
    return "description"


def _enrich_slide_meta(slide: dict, placeholders: list) -> dict:
    """슬라이드 오브젝트 분석으로 slide_meta 자동 보강

    관리자가 content_type, has_title 등을 설정하지 않았을 때
    오브젝트의 역할에서 자동으로 추론합니다.
    """
    meta = dict(slide.get("slide_meta", {}))
    roles = [ph.get("role", "") for ph in placeholders]

    # has_title 자동 감지
    if "has_title" not in meta:
        meta["has_title"] = "title" in roles

    # has_governance 자동 감지
    if "has_governance" not in meta:
        meta["has_governance"] = "governance" in roles

    # subtitle_count: 부제목 오브젝트 수
    if "subtitle_count" not in meta:
        meta["subtitle_count"] = roles.count("subtitle")

    # description_count: LLM이 생성할 items 수 (본문 슬라이드 최소 3개)
    content_type = meta.get("content_type", "body")
    admin_desc_count = meta.get("description_count", 0)

    if not admin_desc_count:
        # 관리자가 미설정 → 역할 기반 자동 계산
        desc_count = roles.count("description")
        body_count = roles.count("body")
        if desc_count == 0 and body_count > 0:
            desc_count = body_count
    else:
        desc_count = admin_desc_count

    # 본문 슬라이드는 최소 3개 items 보장 (풍부한 콘텐츠 생성)
    if content_type == "body":
        desc_count = max(desc_count, 3)

    meta["description_count"] = desc_count

    # content_type이 없으면 기본 body
    if not meta.get("content_type"):
        meta["content_type"] = "body"

    return meta


def _get_type_availability(slides_meta: list[dict]) -> dict:
    """템플릿에 등록된 슬라이드 타입 분석"""
    types = {
        "title_slide": {"label": "타이틀 슬라이드", "available": False, "indices": [], "count": 0},
        "toc": {"label": "목차 슬라이드", "available": False, "indices": [], "count": 0},
        "section_divider": {"label": "섹션 간지", "available": False, "indices": [], "count": 0},
        "body": {"label": "본문 슬라이드", "available": False, "indices": [], "count": 0},
        "closing": {"label": "마지막 슬라이드", "available": False, "indices": [], "count": 0},
    }

    for sm in slides_meta:
        ct = sm.get("slide_meta", {}).get("content_type", "body")
        if ct in types:
            types[ct]["available"] = True
            types[ct]["indices"].append(sm["slide_index"])
            types[ct]["count"] += 1

    return types


def _find_fallback_template(slides: list) -> int:
    """잘못된 template_index 대응: body 타입 템플릿 또는 첫 번째 템플릿 반환"""
    for idx, slide in enumerate(slides):
        if slide.get("slide_meta", {}).get("content_type") == "body":
            return idx
    return 0
