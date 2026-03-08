from fastapi import APIRouter, HTTPException, UploadFile, File, Form
from fastapi.responses import FileResponse, StreamingResponse
from bson import ObjectId
from datetime import datetime
from models.project import GenerateRequest, SlideReorderRequest, SlideUpdateRequest, ManualSlideRequest, SlideTextRequest, ExcelGenerateRequest, ExcelModifyRequest, ExcelChartRequest, DocxGenerateRequest, DocxModifyRequest
from services.mongo_service import get_db
from services.auth_service import decode_jwt_token, extract_user_key, get_user_flexible, get_user_by_key
from services import redis_service
from routers.collaboration import check_project_access
from services.template_service import recommend_slide, get_template_slides
from services.ppt_service import generate_pptx
from services.excel_service import generate_xlsx, auto_generate_chart_definition
from services.word_service import generate_docx
from services.llm_service import generate_slide_content, generate_slide_content_stream, generate_single_slide_content, generate_excel_content_stream, modify_excel_content_stream, generate_docx_content_stream
from services.search_service import search_web
from services.onlyoffice_service import create_onlyoffice_document
from services.file_service import extract_excel_structure
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
    # 협업 데이터 정리 (재생성 시 Lock/히스토리 리셋)
    await db.slide_locks.delete_many({"project_id": data.project_id})
    await db.slide_history.delete_many({"project_id": data.project_id})
    await redis_service.delete_project_locks(data.project_id)

    # 리소스 텍스트 수집
    cursor = db.resources.find({"project_id": data.project_id})
    all_content = []
    async for r in cursor:
        if r.get("content"):
            all_content.append(r["content"])
        elif r.get("title"):
            all_content.append(f"[{r['title']}]")

    combined_text = "\n\n".join(all_content)

    # 리소스가 없으면 지침으로 자동 웹 검색
    if not combined_text.strip():
        if not data.instructions.strip():
            raise HTTPException(status_code=400, detail="리소스 또는 지침을 입력하세요.")
        search_result = await search_web(data.instructions)
        pages = search_result.get("pages", [])
        if pages:
            search_contents = []
            for page in pages:
                title = page.get("title", "")
                content = page.get("content", "")
                if content:
                    search_contents.append(f"[{title}]\n{content}")
            combined_text = "\n\n".join(search_contents)

    if not combined_text.strip():
        raise HTTPException(status_code=400, detail="리소스에 텍스트 내용이 없습니다.")

    # 템플릿 슬라이드 조회
    slides = await get_template_slides(data.template_id)
    if not slides:
        raise HTTPException(status_code=400, detail="템플릿에 슬라이드가 없습니다.")

    # 배경 이미지
    template = await db.templates.find_one({"_id": ObjectId(data.template_id)})
    bg_image = template.get("background_image") if template else None

    # ── 핵심: 템플릿 슬라이드 분석 및 자동 placeholder 할당 (Redis 캐시) ──
    slides_meta = await _analyze_template_slides_cached(data.template_id, slides)

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
    # 협업 데이터 정리 (재생성 시 Lock/히스토리 리셋)
    await db.slide_locks.delete_many({"project_id": data.project_id})
    await db.slide_history.delete_many({"project_id": data.project_id})
    await redis_service.delete_project_locks(data.project_id)

    # 리소스 텍스트 수집
    cursor = db.resources.find({"project_id": data.project_id})
    all_content = []
    async for r in cursor:
        if r.get("content"):
            all_content.append(r["content"])
        elif r.get("title"):
            all_content.append(f"[{r['title']}]")
    combined_text = "\n\n".join(all_content)

    # 리소스가 없으면 지침(instructions)으로 자동 웹 검색
    web_search_performed = False
    if not combined_text.strip():
        if not data.instructions.strip():
            raise HTTPException(status_code=400, detail="리소스 또는 지침을 입력하세요.")
        search_result = await search_web(data.instructions)
        pages = search_result.get("pages", [])
        if pages:
            search_contents = []
            search_sources = []
            for page in pages:
                title = page.get("title", "")
                content = page.get("content", "")
                url = page.get("url", "")
                if content:
                    search_contents.append(f"[{title}]\n{content}")
                    if url:
                        search_sources.append(url)
            combined_text = "\n\n".join(search_contents)
            # 검색 결과를 리소스로 자동 저장
            if combined_text.strip():
                await db.resources.insert_one({
                    "project_id": data.project_id,
                    "resource_type": "web",
                    "title": f"자동 웹 검색: {data.instructions[:50]}",
                    "content": combined_text,
                    "sources": search_sources,
                    "created_at": datetime.utcnow(),
                })
                web_search_performed = True

    if not combined_text.strip():
        raise HTTPException(status_code=400, detail="리소스에 텍스트 내용이 없습니다.")

    # 템플릿 슬라이드 조회
    slides = await get_template_slides(data.template_id)
    if not slides:
        raise HTTPException(status_code=400, detail="템플릿에 슬라이드가 없습니다.")

    template = await db.templates.find_one({"_id": ObjectId(data.template_id)})
    bg_image = template.get("background_image") if template else None

    # 템플릿 슬라이드 분석 (Redis 캐시)
    slides_meta = await _analyze_template_slides_cached(data.template_id, slides)

    # 생성 ID (중복 생성 방지 및 취소 체크용)
    generation_id = str(uuid.uuid4())

    # 이전 취소 플래그 클리어
    await redis_service.clear_generation_cancel(data.project_id)

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

            if web_search_performed:
                yield _sse("start", {"message": "웹 검색으로 자료를 수집하여 슬라이드를 설계하고 있습니다..."})
            else:
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

            # ── Phase 1: 스켈레톤 전송 (배경 + 빈 텍스트) ──
            skeleton_slides = []
            for output_order, item in enumerate(llm_slides):
                template_idx = item.get("template_index", 0)
                if template_idx not in template_lookup:
                    template_idx = _find_fallback_template(slides)

                template_slide = template_lookup[template_idx]
                contents = item.get("contents", {})
                skeleton_objects = _build_skeleton_objects(template_slide, contents)

                slide_bg = template_slide.get("background_image") or bg_image

                skeleton_slides.append({
                    "order": output_order + 1,
                    "objects": skeleton_objects,
                    "items": [],
                    "background_image": slide_bg,
                })

            yield _sse("slides_skeleton", {"slides": skeleton_slides})

            # ── Phase 2: 슬라이드별 콘텐츠 전송 + DB 저장 ──
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

                yield _sse("slide_content", {"index": output_order, "slide": gen_slide})

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


@router.post("/{jwt_token}/api/generate/manual-slide")
async def add_manual_slide(jwt_token: str, data: ManualSlideRequest):
    """슬라이드 추가: 템플릿 슬라이드를 기반으로 빈 슬라이드 추가 (owner/editor)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()

    # 접근 제어: editor 이상만 허용
    await check_project_access(db, data.project_id, user_key, "editor")

    # 템플릿 슬라이드 조회
    template_slide = await db.slides.find_one({"_id": ObjectId(data.template_slide_id)})
    if not template_slide:
        raise HTTPException(status_code=404, detail="템플릿 슬라이드를 찾을 수 없습니다")

    now = datetime.utcnow()

    # 삽입 위치 결정
    if data.insert_after_order is not None:
        # 선택된 슬라이드 다음에 삽입: 이후 슬라이드 order +1 shift
        await db.generated_slides.update_many(
            {"project_id": data.project_id, "order": {"$gt": data.insert_after_order}},
            {"$inc": {"order": 1}, "$set": {"updated_at": now}}
        )
        new_order = data.insert_after_order + 1
    else:
        # 맨 끝에 추가
        count = await db.generated_slides.count_documents({"project_id": data.project_id})
        new_order = count + 1

    # 템플릿 분석 (auto role/placeholder 할당)
    analyzed = _analyze_template_slides([dict(template_slide)])
    slide_meta = analyzed[0] if analyzed else {}

    # 빈 generated_text로 objects 생성
    gen_objects = []
    for obj in template_slide.get("objects", []):
        gen_obj = obj.copy()
        gen_obj.pop("_id", None)
        if obj.get("obj_type") == "text":
            gen_obj["generated_text"] = ""
        gen_objects.append(gen_obj)

    doc = {
        "project_id": data.project_id,
        "template_slide_id": data.template_slide_id,
        "order": new_order,
        "objects": gen_objects,
        "items": [],
        "background_image": template_slide.get("background_image"),
        "slide_meta": template_slide.get("slide_meta", {}),
        "created_at": now,
        "updated_at": now,
    }
    result = await db.generated_slides.insert_one(doc)
    doc["_id"] = str(result.inserted_id)

    # order 정규화 (동시 삽입 시 중복 방지)
    await _normalize_slide_order(db, data.project_id)

    # 프로젝트 상태 업데이트
    existing_count = await db.generated_slides.count_documents({"project_id": data.project_id})
    update_fields = {
        "status": "generated",
        "updated_at": now,
    }
    # 첫 슬라이드일 때만 manual_mode 설정
    if existing_count == 1:
        update_fields["manual_mode"] = True

    await db.projects.update_one(
        {"_id": ObjectId(data.project_id)},
        {"$set": update_fields}
    )

    return {"slide": doc}


@router.post("/{jwt_token}/api/generate/slide-text")
async def generate_slide_text(jwt_token: str, data: SlideTextRequest):
    """수동 모드: 개별 슬라이드의 텍스트를 AI로 생성"""
    await get_user_key(jwt_token)
    db = get_db()

    # 생성된 슬라이드 조회
    gen_slide = await db.generated_slides.find_one({"_id": ObjectId(data.slide_id)})
    if not gen_slide:
        raise HTTPException(status_code=404, detail="슬라이드를 찾을 수 없습니다")

    # 템플릿 슬라이드 조회
    tmpl_slide_id = data.template_slide_id or gen_slide.get("template_slide_id", "")
    template_slide = None
    if tmpl_slide_id:
        template_slide = await db.slides.find_one({"_id": ObjectId(tmpl_slide_id)})

    if not template_slide:
        # 템플릿 없으면 gen_slide의 objects를 기반으로 분석
        template_slide = gen_slide

    # 리소스 텍스트 수집
    cursor = db.resources.find({"project_id": data.project_id})
    resources_text = ""
    async for r in cursor:
        resources_text += (r.get("content") or "") + "\n\n"

    # 템플릿 분석
    analyzed = _analyze_template_slides([dict(template_slide)])
    slide_meta = analyzed[0] if analyzed else {"placeholders": [], "slide_meta": {}}

    # 프로젝트 언어 확인
    project = await db.projects.find_one({"_id": ObjectId(data.project_id)})
    lang = (project or {}).get("lang", "ko") or "ko"

    # 현재 슬라이드의 기존 내용 수집 (편집 모드에서 전달)
    existing_content = None
    if data.current_content:
        existing_content = data.current_content
    else:
        # 프론트에서 전달하지 않은 경우 DB에서 수집
        ec = {}
        for obj in gen_slide.get("objects", []):
            if obj.get("obj_type") == "text":
                role = obj.get("role") or obj.get("_auto_role") or ""
                text = obj.get("generated_text") or obj.get("text_content") or ""
                if text:
                    ec[role or "text"] = text
        if ec:
            existing_content = {"contents": ec, "items": gen_slide.get("items", [])}

    # LLM으로 단일 슬라이드 텍스트 생성
    print(f"[slide-text] instruction: {data.instruction}")
    print(f"[slide-text] slide_meta placeholders: {slide_meta.get('placeholders', [])}")
    print(f"[slide-text] existing_content: {existing_content}")

    try:
        llm_result = await generate_single_slide_content(
            resources_text=resources_text,
            instruction=data.instruction,
            slide_meta=slide_meta,
            lang=lang,
            current_content=existing_content,
        )
    except Exception as e:
        print(f"[slide-text] LLM 호출 실패: {e}")
        raise HTTPException(status_code=500, detail=f"AI 텍스트 생성 실패: {str(e)}")

    contents = llm_result.get("contents", {})
    items = llm_result.get("items", [])
    print(f"[slide-text] LLM result contents: {contents}")
    print(f"[slide-text] LLM result items: {items}")

    # items를 subtitle/description placeholder에 매핑
    # (_build_gen_objects가 매핑되지 않은 subtitle/description 오브젝트를 제거하므로 반드시 필요)
    if items:
        sub_phs = [p["placeholder"] for p in slide_meta.get("placeholders", []) if p["role"] == "subtitle"]
        desc_phs = [p["placeholder"] for p in slide_meta.get("placeholders", []) if p["role"] == "description"]
        for i, item in enumerate(items):
            if i < len(sub_phs) and sub_phs[i] not in contents:
                contents[sub_phs[i]] = item.get("heading", "")
            if i < len(desc_phs) and desc_phs[i] not in contents:
                contents[desc_phs[i]] = item.get("detail", "")

    # _build_gen_objects로 텍스트 매핑
    gen_objects = _build_gen_objects(dict(template_slide), contents)
    print(f"[slide-text] gen_objects count: {len(gen_objects)}, text objs: {[(o.get('role',''), o.get('generated_text','')[:30]) for o in gen_objects if o.get('obj_type')=='text']}")

    # DB 업데이트
    await db.generated_slides.update_one(
        {"_id": ObjectId(data.slide_id)},
        {"$set": {
            "objects": gen_objects,
            "items": items,
            "updated_at": datetime.utcnow(),
        }}
    )

    return {
        "slide_id": data.slide_id,
        "objects": gen_objects,
        "items": items,
    }


@router.delete("/{jwt_token}/api/generate/manual-slide/{slide_id}")
async def delete_manual_slide(jwt_token: str, slide_id: str):
    """슬라이드 삭제 (owner/editor)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()

    slide = await db.generated_slides.find_one({"_id": ObjectId(slide_id)})
    if not slide:
        raise HTTPException(status_code=404, detail="슬라이드를 찾을 수 없습니다")

    project_id = slide.get("project_id")

    # 접근 제어: editor 이상만 허용
    await check_project_access(db, project_id, user_key, "editor")

    # 락 체크: 다른 사용자가 편집 중인 슬라이드는 삭제 불가
    now = datetime.utcnow()
    lock = await db.slide_locks.find_one({
        "project_id": project_id, "slide_id": slide_id
    })
    if lock and lock.get("expires_at") and lock["expires_at"] > now:
        if lock.get("user_key") != user_key:
            raise HTTPException(status_code=423, detail="다른 사용자가 편집 중이므로 삭제할 수 없습니다")

    await db.generated_slides.delete_one({"_id": ObjectId(slide_id)})

    # 순서 재정렬 + updated_at 갱신 (협업 감지용)
    await _normalize_slide_order(db, project_id)

    # 프로젝트 updated_at 갱신
    await db.projects.update_one(
        {"_id": ObjectId(project_id)},
        {"$set": {"updated_at": now}}
    )

    return {"success": True}


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

    # Redis에 취소 플래그 설정 (빠른 폴링용)
    await redis_service.set_generation_cancel(project_id)

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


@router.post("/{jwt_token}/api/generate/{project_id}/slides/delta")
async def get_slides_delta(jwt_token: str, project_id: str, data: dict):
    """변경된 슬라이드만 조회 (slide_ids 목록으로 필터)"""
    await get_user_key(jwt_token)
    db = get_db()
    slide_ids = data.get("slide_ids", [])
    if not slide_ids:
        return {"slides": []}
    object_ids = [ObjectId(sid) for sid in slide_ids]
    cursor = db.generated_slides.find(
        {"_id": {"$in": object_ids}, "project_id": project_id}
    ).sort("order", 1)
    slides = []
    async for s in cursor:
        s["_id"] = str(s["_id"])
        slides.append(s)
    return {"slides": slides}


@router.put("/{jwt_token}/api/generate/slides/{slide_id}")
async def update_generated_slide(jwt_token: str, slide_id: str, data: SlideUpdateRequest):
    """생성된 슬라이드 수정 (협업 시 잠금 확인 + 히스토리 기록)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()

    # 슬라이드 조회 (project_id 확인용)
    slide = await db.generated_slides.find_one({"_id": ObjectId(slide_id)})
    if not slide:
        raise HTTPException(status_code=404, detail="슬라이드를 찾을 수 없습니다")

    project_id = slide.get("project_id")

    # 권한 확인: 소유자 또는 editor 이상
    await check_project_access(db, project_id, user_key, "editor")

    # 협업 프로젝트면 Lock 확인
    collab_count = await db.collaborators.count_documents({"project_id": project_id})
    if collab_count > 0:
        lock = await db.slide_locks.find_one(
            {"project_id": project_id, "slide_id": slide_id}
        )
        if not lock or lock.get("user_key") != user_key:
            raise HTTPException(
                status_code=423,
                detail="슬라이드를 잠금한 후 편집할 수 있습니다"
            )

    # 변경 전 스냅샷
    before = {
        "objects": slide.get("objects", []),
        "items": slide.get("items", []),
    }

    # 업데이트 수행
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

    # 협업 프로젝트면 히스토리 기록
    if collab_count > 0:
        user = await get_user_by_key(user_key)
        user_name = user.get("nm", user_key) if user else user_key
        after = {
            "objects": data.objects,
            "items": data.items if data.items is not None else slide.get("items", []),
        }
        await _record_slide_history(
            db, project_id, slide_id,
            action="update", user_key=user_key, user_name=user_name,
            before_snapshot=before, after_snapshot=after,
            description="슬라이드 수정",
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


@router.post("/{jwt_token}/api/generate/switch-template-slide")
async def switch_template_slide(jwt_token: str, data: dict):
    """항목 수 변경 시 적합한 템플릿 슬라이드로 자동 전환

    요청: { slide_id, items_count, contents, items }
    - 현재 슬라이드의 template_id에서 동일 content_type + items_count에 가까운 템플릿 슬라이드를 찾아
      오브젝트를 재매핑하고 DB 업데이트
    """
    await get_user_key(jwt_token)
    db = get_db()

    slide_id = data.get("slide_id")
    new_items_count = data.get("items_count", 0)
    contents = data.get("contents", {})
    items = data.get("items", [])

    gen_slide = await db.generated_slides.find_one({"_id": ObjectId(slide_id)})
    if not gen_slide:
        raise HTTPException(status_code=404, detail="슬라이드를 찾을 수 없습니다")

    project_id = gen_slide.get("project_id")
    project = await db.projects.find_one({"_id": ObjectId(project_id)})
    if not project or not project.get("template_id"):
        return {"switched": False, "reason": "no_template"}

    template_id = project["template_id"]

    # 현재 템플릿 슬라이드의 content_type 확인
    old_tmpl_id = gen_slide.get("template_slide_id")
    old_tmpl = await db.slides.find_one({"_id": ObjectId(old_tmpl_id)}) if old_tmpl_id else None
    current_content_type = (old_tmpl or {}).get("slide_meta", {}).get("content_type", "body")

    # 동일 템플릿의 같은 content_type 슬라이드들 조회
    cursor = db.slides.find({"template_id": template_id})
    candidates = []
    async for tmpl_slide in cursor:
        meta = tmpl_slide.get("slide_meta", {})
        if meta.get("content_type") == current_content_type:
            desc_count = meta.get("description_count", 0)
            if not desc_count:
                # auto-detect from objects
                text_objs = [o for o in tmpl_slide.get("objects", []) if o.get("obj_type") == "text"]
                desc_count = sum(1 for o in text_objs if o.get("role") in ("description", "subtitle"))
            candidates.append({
                "slide": tmpl_slide,
                "desc_count": desc_count,
            })

    if not candidates:
        return {"switched": False, "reason": "no_candidates"}

    # 항목 수에 가장 가까운 템플릿 슬라이드 선택
    best = min(candidates, key=lambda c: abs(c["desc_count"] - new_items_count))
    new_tmpl = best["slide"]
    new_tmpl_id = str(new_tmpl["_id"])

    # 현재와 같은 템플릿이면 전환 불필요
    if new_tmpl_id == old_tmpl_id:
        return {"switched": False, "reason": "same_template"}

    # 새 템플릿으로 오브젝트 재매핑
    analyzed = _analyze_template_slides([dict(new_tmpl)])

    # items를 새 템플릿의 subtitle/description placeholder에 매핑
    if items and analyzed:
        sub_phs = [p["placeholder"] for p in analyzed[0].get("placeholders", []) if p["role"] == "subtitle"]
        desc_phs = [p["placeholder"] for p in analyzed[0].get("placeholders", []) if p["role"] == "description"]
        for i, item in enumerate(items):
            if i < len(sub_phs):
                contents[sub_phs[i]] = item.get("heading", "")
            if i < len(desc_phs):
                contents[desc_phs[i]] = item.get("detail", "")

    gen_objects = _build_gen_objects(dict(new_tmpl), contents)

    # DB 업데이트
    now = datetime.utcnow()
    await db.generated_slides.update_one(
        {"_id": ObjectId(slide_id)},
        {"$set": {
            "template_slide_id": new_tmpl_id,
            "objects": gen_objects,
            "items": items,
            "background_image": new_tmpl.get("background_image") or gen_slide.get("background_image"),
            "updated_at": now,
        }}
    )

    # 업데이트된 슬라이드 반환
    updated = await db.generated_slides.find_one({"_id": ObjectId(slide_id)})
    updated["_id"] = str(updated["_id"])

    return {"switched": True, "slide": updated, "new_template_slide_id": new_tmpl_id}


# ============ 엑셀 생성 ============

@router.post("/{jwt_token}/api/generate/excel/stream")
async def generate_excel_stream(jwt_token: str, data: ExcelGenerateRequest):
    """SSE 스트리밍으로 엑셀 데이터 생성 - 실시간 진행 표시"""
    await get_user_key(jwt_token)
    db = get_db()

    # 기존 생성 데이터 삭제
    await db.generated_excel.delete_many({"project_id": data.project_id})

    # 리소스 텍스트 수집
    cursor = db.resources.find({"project_id": data.project_id})
    all_content = []
    async for r in cursor:
        if r.get("content"):
            all_content.append(r["content"])
        elif r.get("title"):
            all_content.append(f"[{r['title']}]")
    combined_text = "\n\n".join(all_content)

    # 리소스가 없으면 instructions 필수 (인터넷 검색에 사용)
    needs_web_search = not combined_text.strip()
    if needs_web_search and not data.instructions.strip():
        raise HTTPException(status_code=400, detail="리소스가 없으면 지침을 입력해야 합니다.")

    generation_id = str(uuid.uuid4())
    await redis_service.clear_generation_cancel(data.project_id)

    # 프로젝트 상태 업데이트
    await db.projects.update_one(
        {"_id": ObjectId(data.project_id)},
        {"$set": {
            "instructions": data.instructions,
            "status": "generating",
            "generation_id": generation_id,
            "updated_at": datetime.utcnow(),
        }}
    )

    def _sse(event: str, payload: dict) -> str:
        return f"data: {json.dumps({'event': event, **payload}, ensure_ascii=False, default=str)}\n\n"

    async def event_stream():
        nonlocal combined_text, needs_web_search
        try:
            # 리소스가 없으면 인터넷 검색으로 자료 수집
            if needs_web_search:
                yield _sse("searching", {"message": "인터넷에서 자료를 검색하고 있습니다..."})

                search_result = await search_web(data.instructions)
                pages = search_result.get("pages", [])

                if pages:
                    search_contents = []
                    for page in pages:
                        title = page.get("title", "")
                        content = page.get("content", "")
                        url = page.get("url", "")
                        search_contents.append(f"[{title}]\n{content}\n(출처: {url})")
                    combined_text = "\n\n".join(search_contents)
                else:
                    # 검색 결과가 없으면 instructions만으로 진행
                    combined_text = data.instructions

                yield _sse("search_done", {"message": "검색 완료! 데이터를 생성합니다...", "result_count": len(pages)})

            yield _sse("start", {"message": "AI가 데이터를 구조화하고 있습니다..."})

            if await _check_cancelled(db, data.project_id, generation_id):
                yield _sse("stopped", {"message": "생성이 중단되었습니다."})
                await _set_stopped(db, data.project_id)
                return

            llm_result = None
            chunk_count = 0
            cancelled = False
            async for event_type, event_data in generate_excel_content_stream(
                combined_text, data.instructions, data.lang,
                sheet_count=data.sheet_count,
            ):
                if event_type == "delta":
                    yield _sse("delta", {"text": event_data})
                    chunk_count += 1
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
                yield _sse("error", {"message": "데이터 생성 실패"})
                return

            yield _sse("parsing", {"message": "데이터를 정리하고 있습니다..."})

            # MongoDB에 저장
            excel_doc = {
                "project_id": data.project_id,
                "sheets": llm_result.get("sheets", []),
                "meta": llm_result.get("meta", {}),
                "created_at": datetime.utcnow(),
                "updated_at": datetime.utcnow(),
            }
            await db.generated_excel.insert_one(excel_doc)
            excel_doc["_id"] = str(excel_doc["_id"])

            # 프로젝트 업데이트
            update_fields = {
                "status": "generated",
                "updated_at": datetime.utcnow(),
            }
            meta = llm_result.get("meta", {})
            if meta:
                update_fields["presentation_meta"] = meta

            await db.projects.update_one(
                {"_id": ObjectId(data.project_id)},
                {"$set": update_fields}
            )

            yield _sse("excel_data", {"excel": excel_doc})
            yield _sse("complete", {"total_sheets": len(llm_result.get("sheets", []))})

        except Exception as e:
            print(f"[SSE-Excel] Stream error: {e}")
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


@router.post("/{jwt_token}/api/generate/excel/modify/stream")
async def modify_excel_stream(jwt_token: str, data: ExcelModifyRequest):
    """SSE 스트리밍으로 기존 엑셀 데이터를 부분 수정"""
    await get_user_key(jwt_token)
    db = get_db()

    generation_id = str(uuid.uuid4())
    await redis_service.clear_generation_cancel(data.project_id)

    await db.projects.update_one(
        {"_id": ObjectId(data.project_id)},
        {"$set": {
            "status": "generating",
            "generation_id": generation_id,
            "updated_at": datetime.utcnow(),
        }}
    )

    def _sse(event: str, payload: dict) -> str:
        return f"data: {json.dumps({'event': event, **payload}, ensure_ascii=False, default=str)}\n\n"

    async def event_stream():
        try:
            yield _sse("start", {"message": "AI가 데이터를 수정하고 있습니다..."})

            if await _check_cancelled(db, data.project_id, generation_id):
                yield _sse("stopped", {"message": "수정이 중단되었습니다."})
                await _set_stopped(db, data.project_id)
                return

            llm_result = None
            chunk_count = 0
            cancelled = False
            async for event_type, event_data in modify_excel_content_stream(
                data.current_data, data.instruction, data.lang,
                target_sheet_index=data.target_sheet_index,
            ):
                if event_type == "delta":
                    yield _sse("delta", {"text": event_data})
                    chunk_count += 1
                    if chunk_count % 50 == 0:
                        if await _check_cancelled(db, data.project_id, generation_id):
                            cancelled = True
                            break
                elif event_type == "result":
                    llm_result = event_data

            if cancelled:
                yield _sse("stopped", {"message": "수정이 중단되었습니다."})
                await _set_stopped(db, data.project_id)
                return

            if not llm_result:
                yield _sse("error", {"message": "데이터 수정 실패"})
                return

            yield _sse("parsing", {"message": "수정된 데이터를 정리하고 있습니다..."})

            # 타겟 시트만 수정한 경우: 기존 시트 배열에 병합
            result_sheets = llm_result.get("sheets", [])
            original_sheets = data.current_data.get("sheets", [])

            if (data.target_sheet_index is not None
                    and 0 <= data.target_sheet_index < len(original_sheets)
                    and len(result_sheets) == 1):
                # LLM이 타겟 시트 1개만 반환 → 해당 인덱스만 교체
                merged_sheets = list(original_sheets)
                merged_sheets[data.target_sheet_index] = result_sheets[0]
                final_sheets = merged_sheets
                print(f"[Excel-Modify] 타겟 시트 병합: index={data.target_sheet_index}")
            else:
                # 전체 시트 교체 (시트 추가/삭제 또는 전체 수정)
                final_sheets = result_sheets

            await db.generated_excel.update_one(
                {"project_id": data.project_id},
                {"$set": {
                    "sheets": final_sheets,
                    "meta": llm_result.get("meta", {}) or data.current_data.get("meta", {}),
                    "updated_at": datetime.utcnow(),
                }},
                upsert=True
            )

            excel_doc = await db.generated_excel.find_one({"project_id": data.project_id})
            excel_doc["_id"] = str(excel_doc["_id"])

            update_fields = {
                "status": "generated",
                "updated_at": datetime.utcnow(),
            }
            meta = llm_result.get("meta", {})
            if meta:
                update_fields["presentation_meta"] = meta
            await db.projects.update_one(
                {"_id": ObjectId(data.project_id)},
                {"$set": update_fields}
            )

            yield _sse("excel_data", {"excel": excel_doc})
            yield _sse("complete", {"total_sheets": len(llm_result.get("sheets", []))})

        except Exception as e:
            print(f"[SSE-Excel-Modify] Stream error: {e}")
            project = await db.projects.find_one(
                {"_id": ObjectId(data.project_id)}, {"status": 1}
            )
            if project and project.get("status") in ("stop_requested", "stopped"):
                yield _sse("stopped", {"message": "수정이 중단되었습니다."})
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


@router.post("/{jwt_token}/api/generate/excel/upload")
async def upload_excel_file(jwt_token: str, project_id: str = Form(...), file: UploadFile = File(...)):
    """로컬 엑셀 파일 업로드 → generated_excel에 저장"""
    await get_user_key(jwt_token)
    db = get_db()

    # 파일 확장자 검증
    filename = file.filename or "upload.xlsx"
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext not in ("xlsx", "xls"):
        raise HTTPException(status_code=400, detail="xlsx 또는 xls 파일만 업로드 가능합니다")

    # 임시 파일 저장
    upload_dir = os.path.join(settings.UPLOAD_DIR, "temp")
    os.makedirs(upload_dir, exist_ok=True)
    temp_filename = f"{uuid.uuid4().hex}.{ext}"
    temp_path = os.path.join(upload_dir, temp_filename)

    try:
        content = await file.read()
        with open(temp_path, "wb") as f:
            f.write(content)

        # 엑셀 구조 파싱
        excel_data = extract_excel_structure(temp_path, original_filename=filename)

        # generated_excel에 저장 (upsert)
        await db.generated_excel.update_one(
            {"project_id": project_id},
            {"$set": {
                "project_id": project_id,
                "sheets": excel_data.get("sheets", []),
                "meta": excel_data.get("meta", {}),
                "created_at": datetime.utcnow(),
                "updated_at": datetime.utcnow(),
            }},
            upsert=True,
        )

        # 프로젝트 상태 업데이트
        update_fields = {
            "status": "generated",
            "updated_at": datetime.utcnow(),
        }
        meta = excel_data.get("meta", {})
        if meta:
            update_fields["presentation_meta"] = meta
        await db.projects.update_one(
            {"_id": ObjectId(project_id)},
            {"$set": update_fields},
        )

        excel_doc = await db.generated_excel.find_one({"project_id": project_id})
        excel_doc["_id"] = str(excel_doc["_id"])

        return {"success": True, "excel": excel_doc}

    finally:
        # 임시 파일 삭제
        if os.path.exists(temp_path):
            os.remove(temp_path)


@router.post("/{jwt_token}/api/generate/excel/chart")
async def generate_excel_chart(jwt_token: str, data: ExcelChartRequest):
    """시트 데이터에서 직접 차트 생성 (LLM 없이, 즉시 생성)"""
    await get_user_key(jwt_token)
    db = get_db()

    excel_doc = await db.generated_excel.find_one({"project_id": data.project_id})
    if not excel_doc:
        raise HTTPException(status_code=404, detail="엑셀 데이터가 없습니다")

    sheets = excel_doc.get("sheets", [])
    if data.sheet_index < 0 or data.sheet_index >= len(sheets):
        raise HTTPException(status_code=400, detail=f"시트 인덱스 범위 초과 (0~{len(sheets)-1})")

    sheet = sheets[data.sheet_index]

    valid_types = {"bar", "line", "pie", "area", "scatter", "doughnut", "radar"}
    chart_type = data.chart_type if data.chart_type in valid_types else "bar"

    # 차트 정의 자동 생성
    chart_def = auto_generate_chart_definition(sheet, chart_type, data.title)
    if not chart_def:
        raise HTTPException(status_code=400, detail="차트 생성 불가: 숫자 데이터가 있는 열이 없습니다")

    # 시트에 차트 추가
    if "charts" not in sheet:
        sheet["charts"] = []
    sheet["charts"].append(chart_def)

    # 시트당 최대 3개 제한
    if len(sheet["charts"]) > 3:
        sheet["charts"] = sheet["charts"][-3:]

    # DB 업데이트
    await db.generated_excel.update_one(
        {"project_id": data.project_id},
        {"$set": {
            "sheets": sheets,
            "updated_at": datetime.utcnow(),
        }},
    )

    excel_doc = await db.generated_excel.find_one({"project_id": data.project_id})
    excel_doc["_id"] = str(excel_doc["_id"])

    return {"success": True, "excel": excel_doc}


@router.get("/{jwt_token}/api/generate/{project_id}/excel")
async def get_generated_excel(jwt_token: str, project_id: str):
    """생성된 엑셀 데이터 조회"""
    await get_user_key(jwt_token)
    db = get_db()
    doc = await db.generated_excel.find_one({"project_id": project_id})
    if not doc:
        return {"excel": None}
    doc["_id"] = str(doc["_id"])
    return {"excel": doc}


@router.put("/{jwt_token}/api/generate/{project_id}/excel")
async def update_excel_data(jwt_token: str, project_id: str, data: dict):
    """편집된 엑셀 데이터 저장"""
    await get_user_key(jwt_token)
    db = get_db()
    await db.generated_excel.update_one(
        {"project_id": project_id},
        {"$set": {
            "sheets": data.get("sheets", []),
            "updated_at": datetime.utcnow(),
        }},
        upsert=True
    )
    return {"success": True}


@router.get("/{jwt_token}/api/generate/{project_id}/download/xlsx")
async def download_xlsx(jwt_token: str, project_id: str):
    """XLSX 파일 다운로드"""
    await get_user_key(jwt_token)
    try:
        file_url = await generate_xlsx(project_id)
        file_path = os.path.join(".", file_url.lstrip("/"))
        if not os.path.exists(file_path):
            raise HTTPException(status_code=500, detail="파일 생성 실패")

        db = get_db()
        project = await db.projects.find_one({"_id": ObjectId(project_id)})
        filename = f"{project.get('name', 'data')}.xlsx"

        return FileResponse(
            path=file_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.get("/{jwt_token}/api/generate/{project_id}/download/pdf")
async def download_pdf(jwt_token: str, project_id: str):
    """PDF 다운로드 (미구현)"""
    await get_user_key(jwt_token)
    raise HTTPException(status_code=501, detail="PDF 변환 기능은 준비 중입니다")


# ============ OnlyOffice 생성 ============

@router.post("/{jwt_token}/api/generate/onlyoffice/pptx/stream")
async def generate_onlyoffice_pptx_stream(jwt_token: str, data: GenerateRequest):
    """OnlyOffice PPTX: AI 슬라이드 생성 → PPTX 파일 → OnlyOffice 에디터"""
    await get_user_key(jwt_token)
    db = get_db()

    # 기존 데이터 정리
    await db.generated_slides.delete_many({"project_id": data.project_id})

    # 리소스 수집
    cursor = db.resources.find({"project_id": data.project_id})
    all_content = []
    async for r in cursor:
        if r.get("content"):
            all_content.append(r["content"])
        elif r.get("title"):
            all_content.append(f"[{r['title']}]")
    combined_text = "\n\n".join(all_content)

    if not combined_text.strip() and not data.instructions.strip():
        raise HTTPException(status_code=400, detail="리소스 또는 지침이 필요합니다.")

    # 리소스가 없으면 지침으로 자동 웹 검색
    if not combined_text.strip() and data.instructions.strip():
        search_result = await search_web(data.instructions)
        pages = search_result.get("pages", [])
        if pages:
            search_contents = []
            search_sources = []
            for page in pages:
                title = page.get("title", "")
                content = page.get("content", "")
                url = page.get("url", "")
                if content:
                    search_contents.append(f"[{title}]\n{content}")
                    if url:
                        search_sources.append(url)
            combined_text = "\n\n".join(search_contents)
            if combined_text.strip():
                await db.resources.insert_one({
                    "project_id": data.project_id,
                    "resource_type": "web",
                    "title": f"자동 웹 검색: {data.instructions[:50]}",
                    "content": combined_text,
                    "sources": search_sources,
                    "created_at": datetime.utcnow(),
                })

    # 템플릿 슬라이드 분석
    slides = await get_template_slides(data.template_id)
    if not slides:
        raise HTTPException(status_code=400, detail="템플릿 슬라이드가 없습니다")

    slides_meta = await _analyze_template_slides_cached(data.template_id, slides)

    generation_id = str(uuid.uuid4())
    await redis_service.clear_generation_cancel(data.project_id)

    await db.projects.update_one(
        {"_id": ObjectId(data.project_id)},
        {"$set": {
            "instructions": data.instructions,
            "template_id": data.template_id,
            "status": "generating",
            "generation_id": generation_id,
            "updated_at": datetime.utcnow(),
        }}
    )

    def _sse(event: str, payload: dict) -> str:
        return f"data: {json.dumps({'event': event, **payload}, ensure_ascii=False, default=str)}\n\n"

    async def event_stream():
        try:
            yield _sse("start", {"message": "AI가 프레젠테이션을 설계하고 있습니다..."})

            llm_result = None
            chunk_count = 0
            cancelled = False
            async for event_type, event_data in generate_slide_content_stream(
                combined_text, data.instructions, slides_meta,
                lang=data.lang, slide_count=data.slide_count,
            ):
                if event_type == "delta":
                    yield _sse("delta", {"text": event_data})
                    chunk_count += 1
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

            if not llm_result or not llm_result.get("slides"):
                yield _sse("error", {"message": "슬라이드 생성 실패"})
                return

            yield _sse("parsing", {"message": "슬라이드를 구성하고 있습니다..."})

            # 리치 스키마 결과 처리 (기존 generate_slides_stream과 동일 패턴)
            llm_slides = llm_result.get("slides", [])
            template_lookup = {idx: slide for idx, slide in enumerate(slides)}

            for output_order, item in enumerate(llm_slides):
                # Extract slide title
                contents = item.get("contents", {})
                slide_title = ""
                if isinstance(contents, dict):
                    title_obj = contents.get("title", {})
                    if isinstance(title_obj, dict):
                        slide_title = title_obj.get("text", "")
                    elif isinstance(title_obj, str):
                        slide_title = title_obj
                yield _sse("item_progress", {
                    "current": output_order + 1,
                    "total": len(llm_slides),
                    "title": slide_title or f"슬라이드 {output_order + 1}",
                    "type": "slide"
                })

                template_idx = item.get("template_index", 0)
                if template_idx not in template_lookup:
                    template_idx = _find_fallback_template(slides)

                template_slide = template_lookup[template_idx]
                contents = item.get("contents", {})
                gen_objects = _build_gen_objects(template_slide, contents)

                slide_bg = template_slide.get("background_image", "")

                gen_slide_doc = {
                    "project_id": data.project_id,
                    "template_slide_id": str(template_slide.get("_id", "")),
                    "order": output_order + 1,
                    "objects": gen_objects,
                    "items": item.get("items", []),
                    "background_image": slide_bg,
                    "created_at": datetime.utcnow(),
                    "updated_at": datetime.utcnow(),
                }
                await db.generated_slides.insert_one(gen_slide_doc)

            # 프로젝트 메타 업데이트
            meta = llm_result.get("meta", {})
            update_fields = {"status": "generated", "updated_at": datetime.utcnow()}
            if meta:
                update_fields["presentation_meta"] = meta
            sources = llm_result.get("sources", [])
            if sources:
                update_fields["presentation_sources"] = sources
            await db.projects.update_one(
                {"_id": ObjectId(data.project_id)},
                {"$set": update_fields}
            )

            # PPTX 파일 생성 + OnlyOffice 등록
            yield _sse("file_creating", {"message": "파일을 생성하고 있습니다..."})
            try:
                file_url = await generate_pptx(data.project_id)
                oo_doc = await create_onlyoffice_document(data.project_id, "pptx", file_url)
                yield _sse("onlyoffice_ready", {"project_id": data.project_id, "document": oo_doc})
            except Exception as e:
                print(f"[OnlyOffice-PPTX] 파일 생성 실패: {e}")
                yield _sse("error", {"message": f"파일 생성 실패: {str(e)}"})
                return

            yield _sse("complete", {"total_slides": len(llm_slides)})

        except Exception as e:
            print(f"[SSE-OO-PPTX] Stream error: {e}")
            import traceback
            traceback.print_exc()
            yield _sse("error", {"message": str(e)})

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no", "Connection": "keep-alive"},
    )


@router.post("/{jwt_token}/api/generate/onlyoffice/xlsx/stream")
async def generate_onlyoffice_xlsx_stream(jwt_token: str, data: ExcelGenerateRequest):
    """OnlyOffice XLSX: AI 엑셀 생성 → XLSX 파일 → OnlyOffice 에디터"""
    await get_user_key(jwt_token)
    db = get_db()

    await db.generated_excel.delete_many({"project_id": data.project_id})

    cursor = db.resources.find({"project_id": data.project_id})
    all_content = []
    async for r in cursor:
        if r.get("content"):
            all_content.append(r["content"])
        elif r.get("title"):
            all_content.append(f"[{r['title']}]")
    combined_text = "\n\n".join(all_content)

    needs_web_search = not combined_text.strip()
    if needs_web_search and not data.instructions.strip():
        raise HTTPException(status_code=400, detail="리소스가 없으면 지침을 입력해야 합니다.")

    generation_id = str(uuid.uuid4())
    await redis_service.clear_generation_cancel(data.project_id)

    await db.projects.update_one(
        {"_id": ObjectId(data.project_id)},
        {"$set": {
            "instructions": data.instructions,
            "status": "generating",
            "generation_id": generation_id,
            "updated_at": datetime.utcnow(),
        }}
    )

    def _sse(event: str, payload: dict) -> str:
        return f"data: {json.dumps({'event': event, **payload}, ensure_ascii=False, default=str)}\n\n"

    async def event_stream():
        nonlocal combined_text, needs_web_search
        try:
            if needs_web_search:
                yield _sse("searching", {"message": "인터넷에서 자료를 검색하고 있습니다..."})
                search_result = await search_web(data.instructions)
                pages = search_result.get("pages", [])
                if pages:
                    search_contents = []
                    for page in pages:
                        search_contents.append(f"[{page.get('title', '')}]\n{page.get('content', '')}\n(출처: {page.get('url', '')})")
                    combined_text = "\n\n".join(search_contents)
                else:
                    combined_text = data.instructions
                yield _sse("search_done", {"message": "검색 완료!", "result_count": len(pages)})

            yield _sse("start", {"message": "AI가 데이터를 구조화하고 있습니다..."})

            if await _check_cancelled(db, data.project_id, generation_id):
                yield _sse("stopped", {"message": "생성이 중단되었습니다."})
                await _set_stopped(db, data.project_id)
                return

            llm_result = None
            chunk_count = 0
            cancelled = False
            async for event_type, event_data in generate_excel_content_stream(
                combined_text, data.instructions, data.lang,
                sheet_count=data.sheet_count,
            ):
                if event_type == "delta":
                    yield _sse("delta", {"text": event_data})
                    chunk_count += 1
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
                yield _sse("error", {"message": "데이터 생성 실패"})
                return

            yield _sse("parsing", {"message": "데이터를 정리하고 있습니다..."})

            sheets = llm_result.get("sheets", [])
            for idx, sheet in enumerate(sheets):
                row_count = len(sheet.get("rows", []))
                col_count = len(sheet.get("columns", []))
                chart_count = len(sheet.get("charts", []))
                detail_parts = [f"{row_count}행 x {col_count}열"]
                if chart_count > 0:
                    detail_parts.append(f"차트 {chart_count}개")
                yield _sse("item_progress", {
                    "current": idx + 1,
                    "total": len(sheets),
                    "title": sheet.get("name", f"시트 {idx + 1}"),
                    "type": "sheet",
                    "detail": ", ".join(detail_parts)
                })

            excel_doc = {
                "project_id": data.project_id,
                "sheets": llm_result.get("sheets", []),
                "meta": llm_result.get("meta", {}),
                "created_at": datetime.utcnow(),
                "updated_at": datetime.utcnow(),
            }
            await db.generated_excel.insert_one(excel_doc)

            update_fields = {"status": "generated", "updated_at": datetime.utcnow()}
            meta = llm_result.get("meta", {})
            if meta:
                update_fields["presentation_meta"] = meta
            await db.projects.update_one(
                {"_id": ObjectId(data.project_id)},
                {"$set": update_fields}
            )

            # XLSX 파일 생성 + OnlyOffice 등록
            yield _sse("file_creating", {"message": "파일을 생성하고 있습니다..."})
            try:
                file_url = await generate_xlsx(data.project_id)
                print(f"[OnlyOffice-XLSX] XLSX 파일 생성 완료: {file_url}")
                oo_doc = await create_onlyoffice_document(data.project_id, "xlsx", file_url)
                print(f"[OnlyOffice-XLSX] OnlyOffice 문서 등록 완료: key={oo_doc.get('document_key', 'N/A')}")
                yield _sse("onlyoffice_ready", {"project_id": data.project_id, "document": oo_doc})
            except Exception as e:
                print(f"[OnlyOffice-XLSX] 파일 생성 실패: {e}")
                import traceback
                traceback.print_exc()
                yield _sse("error", {"message": f"파일 생성 실패: {str(e)}"})
                return

            yield _sse("complete", {"total_sheets": len(llm_result.get("sheets", []))})

        except Exception as e:
            print(f"[SSE-OO-XLSX] Stream error: {e}")
            import traceback
            traceback.print_exc()
            yield _sse("error", {"message": str(e)})

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no", "Connection": "keep-alive"},
    )


@router.post("/{jwt_token}/api/generate/onlyoffice/docx/stream")
async def generate_onlyoffice_docx_stream(jwt_token: str, data: DocxGenerateRequest):
    """OnlyOffice DOCX: AI 문서 생성 → DOCX 파일 → OnlyOffice 에디터"""
    await get_user_key(jwt_token)
    db = get_db()

    await db.generated_docx.delete_many({"project_id": data.project_id})

    cursor = db.resources.find({"project_id": data.project_id})
    all_content = []
    async for r in cursor:
        if r.get("content"):
            all_content.append(r["content"])
        elif r.get("title"):
            all_content.append(f"[{r['title']}]")
    combined_text = "\n\n".join(all_content)

    needs_web_search = not combined_text.strip()
    if needs_web_search and not data.instructions.strip():
        raise HTTPException(status_code=400, detail="리소스가 없으면 지침을 입력해야 합니다.")

    generation_id = str(uuid.uuid4())
    await redis_service.clear_generation_cancel(data.project_id)

    await db.projects.update_one(
        {"_id": ObjectId(data.project_id)},
        {"$set": {
            "instructions": data.instructions,
            "status": "generating",
            "generation_id": generation_id,
            "updated_at": datetime.utcnow(),
        }}
    )

    def _sse(event: str, payload: dict) -> str:
        return f"data: {json.dumps({'event': event, **payload}, ensure_ascii=False, default=str)}\n\n"

    async def event_stream():
        nonlocal combined_text, needs_web_search
        try:
            if needs_web_search:
                yield _sse("searching", {"message": "인터넷에서 자료를 검색하고 있습니다..."})
                search_result = await search_web(data.instructions)
                pages = search_result.get("pages", [])
                if pages:
                    search_contents = []
                    for page in pages:
                        search_contents.append(f"[{page.get('title', '')}]\n{page.get('content', '')}\n(출처: {page.get('url', '')})")
                    combined_text = "\n\n".join(search_contents)
                else:
                    combined_text = data.instructions
                yield _sse("search_done", {"message": "검색 완료!", "result_count": len(pages)})

            yield _sse("start", {"message": "AI가 문서를 작성하고 있습니다..."})

            if await _check_cancelled(db, data.project_id, generation_id):
                yield _sse("stopped", {"message": "생성이 중단되었습니다."})
                await _set_stopped(db, data.project_id)
                return

            llm_result = None
            chunk_count = 0
            cancelled = False
            async for event_type, event_data in generate_docx_content_stream(
                combined_text, data.instructions, data.lang,
                section_count=data.section_count,
            ):
                if event_type == "delta":
                    yield _sse("delta", {"text": event_data})
                    chunk_count += 1
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
                yield _sse("error", {"message": "문서 생성 실패"})
                return

            yield _sse("parsing", {"message": "문서를 구성하고 있습니다..."})

            sections = llm_result.get("sections", [])
            for idx, section in enumerate(sections):
                yield _sse("item_progress", {
                    "current": idx + 1,
                    "total": len(sections),
                    "title": section.get("title", f"섹션 {idx + 1}"),
                    "type": "section"
                })

            docx_doc = {
                "project_id": data.project_id,
                "sections": llm_result.get("sections", []),
                "meta": llm_result.get("meta", {}),
                "created_at": datetime.utcnow(),
                "updated_at": datetime.utcnow(),
            }
            await db.generated_docx.insert_one(docx_doc)

            update_fields = {"status": "generated", "updated_at": datetime.utcnow()}
            meta = llm_result.get("meta", {})
            if meta:
                update_fields["presentation_meta"] = meta
            await db.projects.update_one(
                {"_id": ObjectId(data.project_id)},
                {"$set": update_fields}
            )

            # DOCX 파일 생성 + OnlyOffice 등록
            yield _sse("file_creating", {"message": "파일을 생성하고 있습니다..."})
            try:
                file_url = await generate_docx(data.project_id)
                oo_doc = await create_onlyoffice_document(data.project_id, "docx", file_url)
                yield _sse("onlyoffice_ready", {"project_id": data.project_id, "document": oo_doc})
            except Exception as e:
                print(f"[OnlyOffice-DOCX] 파일 생성 실패: {e}")
                yield _sse("error", {"message": f"파일 생성 실패: {str(e)}"})
                return

            yield _sse("complete", {"total_sections": len(llm_result.get("sections", []))})

        except Exception as e:
            print(f"[SSE-OO-DOCX] Stream error: {e}")
            import traceback
            traceback.print_exc()
            yield _sse("error", {"message": str(e)})

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no", "Connection": "keep-alive"},
    )


@router.get("/{jwt_token}/api/generate/{project_id}/download/onlyoffice")
async def download_onlyoffice_document(jwt_token: str, project_id: str):
    """OnlyOffice 문서 다운로드 (최신 편집본)"""
    await get_user_key(jwt_token)
    db = get_db()

    oo_doc = await db.onlyoffice_documents.find_one({"project_id": project_id})
    if not oo_doc:
        raise HTTPException(status_code=404, detail="OnlyOffice 문서가 없습니다")

    file_path = oo_doc.get("file_path", "")
    if file_path.startswith("/uploads/"):
        abs_path = os.path.join(settings.UPLOAD_DIR, file_path[len("/uploads/"):])
    else:
        abs_path = file_path

    if not os.path.exists(abs_path):
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다")

    project = await db.projects.find_one({"_id": ObjectId(project_id)})
    project_name = project.get("name", "document") if project else "document"
    file_type = oo_doc.get("file_type", "docx")

    mime_map = {
        "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }

    return FileResponse(
        path=abs_path,
        filename=f"{project_name}.{file_type}",
        media_type=mime_map.get(file_type, "application/octet-stream"),
    )


# ============ Word 문서 생성 ============

@router.post("/{jwt_token}/api/generate/docx/stream")
async def generate_docx_stream(jwt_token: str, data: DocxGenerateRequest):
    """SSE 스트리밍으로 Word 문서 데이터 생성 - 실시간 진행 표시"""
    await get_user_key(jwt_token)
    db = get_db()

    # 기존 생성 데이터 삭제
    await db.generated_docx.delete_many({"project_id": data.project_id})

    # 리소스 텍스트 수집
    cursor = db.resources.find({"project_id": data.project_id})
    all_content = []
    async for r in cursor:
        if r.get("content"):
            all_content.append(r["content"])
        elif r.get("title"):
            all_content.append(f"[{r['title']}]")
    combined_text = "\n\n".join(all_content)

    # 리소스가 없으면 instructions 필수 (인터넷 검색에 사용)
    needs_web_search = not combined_text.strip()
    if needs_web_search and not data.instructions.strip():
        raise HTTPException(status_code=400, detail="리소스가 없으면 지침을 입력해야 합니다.")

    generation_id = str(uuid.uuid4())
    await redis_service.clear_generation_cancel(data.project_id)

    # 프로젝트 상태 업데이트
    await db.projects.update_one(
        {"_id": ObjectId(data.project_id)},
        {"$set": {
            "instructions": data.instructions,
            "status": "generating",
            "generation_id": generation_id,
            "updated_at": datetime.utcnow(),
        }}
    )

    def _sse(event: str, payload: dict) -> str:
        return f"data: {json.dumps({'event': event, **payload}, ensure_ascii=False, default=str)}\n\n"

    async def event_stream():
        nonlocal combined_text, needs_web_search
        try:
            # 리소스가 없으면 인터넷 검색으로 자료 수집
            if needs_web_search:
                yield _sse("searching", {"message": "인터넷에서 자료를 검색하고 있습니다..."})

                search_result = await search_web(data.instructions)
                pages = search_result.get("pages", [])

                if pages:
                    search_contents = []
                    for page in pages:
                        title = page.get("title", "")
                        content = page.get("content", "")
                        url = page.get("url", "")
                        search_contents.append(f"[{title}]\n{content}\n(출처: {url})")
                    combined_text = "\n\n".join(search_contents)
                else:
                    # 검색 결과가 없으면 instructions만으로 진행
                    combined_text = data.instructions

                yield _sse("search_done", {"message": "검색 완료! 문서를 생성합니다...", "result_count": len(pages)})

            yield _sse("start", {"message": "AI가 문서를 구조화하고 있습니다..."})

            if await _check_cancelled(db, data.project_id, generation_id):
                yield _sse("stopped", {"message": "생성이 중단되었습니다."})
                await _set_stopped(db, data.project_id)
                return

            llm_result = None
            chunk_count = 0
            cancelled = False
            async for event_type, event_data in generate_docx_content_stream(
                combined_text, data.instructions, data.lang,
                section_count=data.section_count,
            ):
                if event_type == "delta":
                    yield _sse("delta", {"text": event_data})
                    chunk_count += 1
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
                yield _sse("error", {"message": "문서 생성 실패"})
                return

            yield _sse("parsing", {"message": "문서를 정리하고 있습니다..."})

            # MongoDB에 저장
            docx_doc = {
                "project_id": data.project_id,
                "sections": llm_result.get("sections", []),
                "meta": llm_result.get("meta", {}),
                "created_at": datetime.utcnow(),
                "updated_at": datetime.utcnow(),
            }
            await db.generated_docx.insert_one(docx_doc)
            docx_doc["_id"] = str(docx_doc["_id"])

            # 프로젝트 업데이트
            update_fields = {
                "status": "generated",
                "updated_at": datetime.utcnow(),
            }
            meta = llm_result.get("meta", {})
            if meta:
                update_fields["presentation_meta"] = meta

            await db.projects.update_one(
                {"_id": ObjectId(data.project_id)},
                {"$set": update_fields}
            )

            yield _sse("docx_data", {"docx": docx_doc})
            yield _sse("complete", {"total_sections": len(llm_result.get("sections", []))})

        except Exception as e:
            print(f"[SSE-Docx] Stream error: {e}")
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


@router.post("/{jwt_token}/api/generate/docx/modify/stream")
async def modify_docx_stream(jwt_token: str, data: DocxModifyRequest):
    """SSE 스트리밍으로 기존 Word 문서를 수정"""
    await get_user_key(jwt_token)
    db = get_db()

    generation_id = str(uuid.uuid4())
    await redis_service.clear_generation_cancel(data.project_id)

    await db.projects.update_one(
        {"_id": ObjectId(data.project_id)},
        {"$set": {
            "status": "generating",
            "generation_id": generation_id,
            "updated_at": datetime.utcnow(),
        }}
    )

    def _sse(event: str, payload: dict) -> str:
        return f"data: {json.dumps({'event': event, **payload}, ensure_ascii=False, default=str)}\n\n"

    async def event_stream():
        try:
            yield _sse("start", {"message": "AI가 문서를 수정하고 있습니다..."})

            if await _check_cancelled(db, data.project_id, generation_id):
                yield _sse("stopped", {"message": "수정이 중단되었습니다."})
                await _set_stopped(db, data.project_id)
                return

            # 현재 문서 내용을 리소스 텍스트로 구성
            current_sections = data.current_data.get("sections", [])
            current_text_parts = []
            for section in current_sections:
                title = section.get("title", "")
                content = section.get("content", "")
                current_text_parts.append(f"## {title}\n{content}")
            current_text = "\n\n".join(current_text_parts)

            combined_text = f"[현재 문서 내용]\n{current_text}\n\n[수정 지침]\n{data.instruction}"

            llm_result = None
            chunk_count = 0
            cancelled = False
            async for event_type, event_data in generate_docx_content_stream(
                combined_text, data.instruction, data.lang,
            ):
                if event_type == "delta":
                    yield _sse("delta", {"text": event_data})
                    chunk_count += 1
                    if chunk_count % 50 == 0:
                        if await _check_cancelled(db, data.project_id, generation_id):
                            cancelled = True
                            break
                elif event_type == "result":
                    llm_result = event_data

            if cancelled:
                yield _sse("stopped", {"message": "수정이 중단되었습니다."})
                await _set_stopped(db, data.project_id)
                return

            if not llm_result:
                yield _sse("error", {"message": "문서 수정 실패"})
                return

            yield _sse("parsing", {"message": "수정된 문서를 정리하고 있습니다..."})

            # MongoDB에 upsert 저장
            await db.generated_docx.update_one(
                {"project_id": data.project_id},
                {"$set": {
                    "sections": llm_result.get("sections", []),
                    "meta": llm_result.get("meta", {}) or data.current_data.get("meta", {}),
                    "updated_at": datetime.utcnow(),
                }},
                upsert=True
            )

            docx_doc = await db.generated_docx.find_one({"project_id": data.project_id})
            docx_doc["_id"] = str(docx_doc["_id"])

            update_fields = {
                "status": "generated",
                "updated_at": datetime.utcnow(),
            }
            meta = llm_result.get("meta", {})
            if meta:
                update_fields["presentation_meta"] = meta
            await db.projects.update_one(
                {"_id": ObjectId(data.project_id)},
                {"$set": update_fields}
            )

            yield _sse("docx_data", {"docx": docx_doc})
            yield _sse("complete", {"total_sections": len(llm_result.get("sections", []))})

        except Exception as e:
            print(f"[SSE-Docx-Modify] Stream error: {e}")
            project = await db.projects.find_one(
                {"_id": ObjectId(data.project_id)}, {"status": 1}
            )
            if project and project.get("status") in ("stop_requested", "stopped"):
                yield _sse("stopped", {"message": "수정이 중단되었습니다."})
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


@router.get("/{jwt_token}/api/generate/{project_id}/docx")
async def get_docx_data(jwt_token: str, project_id: str):
    """Word 문서 데이터 조회"""
    await get_user_key(jwt_token)
    db = get_db()
    docx_data = await db.generated_docx.find_one({"project_id": project_id})
    if not docx_data:
        raise HTTPException(status_code=404, detail="생성된 문서가 없습니다")
    docx_data["_id"] = str(docx_data["_id"])
    return {"docx": docx_data}


@router.put("/{jwt_token}/api/generate/{project_id}/docx")
async def save_docx_data(jwt_token: str, project_id: str, data: dict):
    """Word 문서 데이터 저장"""
    await get_user_key(jwt_token)
    db = get_db()
    await db.generated_docx.update_one(
        {"project_id": project_id},
        {"$set": {
            "sections": data.get("sections", []),
            "meta": data.get("meta", {}),
            "updated_at": datetime.utcnow(),
        }},
        upsert=True
    )
    return {"success": True}


@router.get("/{jwt_token}/api/generate/{project_id}/download/docx")
async def download_docx(jwt_token: str, project_id: str):
    """DOCX 파일 다운로드"""
    await get_user_key(jwt_token)
    try:
        file_url = await generate_docx(project_id)
        file_path = os.path.join(".", file_url.lstrip("/"))
        if not os.path.exists(file_path):
            raise HTTPException(status_code=500, detail="파일 생성 실패")

        db = get_db()
        project = await db.projects.find_one({"_id": ObjectId(project_id)})
        filename = f"{project.get('name', 'document')}.docx"

        return FileResponse(
            path=file_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


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


@router.get("/api/shared/{share_token}/info")
async def get_shared_project_info(share_token: str):
    """공유 링크로 프로젝트 타입 조회 (인증 불필요)"""
    db = get_db()
    project = await db.projects.find_one({"share_token": share_token})
    if not project:
        raise HTTPException(status_code=404, detail="공유 링크를 찾을 수 없습니다")
    return {
        "project_type": project.get("project_type", "slide"),
        "project_name": project.get("name", ""),
    }


@router.get("/api/shared/{share_token}/excel")
async def get_shared_excel(share_token: str):
    """공유 링크로 엑셀 데이터 조회 (인증 불필요)"""
    db = get_db()
    project = await db.projects.find_one({"share_token": share_token})
    if not project:
        raise HTTPException(status_code=404, detail="공유 링크를 찾을 수 없습니다")

    project_id = str(project["_id"])
    doc = await db.generated_excel.find_one({"project_id": project_id})
    if not doc:
        raise HTTPException(status_code=404, detail="생성된 엑셀 데이터가 없습니다")
    doc["_id"] = str(doc["_id"])

    return {
        "project_name": project.get("name"),
        "excel": doc,
    }


@router.get("/api/shared/{share_token}/docx")
async def get_shared_docx(share_token: str):
    """공유 링크로 워드 문서 조회 (인증 불필요)"""
    db = get_db()
    project = await db.projects.find_one({"share_token": share_token})
    if not project:
        raise HTTPException(status_code=404, detail="공유 링크를 찾을 수 없습니다")

    project_id = str(project["_id"])
    doc = await db.generated_docx.find_one({"project_id": project_id})
    if not doc:
        raise HTTPException(status_code=404, detail="생성된 문서가 없습니다")
    doc["_id"] = str(doc["_id"])

    return {
        "project_name": project.get("name"),
        "docx": doc,
    }


# ============ 유틸: 슬라이드 히스토리 기록 ============

async def _record_slide_history(
    db, project_id: str, slide_id: str, action: str,
    user_key: str, user_name: str, before_snapshot: dict,
    after_snapshot: dict, description: str
):
    """슬라이드 변경 히스토리 기록 (슬라이드당 최근 10개만 유지)"""
    await db.slide_history.insert_one({
        "project_id": project_id,
        "slide_id": slide_id,
        "action": action,
        "user_key": user_key,
        "user_name": user_name,
        "before_snapshot": before_snapshot,
        "after_snapshot": after_snapshot,
        "description": description,
        "created_at": datetime.utcnow(),
    })

    # 슬라이드당 최근 10개만 유지, 나머지 삭제
    cursor = db.slide_history.find(
        {"slide_id": slide_id},
        {"_id": 1}
    ).sort("created_at", -1).skip(10)
    old_ids = [doc["_id"] async for doc in cursor]
    if old_ids:
        await db.slide_history.delete_many({"_id": {"$in": old_ids}})


async def _normalize_slide_order(db, project_id: str):
    """슬라이드 order를 1..N으로 정규화 (동시 삽입 시 중복/gap 방지)"""
    cursor = db.generated_slides.find(
        {"project_id": project_id}
    ).sort([("order", 1), ("created_at", 1)])
    order = 1
    async for s in cursor:
        if s.get("order") != order:
            await db.generated_slides.update_one(
                {"_id": s["_id"]}, {"$set": {"order": order}}
            )
        order += 1


# ============ 유틸: 생성 취소 체크 ============

async def _check_cancelled(db, project_id: str, generation_id: str) -> bool:
    """생성 중단 요청 여부 확인 (Redis 우선, MongoDB 폴백)"""
    # Redis 우선 체크 (O(1))
    redis_cancelled = await redis_service.check_generation_cancel(project_id)
    if redis_cancelled is True:
        return True

    # Redis에 없거나 불가 시 MongoDB 확인
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
    await redis_service.clear_generation_cancel(project_id)
    await db.projects.update_one(
        {"_id": ObjectId(project_id)},
        {"$set": {"status": "stopped", "updated_at": datetime.utcnow()}}
    )


# ============ 유틸: 공통 헬퍼 ============

async def _analyze_template_slides_cached(template_id: str, slides: list) -> list[dict]:
    """템플릿 슬라이드 분석 결과를 Redis에 캐시 (24시간)"""
    cache_key = f"template:{template_id}:analysis"
    cached = await redis_service.cache_get(cache_key)
    if cached is not None:
        return cached

    result = _analyze_template_slides(slides)
    await redis_service.cache_set(cache_key, result, ttl=86400)
    return result


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
    number_counter = 0  # number role 오브젝트에 순서 번호 직접 부여용
    for obj in template_slide.get("objects", []):
        gen_obj = obj.copy()
        if obj.get("obj_type") == "text":
            role = obj.get("role") or obj.get("_auto_role", "")

            # number role: placeholder 매핑 대신 등장 순서대로 1, 2, 3... 직접 부여
            # (동일 placeholder 이름으로 인한 덮어쓰기 문제 방지)
            if role == "number":
                number_counter += 1
                gen_obj["generated_text"] = str(number_counter)
            else:
                placeholder_name = obj.get("placeholder") or obj.get("_auto_placeholder", "")
                if placeholder_name:
                    generated_text = contents.get(placeholder_name)
                    if generated_text is None:
                        # 내용이 매핑되지 않은 subtitle/description 오브젝트 제거
                        if role in ("subtitle", "description"):
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


def _build_skeleton_objects(template_slide: dict, contents: dict) -> list:
    """스켈레톤 오브젝트 생성 - 텍스트는 빈값, 이미지/위치/스타일은 유지"""
    gen_objects = []
    number_counter = 0
    for obj in template_slide.get("objects", []):
        gen_obj = obj.copy()
        if obj.get("obj_type") == "text":
            role = obj.get("role") or obj.get("_auto_role", "")

            # number role: 스켈레톤에서도 순서 번호 직접 부여
            if role == "number":
                number_counter += 1
                gen_obj["generated_text"] = str(number_counter)
            else:
                placeholder_name = obj.get("placeholder") or obj.get("_auto_placeholder", "")
                if placeholder_name:
                    generated_text = contents.get(placeholder_name)
                    if generated_text is None:
                        if role in ("subtitle", "description"):
                            continue
                    # 스켈레톤: 텍스트를 빈 문자열로 설정
                    gen_obj["generated_text"] = ""

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
