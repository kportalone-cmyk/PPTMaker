from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from services.mongo_service import get_db
from bson import ObjectId
from datetime import datetime

router = APIRouter(tags=["prompts"])


class PromptUpdate(BaseModel):
    content: str


# ── 기본 프롬프트 정의 (DB 초기화용) ──

DEFAULT_PROMPTS = [
    {
        "key": "slide_generation_system",
        "name": "슬라이드 생성 시스템 프롬프트",
        "description": "리소스를 분석하여 슬라이드 구조를 설계하는 시스템 프롬프트. 변수: {lang_instruction}, {slide_count_instruction}",
        "content": """당신은 기업용 프레젠테이션 구조 설계 및 콘텐츠 전문가입니다.
주어진 리소스 자료를 분석하여 전문적인 프레젠테이션을 설계합니다.

## 슬라이드 타입
1. **title** - 타이틀 슬라이드 (프레젠테이션 시작)
   필드: title, subtitle, meta_line
2. **toc** - 목차 슬라이드
   필드: title, items[] (각 항목: num, text)
3. **section** - 섹션 구분 간지
   필수 필드: section_title (제목은 항상 생성해야 합니다)
   선택 필드: section_subtitle (카탈로그에서 해당 간지 템플릿에 "부제목 있음"으로 표시된 경우에만 생성. 없으면 section_subtitle 필드를 생략하세요)
   기타 필드: section_num
4. **content** - 본문 콘텐츠 슬라이드
   필드: title(제목), governance(거버넌스/섹션태그), items[] (heading=부제목 + detail=설명, 순서대로 매핑), sources[]
5. **closing** - 마무리 슬라이드
   필드: title, message, contact

## 콘텐츠 작성 규칙
1. 제목(title)은 간결하고 임팩트 있게 작성합니다 (최대 30자).
2. **governance는 해당 슬라이드의 부제목과 설명 내용을 전체적으로 요약한 문장을 작성합니다 (20~50자).** 단순한 섹션 이름이 아니라, 슬라이드 전체 내용의 핵심을 한 문장으로 압축하세요.
3. **[필수] 본문(content) 슬라이드의 구조**:
   - 본문 슬라이드는 반드시 title(제목), governance(거버넌스), items[](부제목+설명 쌍) 를 생성합니다.
   - 각 item은 heading(부제목, 키워드 1~5단어) + detail(설명, 1~3문장) 구조입니다.
   - heading은 템플릿의 "부제목" 필드에, detail은 "설명" 필드에 순서대로 매핑됩니다.
   - **[필수] 각 content 슬라이드마다 items를 반드시 최소 3개 이상, 최대 4개 생성하세요.** items를 1~2개만 작성하면 안 됩니다.
   - 카탈로그에 표시 가능 개수가 적더라도, 아웃라인 표시를 위해 반드시 3개 이상의 items를 생성해야 합니다.
   - 각 item의 heading은 서로 다른 관점/주제를 다뤄야 합니다.
4. 예시 - subtitle_count=3, description_count=3인 content 슬라이드:
   {{"type":"content","template_index":3,"title":"디지털 전환 핵심 전략","governance":"클라우드 전환, 데이터 분석, 업무 자동화를 통한 디지털 혁신 추진",
     "items":[
       {{"heading":"클라우드 마이그레이션","detail":"기존 온프레미스 인프라를 클라우드로 전환하여 운영 비용을 30% 절감하고 확장성을 확보합니다."}},
       {{"heading":"데이터 기반 의사결정","detail":"빅데이터 분석 플랫폼을 구축하여 실시간 시장 동향 파악과 고객 행동 예측이 가능해집니다."}},
       {{"heading":"업무 자동화 도입","detail":"RPA와 AI를 활용한 반복 업무 자동화로 직원 생산성을 40% 이상 향상시킬 수 있습니다."}}
     ]}}
5. sources가 있으면 출처를 명시합니다.

## 구조 설계 규칙
1. 권장 순서: title → toc → (section → content 슬라이드들)... → closing
2. 각 섹션마다 section(간지) 슬라이드 + 1~3개의 content 슬라이드를 배치합니다.
   **[필수] 목차(toc) 슬라이드의 items 텍스트는 반드시 section 슬라이드의 section_title과 정확히 일치해야 합니다.** 예: section이 3개면 toc items도 3개이며, 각 text가 해당 section_title과 동일합니다.
3. **[필수] 전체 슬라이드가 8장 이하일 경우, 목차(toc)와 섹션 간지(section)는 생략하세요.** 구성: title → content 슬라이드들 → closing. 9장 이상일 때만 toc와 section을 포함합니다.
4. 리소스 내용의 양과 복잡도에 맞게 슬라이드 수를 자유롭게 결정합니다.
5. template_index는 사용 가능한 템플릿 슬라이드 번호입니다 (카탈로그 참조).
6. 같은 template_index를 여러 번 사용할 수 있지만, **같은 타입의 템플릿이 여러 개 있으면 돌아가며 다양하게 사용하세요.** 예를 들어 본문 템플릿 3,4,5번이 모두 content 타입이면 3→4→5→3 순으로 번갈아 사용합니다.
7. 콘텐츠에 맞지 않는 템플릿은 사용하지 않아도 됩니다.
8. **[필수] 카탈로그의 "템플릿 타입 현황"을 반드시 확인하세요. 미등록(✗) 타입은 절대 생성하지 마세요.**
9. {lang_instruction}
10. 반드시 JSON 형식으로만 응답합니다.{slide_count_instruction}""",
    },
    {
        "key": "slide_generation_user",
        "name": "슬라이드 생성 사용자 프롬프트",
        "description": "리소스와 템플릿 카탈로그를 포함한 사용자 프롬프트. 변수: {lang_instruction}, {instructions}, {resources_text}, {slides_description}",
        "content": """아래 리소스 자료를 분석하여 프레젠테이션을 설계하고 콘텐츠를 생성해주세요.

## 출력 언어
{lang_instruction}

## 사용자 지침
{instructions}

## 리소스 자료
{resources_text}

## 사용 가능한 템플릿 슬라이드 카탈로그
{slides_description}

## 응답 형식
반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트는 포함하지 마세요.

```json
{{
    "meta": {{
        "title": "프레젠테이션 제목",
        "subtitle": "부제목",
        "author": "작성자/부서",
        "date": "날짜"
    }},
    "slides": [
        {{
            "type": "title",
            "template_index": 0,
            "title": "프레젠테이션 제목",
            "subtitle": "부제목",
            "meta_line": "작성자 | 날짜"
        }},
        {{
            "type": "toc",
            "template_index": 1,
            "title": "목차",
            "items": [
                {{"num": "01", "text": "섹션 제목"}},
                {{"num": "02", "text": "섹션 제목"}}
            ]
        }},
        {{
            "type": "section",
            "template_index": 2,
            "section_num": "01",
            "section_title": "섹션 제목",
            "section_subtitle": "섹션 부제목"
        }},
        {{
            "type": "content",
            "template_index": 3,
            "title": "슬라이드 제목",
            "governance": "핵심 포인트들의 내용을 종합적으로 요약한 문장 (20~50자)",
            "items": [
                {{"heading": "첫 번째 핵심 포인트", "detail": "첫 번째 포인트에 대한 상세 설명. 구체적인 수치나 사례를 포함하여 1~3문장으로 작성합니다."}},
                {{"heading": "두 번째 핵심 포인트", "detail": "두 번째 포인트에 대한 상세 설명. 논리적 근거와 함께 1~3문장으로 작성합니다."}},
                {{"heading": "세 번째 핵심 포인트", "detail": "세 번째 포인트에 대한 상세 설명. 결론이나 시사점을 1~3문장으로 작성합니다."}}
            ],
            "sources": ["출처1"]
        }},
        {{
            "type": "closing",
            "template_index": 4,
            "title": "감사합니다",
            "message": "마무리 메시지",
            "contact": "팀명 | 이메일"
        }}
    ],
    "sources": [
        {{"ref": "source_id", "title": "출처 제목"}}
    ]
}}
```""",
    },
]


async def ensure_default_prompts():
    """DB에 기본 프롬프트가 없으면 삽입"""
    db = get_db()
    for prompt in DEFAULT_PROMPTS:
        existing = await db.prompts.find_one({"key": prompt["key"]})
        if not existing:
            doc = {**prompt, "created_at": datetime.utcnow(), "updated_at": datetime.utcnow()}
            await db.prompts.insert_one(doc)
            print(f"[Prompt] 기본 프롬프트 등록: {prompt['key']}")


async def get_prompt_content(key: str) -> str:
    """DB에서 프롬프트 content 조회. 없으면 기본값 반환"""
    db = get_db()
    doc = await db.prompts.find_one({"key": key})
    if doc:
        return doc["content"]
    # DB에 없으면 기본 프롬프트에서 찾기
    for p in DEFAULT_PROMPTS:
        if p["key"] == key:
            return p["content"]
    return ""


# ── API 엔드포인트 ──

@router.get("/{jwt_token}/api/admin/prompts")
async def list_prompts(jwt_token: str):
    """프롬프트 목록 조회"""
    db = get_db()
    cursor = db.prompts.find().sort("key", 1)
    prompts = []
    async for p in cursor:
        p["_id"] = str(p["_id"])
        prompts.append(p)
    return {"prompts": prompts}


@router.get("/{jwt_token}/api/admin/prompts/{prompt_id}")
async def get_prompt(jwt_token: str, prompt_id: str):
    """프롬프트 상세 조회"""
    db = get_db()
    doc = await db.prompts.find_one({"_id": ObjectId(prompt_id)})
    if not doc:
        raise HTTPException(status_code=404, detail="프롬프트를 찾을 수 없습니다")
    doc["_id"] = str(doc["_id"])
    return {"prompt": doc}


@router.put("/{jwt_token}/api/admin/prompts/{prompt_id}")
async def update_prompt(jwt_token: str, prompt_id: str, body: PromptUpdate):
    """프롬프트 내용 수정"""
    db = get_db()
    result = await db.prompts.update_one(
        {"_id": ObjectId(prompt_id)},
        {"$set": {"content": body.content, "updated_at": datetime.utcnow()}}
    )
    if result.matched_count == 0:
        raise HTTPException(status_code=404, detail="프롬프트를 찾을 수 없습니다")
    return {"success": True}


@router.post("/{jwt_token}/api/admin/prompts/{prompt_id}/reset")
async def reset_prompt(jwt_token: str, prompt_id: str):
    """프롬프트를 기본값으로 초기화"""
    db = get_db()
    doc = await db.prompts.find_one({"_id": ObjectId(prompt_id)})
    if not doc:
        raise HTTPException(status_code=404, detail="프롬프트를 찾을 수 없습니다")

    # 기본 프롬프트에서 해당 key의 content 찾기
    default_content = ""
    for p in DEFAULT_PROMPTS:
        if p["key"] == doc["key"]:
            default_content = p["content"]
            break

    if not default_content:
        raise HTTPException(status_code=400, detail="기본 프롬프트를 찾을 수 없습니다")

    await db.prompts.update_one(
        {"_id": ObjectId(prompt_id)},
        {"$set": {"content": default_content, "updated_at": datetime.utcnow()}}
    )
    return {"success": True, "content": default_content}
