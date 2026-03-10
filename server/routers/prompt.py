import hashlib
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from typing import Optional
from services.mongo_service import get_db
from services import redis_service
from bson import ObjectId
from datetime import datetime

router = APIRouter(tags=["prompts"])

# 솔루션에서 제공하는 LLM 모델 목록
AVAILABLE_MODELS = [
    {"id": "claude-opus-4-6", "name": "Claude Opus 4.6", "description": "최고 성능 모델 (정밀도 높은 작업)"},
    {"id": "claude-sonnet-4-6", "name": "Claude Sonnet 4.6", "description": "균형 잡힌 모델 (속도 + 성능)"},
    {"id": "claude-haiku-4-5-20251001", "name": "Claude Haiku 4.5", "description": "빠른 응답 모델 (간단한 작업)"},
]


class PromptUpdate(BaseModel):
    content: Optional[str] = None
    model: Optional[str] = None


# ── 기본 프롬프트 정의 (DB 초기화용) ──

DEFAULT_PROMPTS = [
    {
        "key": "slide_generation_system",
        "name": "슬라이드 생성 시스템 프롬프트",
        "description": "리소스를 분석하여 슬라이드 구조를 설계하는 시스템 프롬프트. 변수: {lang_instruction}, {slide_count_instruction}",
        "model": "claude-sonnet-4-6",
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
   필수 필드: section_num (첫 번째 간지는 "01", 두 번째는 "02", 세 번째는 "03" 등 순서대로 자동 증가하는 두 자리 번호)
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
   - **[필수] 각 content 슬라이드의 items 개수는 반드시 선택한 템플릿의 표시 가능 개수와 정확히 일치해야 합니다.** 카탈로그에서 "N개 항목을 표시" 라고 되어 있으면 items를 정확히 N개 생성하세요.
   - **[필수] 다양한 템플릿을 골고루 활용하세요.** 1개짜리, 2개짜리, 3개짜리, 4개짜리 템플릿이 있으면 내용의 성격에 따라 적합한 템플릿을 선택하여 다양한 레이아웃으로 구성하세요.
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
        "model": "claude-sonnet-4-6",
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
    {
        "key": "excel_generation_system",
        "name": "엑셀 생성 시스템 프롬프트",
        "description": "리소스를 분석하여 구조화된 스프레드시트 데이터를 생성하는 시스템 프롬프트. 변수: {lang_instruction}, {sheet_count_instruction}",
        "model": "claude-opus-4-6",
        "content": """당신은 데이터 분석 및 구조화 전문가입니다.
주어진 리소스 자료를 분석하여 구조화된 스프레드시트 데이터를 생성합니다.

## 출력 규칙
1. 데이터를 논리적 시트(탭)로 분류합니다. 관련 데이터끼리 같은 시트에 배치하세요.
2. 각 시트는 명확하고 구체적인 열 헤더를 포함합니다.
3. 데이터는 정확한 셀 값으로 정리합니다:
   - 숫자 데이터는 숫자 타입으로 (문자열 X)
   - 날짜는 "YYYY-MM-DD" 형식
   - 빈 셀은 빈 문자열("")
4. 각 시트는 최소 2개 열, 최소 3개 행의 데이터를 포함해야 합니다.
5. 열 헤더는 내용을 명확히 설명하는 이름으로 작성합니다.
6. 데이터는 논리적 순서로 정렬합니다 (날짜순, 가나다순, 크기순 등).
7. {lang_instruction}
8. 반드시 JSON 형식으로만 응답합니다.{sheet_count_instruction}

## 차트 생성 규칙
데이터가 시각화에 적합한 경우, 해당 시트에 charts 배열을 추가하세요.
1. 숫자 데이터가 2개 이상의 행과 비교 가능한 열이 있으면 차트를 생성합니다.
2. 지원 차트 타입: bar(막대), line(선), pie(원형), area(영역), scatter(산점도), doughnut(도넛), radar(레이더)
3. 데이터 특성에 맞는 차트 타입을 선택하세요:
   - 항목 간 비교: bar
   - 시계열/추세 변화: line 또는 area
   - 비율/구성: pie 또는 doughnut
   - 상관관계: scatter
   - 다차원 비교: radar
4. 한 시트에 최대 3개까지 차트를 생성할 수 있습니다.
5. pie/doughnut 차트는 series에 하나의 열만 포함하세요.
6. 차트가 적합하지 않은 데이터(텍스트만 있는 시트 등)에는 charts를 생략하세요.
7. labels_column은 카테고리/라벨 역할을 하는 열의 인덱스(0부터)입니다.
8. series의 column은 숫자 데이터가 있는 열의 인덱스(0부터)입니다.""",
    },
    {
        "key": "excel_generation_user",
        "name": "엑셀 생성 사용자 프롬프트",
        "description": "리소스를 포함한 엑셀 생성 사용자 프롬프트. 변수: {lang_instruction}, {instructions}, {resources_text}",
        "model": "claude-opus-4-6",
        "content": """아래 리소스 자료를 분석하여 구조화된 스프레드시트 데이터를 생성해주세요.

## 출력 언어
{lang_instruction}

## 사용자 지침
{instructions}

## 리소스 자료
{resources_text}

## 응답 형식
반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트는 포함하지 마세요.

```json
{{
    "meta": {{
        "title": "스프레드시트 제목",
        "description": "데이터 설명"
    }},
    "sheets": [
        {{
            "name": "시트명 (최대 31자)",
            "columns": ["열1", "열2", "열3"],
            "rows": [
                ["값1", 100, "2024-01-01"],
                ["값2", 200, "2024-02-01"]
            ],
            "charts": [
                {{
                    "type": "bar",
                    "title": "차트 제목",
                    "data_range": {{
                        "labels_column": 0,
                        "series": [
                            {{"name": "시리즈명", "column": 1}},
                            {{"name": "시리즈명2", "column": 2}}
                        ],
                        "row_start": 0,
                        "row_end": null
                    }},
                    "options": {{
                        "stacked": false,
                        "show_legend": true
                    }}
                }}
            ]
        }}
    ]
}}
```
charts는 선택적입니다. 숫자 데이터가 시각화에 적합한 경우에만 포함하세요. 텍스트만 있는 시트에는 생략합니다.""",
    },
    {
        "key": "excel_modify_system",
        "name": "엑셀 수정 시스템 프롬프트",
        "description": "기존 엑셀 데이터를 사용자 지침에 따라 부분 수정하는 시스템 프롬프트. 변수: {lang_instruction}",
        "model": "claude-sonnet-4-6",
        "content": """당신은 데이터 분석 및 스프레드시트 편집 전문가입니다.
사용자의 지침에 따라 기존 스프레드시트 데이터를 수정합니다.

## 중요 규칙
1. 사용자가 특정 부분만 수정 요청하면 해당 부분만 변경하고 나머지는 그대로 유지하세요.
2. 숫자 데이터는 숫자 타입으로, 날짜는 "YYYY-MM-DD" 형식으로 유지합니다.
3. {lang_instruction}
4. 반드시 JSON 형식으로만 응답합니다.

## 수정 유형별 처리
- 열 추가/삭제: 기존 데이터 유지하면서 열 추가 또는 제거
- 행 추가/삭제: 해당 행만 추가 또는 제거, 나머지 유지
- 셀 값 변경: 지정된 셀만 수정
- 차트 변경: 차트 타입/설정만 수정, 데이터 유지
- 시트 추가/삭제: 기존 시트 유지하면서 추가 또는 제거
- 데이터 재계산: 수식/계산 요청 시 값을 직접 계산하여 결과값으로 반환

## 차트 규칙
- 데이터 변경 시 차트의 data_range(row_start, row_end, series)도 함께 업데이트
- 지원 타입: bar, line, pie, area, scatter, doughnut, radar
- pie/doughnut 차트는 series 1개만""",
    },
    {
        "key": "excel_modify_user",
        "name": "엑셀 수정 사용자 프롬프트",
        "description": "기존 엑셀 데이터와 수정 지침을 포함한 사용자 프롬프트. 변수: {lang_instruction}, {instruction}, {current_excel_data}",
        "model": "claude-sonnet-4-6",
        "content": """아래 기존 스프레드시트 데이터를 사용자의 수정 지침에 따라 수정해주세요.

## 출력 언어
{lang_instruction}

## 현재 스프레드시트 데이터
{current_excel_data}

## 사용자 수정 지침
{instruction}

## 응답 형식
수정된 데이터를 아래 JSON 형식으로 응답하세요.
- "target_sheet" 필드가 있으면: 해당 시트만 수정하여 sheets 배열에 1개만 포함하세요.
- "target_sheet" 필드가 없으면: 기존의 모든 시트를 반드시 포함하세요 (수정하지 않은 시트도 원본 그대로 포함).
  - 시트 추가 요청 시: 기존 시트 전부 + 새 시트를 sheets 배열에 포함하세요. 기존 시트를 절대 제거하지 마세요.
  - 시트 삭제 요청 시: 삭제 대상만 제거하고 나머지 시트는 모두 유지하세요.

```json
{{
    "meta": {{
        "title": "스프레드시트 제목",
        "description": "데이터 설명"
    }},
    "sheets": [
        {{
            "name": "시트명 (최대 31자)",
            "columns": ["열1", "열2", "열3"],
            "rows": [
                ["값1", 100, "2024-01-01"],
                ["값2", 200, "2024-02-01"]
            ],
            "charts": [
                {{
                    "type": "bar",
                    "title": "차트 제목",
                    "data_range": {{
                        "labels_column": 0,
                        "series": [
                            {{"name": "시리즈명", "column": 1}}
                        ],
                        "row_start": 0,
                        "row_end": null
                    }},
                    "options": {{
                        "stacked": false,
                        "show_legend": true
                    }}
                }}
            ]
        }}
    ]
}}
```
charts는 선택적입니다. 기존 차트를 유지/수정/삭제 요청에 따라 처리하세요.""",
    },
    {
        "key": "docx_generation_system",
        "name": "문서 생성 시스템 프롬프트",
        "description": "리소스를 분석하여 구조화된 Word 문서를 생성하는 시스템 프롬프트. 변수: {lang_instruction}, {section_count_instruction}",
        "model": "claude-opus-4-6",
        "content": """당신은 전문 비즈니스 문서 작성 및 구조화 전문가입니다.
주어진 리소스 자료를 분석하여 시각적으로 풍부하고 전문적인 문서를 작성합니다.

## 출력 규칙
1. 문서를 논리적 섹션으로 구성합니다.
2. 각 섹션은 제목(title)과 본문(content)을 포함합니다.
3. 본문은 Markdown 형식으로 작성하되, 시각적 요소를 적극 활용합니다:
   - **굵게**, *기울임* 으로 핵심 내용 강조
   - 불릿 리스트(- 또는 *)와 번호 리스트(1. 2. 3.)로 정보 구조화
   - 데이터 비교, 현황, 분석 시 반드시 Markdown 표(| col1 | col2 |) 사용
   - 핵심 포인트, 요약, 주의사항은 인용(> 텍스트)으로 강조
4. level은 문서 구조의 깊이입니다 (1=대제목, 2=소제목, 3=하위제목).
5. 전문적이고 읽기 쉬운 문서를 작성합니다.
6. 충분한 분량의 본문 내용을 작성합니다. 각 섹션의 content는 최소 3문장 이상이어야 합니다.
7. 수치, 통계, 비교 데이터가 있으면 반드시 표로 정리하세요.
8. 각 주요 섹션 끝에 핵심 요약을 인용(>) 형식으로 추가하세요.
9. {lang_instruction}
10. 반드시 JSON 형식으로만 응답합니다.{section_count_instruction}

## 차트 생성 규칙
수치 데이터가 포함된 경우 **사용자가 별도로 요청하지 않더라도** 차트를 자동으로 추가하세요.
아래 조건 중 하나라도 해당하면 Markdown 표와 함께 ```chart 코드 블록을 반드시 삽입합니다:
- 3개 이상의 수치 비교 데이터 (매출, 인원, 비용, 점수 등)
- 시계열/기간별 추이 데이터 (월별, 분기별, 연도별 등)
- 비율/구성비 데이터 (시장 점유율, 예산 배분 등)
- 순위/랭킹 데이터
사용자가 차트/그래프를 명시적으로 요청하면 더욱 적극적으로 다양한 차트를 추가하세요.

차트 JSON 형식 (content 내부에 ```chart ... ``` 블록으로 삽입):
```chart
{{
  "type": "bar",
  "data": {{
    "labels": ["항목1", "항목2", "항목3"],
    "datasets": [
      {{
        "label": "시리즈명",
        "data": [10, 20, 30],
        "backgroundColor": ["#6366F1", "#EC4899", "#F59E0B"]
      }}
    ]
  }},
  "options": {{
    "title": "차트 제목",
    "scales": {{ "y": {{ "title": "단위" }} }}
  }}
}}
```

지원 차트 타입: bar(막대), line(선), pie(원형), doughnut(도넛), area(영역)
- 비교/순위 데이터 → bar
- 시계열/추이 데이터 → line
- 비율/구성 데이터 → pie 또는 doughnut
- 누적 추이 → area""",
    },
    {
        "key": "docx_generation_user",
        "name": "문서 생성 사용자 프롬프트",
        "description": "리소스를 포함한 문서 생성 사용자 프롬프트. 변수: {lang_instruction}, {instructions}, {resources_text}",
        "model": "claude-opus-4-6",
        "content": """아래 리소스 자료를 분석하여 전문 문서를 작성해주세요.

## 출력 언어
{lang_instruction}

## 사용자 지침
{instructions}

## 리소스 자료
{resources_text}

## 응답 형식
반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트는 포함하지 마세요.

```json
{{
    "meta": {{
        "title": "문서 제목",
        "description": "문서 설명"
    }},
    "sections": [
        {{
            "title": "섹션 제목",
            "level": 1,
            "content": "Markdown 형식의 본문 내용...\\n\\n여러 문단을 포함할 수 있습니다."
        }},
        {{
            "title": "하위 섹션",
            "level": 2,
            "content": "상세 내용...\\n\\n- 불릿 항목 1\\n- 불릿 항목 2"
        }}
    ]
}}
```""",
    },
    {
        "key": "rewrite_system",
        "name": "텍스트 리라이트 시스템 프롬프트",
        "description": "선택한 텍스트를 사용자 지시에 따라 수정하는 시스템 프롬프트. 변수: {lang_instruction}",
        "model": "",
        "content": """당신은 전문 비즈니스 문서 편집 전문가입니다.
사용자가 문서에서 선택한 텍스트를 지시사항에 따라 수정합니다.

## 규칙
1. 선택된 텍스트만 수정합니다. 문서의 나머지 부분은 변경하지 않습니다.
2. 사용자의 지시사항을 정확히 반영합니다.
3. 원문의 전체적인 톤과 스타일을 유지하면서 수정합니다.
4. Markdown 형식을 유지합니다 (굵게, 기울임, 리스트, 표 등).
5. {lang_instruction}
6. 수정된 텍스트만 출력합니다. 설명이나 부가 텍스트를 추가하지 마세요.
7. 절대 ```markdown 등의 코드 블록으로 감싸지 마세요. 순수 텍스트/마크다운만 출력합니다.""",
    },
    {
        "key": "rewrite_user",
        "name": "텍스트 리라이트 사용자 프롬프트",
        "description": "선택 텍스트와 수정 지시를 포함하는 사용자 프롬프트. 변수: {lang_instruction}, {instructions}, {selected_text}, {context_text}",
        "model": "",
        "content": """## 출력 언어
{lang_instruction}

## 수정 지시사항
{instructions}

## 선택된 텍스트 (이 부분만 수정하세요)
{selected_text}

## 주변 문맥 (참고용, 수정하지 마세요)
{context_text}

위의 "선택된 텍스트"를 수정 지시사항에 따라 수정하여 출력하세요. 수정된 텍스트만 출력하세요.""",
    },
]


def _content_hash(content: str) -> str:
    """프롬프트 content의 해시값 계산"""
    return hashlib.md5(content.strip().encode()).hexdigest()


async def ensure_default_prompts():
    """DB에 기본 프롬프트가 없으면 삽입, content가 변경되었으면 업데이트"""
    db = get_db()
    for prompt in DEFAULT_PROMPTS:
        new_hash = _content_hash(prompt["content"])
        existing = await db.prompts.find_one({"key": prompt["key"]})
        if not existing:
            doc = {
                **prompt,
                "content_hash": new_hash,
                "created_at": datetime.utcnow(),
                "updated_at": datetime.utcnow(),
            }
            await db.prompts.insert_one(doc)
            print(f"[Prompt] 기본 프롬프트 등록: {prompt['key']}")
        else:
            update_fields = {}
            # content 변경 감지 (관리자가 수동 수정하지 않은 경우에만 업데이트)
            old_hash = existing.get("content_hash", "")
            existing_content_hash = _content_hash(existing.get("content", ""))
            # 관리자가 수동 수정한 경우: old_hash != existing_content_hash
            # 기본값 그대로인 경우: old_hash == existing_content_hash (또는 old_hash 없음)
            admin_modified = old_hash and old_hash != existing_content_hash
            if not admin_modified and old_hash != new_hash:
                update_fields["content"] = prompt["content"]
                update_fields["content_hash"] = new_hash
                print(f"[Prompt] content 업데이트: {prompt['key']}")
            # model 필드 동기화
            if "model" not in existing or existing.get("model") != prompt.get("model", ""):
                update_fields["model"] = prompt.get("model", "")
            if update_fields:
                update_fields["updated_at"] = datetime.utcnow()
                await db.prompts.update_one(
                    {"_id": existing["_id"]},
                    {"$set": update_fields},
                )
                # Redis 캐시 무효화
                await redis_service.cache_delete(f"prompt:{prompt['key']}")


async def get_prompt_content(key: str) -> str:
    """DB에서 프롬프트 content 조회 (Redis 캐시, 24시간). 없으면 기본값 반환"""
    cache_key = f"prompt:{key}"
    cached = await redis_service.cache_get(cache_key)
    if cached is not None:
        return cached

    db = get_db()
    doc = await db.prompts.find_one({"key": key})
    if doc:
        await redis_service.cache_set(cache_key, doc["content"], ttl=86400)
        return doc["content"]
    # DB에 없으면 기본 프롬프트에서 찾기
    for p in DEFAULT_PROMPTS:
        if p["key"] == key:
            return p["content"]
    return ""


async def get_prompt_model(key: str) -> str:
    """DB에서 프롬프트에 설정된 모델 조회. 없으면 기본값 반환"""
    cache_key = f"prompt_model:{key}"
    cached = await redis_service.cache_get(cache_key)
    if cached is not None:
        return cached

    db = get_db()
    doc = await db.prompts.find_one({"key": key}, {"model": 1})
    if doc and doc.get("model"):
        await redis_service.cache_set(cache_key, doc["model"], ttl=86400)
        return doc["model"]
    # DB에 없으면 기본 프롬프트에서 찾기
    for p in DEFAULT_PROMPTS:
        if p["key"] == key:
            return p.get("model", "")
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


@router.get("/{jwt_token}/api/admin/prompts/models")
async def list_available_models(jwt_token: str):
    """사용 가능한 LLM 모델 목록 조회"""
    return {"models": AVAILABLE_MODELS}


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
    """프롬프트 내용/모델 수정"""
    db = get_db()
    # 캐시 무효화를 위해 key 조회
    doc = await db.prompts.find_one({"_id": ObjectId(prompt_id)}, {"key": 1})
    update_fields = {"updated_at": datetime.utcnow()}
    if body.content is not None:
        update_fields["content"] = body.content
    if body.model is not None:
        update_fields["model"] = body.model
    result = await db.prompts.update_one(
        {"_id": ObjectId(prompt_id)},
        {"$set": update_fields}
    )
    if result.matched_count == 0:
        raise HTTPException(status_code=404, detail="프롬프트를 찾을 수 없습니다")
    if doc:
        await redis_service.cache_delete(f"prompt:{doc['key']}")
        await redis_service.cache_delete(f"prompt_model:{doc['key']}")
    return {"success": True}


@router.post("/{jwt_token}/api/admin/prompts/{prompt_id}/reset")
async def reset_prompt(jwt_token: str, prompt_id: str):
    """프롬프트를 기본값으로 초기화"""
    db = get_db()
    doc = await db.prompts.find_one({"_id": ObjectId(prompt_id)})
    if not doc:
        raise HTTPException(status_code=404, detail="프롬프트를 찾을 수 없습니다")

    # 기본 프롬프트에서 해당 key의 content와 model 찾기
    default_content = ""
    default_model = ""
    for p in DEFAULT_PROMPTS:
        if p["key"] == doc["key"]:
            default_content = p["content"]
            default_model = p.get("model", "")
            break

    if not default_content:
        raise HTTPException(status_code=400, detail="기본 프롬프트를 찾을 수 없습니다")

    await db.prompts.update_one(
        {"_id": ObjectId(prompt_id)},
        {"$set": {"content": default_content, "model": default_model, "updated_at": datetime.utcnow()}}
    )
    await redis_service.cache_delete(f"prompt:{doc['key']}")
    await redis_service.cache_delete(f"prompt_model:{doc['key']}")
    return {"success": True, "content": default_content, "model": default_model}
