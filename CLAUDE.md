# PPTMaker - 기업용 파워포인트 자동 생성 솔루션

## 프로젝트 개요
기업용 파워포인트 자동 생성 솔루션. 관리자가 슬라이드 템플릿을 구축하고, 사용자가 리소스를 등록하여 AI 기반 파워포인트를 자동 생성한다.

## 아키텍처 스택

### 백엔드
- **서버**: Python FastAPI (비동기)
- **DB**: MongoDB v4
  - 접속: `MONGO_URI=mongodb://imadmin:kmslabkm@im.k-portal.co.kr:16270/?authSource=admin`
  - PPTMaker DB: `PPTMaker`
  - 조직도 DB: `im_org_info` / Collection: `user_info`
  - 조직도 필드: 사용자명(nm), 부서명(dp), 메일(em), 키(ky), 역할(role), M365계정(m365)
- **외부 API**: Perplexity Search API (`.env`에 API key)

### 프론트엔드
- HTML, CSS, JS, jQuery
- Tailwind CSS, Chart.js
- **관리자 모듈**: `/admin/` 폴더
- **사용자 모듈**: `/front/` 폴더

## 프로젝트 구조
```
d:/PPTMaker/
├── CLAUDE.md
├── .env                    # 환경설정 (하드코딩 금지)
├── server/                 # FastAPI 서버
│   ├── venv/               # Python 가상환경
│   ├── requirements.txt
│   ├── main.py             # FastAPI 엔트리포인트
│   ├── config.py           # .env 로드 및 설정
│   ├── routers/
│   │   ├── auth.py         # JWT 인증
│   │   ├── admin.py        # 관리자 API
│   │   ├── template.py     # 템플릿 관리 API
│   │   ├── project.py      # 프로젝트 관리 API
│   │   ├── resource.py     # 리소스 관리 API
│   │   ├── generate.py     # PPT 생성 API
│   │   └── font.py         # 폰트 관리 API
│   ├── services/
│   │   ├── mongo_service.py    # MongoDB 연결/관리
│   │   ├── auth_service.py     # JWT/인증 서비스
│   │   ├── template_service.py # 템플릿 비즈니스 로직
│   │   ├── ppt_service.py      # PPT 생성 서비스
│   │   ├── search_service.py   # Perplexity 검색
│   │   └── file_service.py     # 파일 업로드/처리
│   ├── models/
│   │   ├── template.py     # 템플릿 데이터 모델
│   │   ├── project.py      # 프로젝트 데이터 모델
│   │   ├── resource.py     # 리소스 데이터 모델
│   │   └── user.py         # 사용자 데이터 모델
│   └── utils/
│       ├── versioning.py   # 정적파일 버전 관리
│       └── crypto.py       # 암호화 유틸
├── admin/                  # 관리자 프론트엔드
│   ├── index.html          # 관리자 메인
│   ├── css/
│   │   └── admin.css
│   └── js/
│       └── admin.js
├── front/                  # 사용자 프론트엔드
│   ├── index.html          # 사용자 메인
│   ├── css/
│   │   └── app.css
│   └── js/
│       └── app.js
└── uploads/                # 업로드 파일 저장
    ├── backgrounds/        # 배경 이미지
    ├── images/             # 슬라이드 이미지
    └── resources/          # 리소스 파일
```

## 핵심 개발 규칙

### 필수 규칙
1. **환경설정**: 모든 설정은 `.env` 파일 사용. 하드코딩 금지
2. **상태 관리**: 서버 메모리/브라우저 localStorage 사용 금지. 모든 상태는 MongoDB에 저장
3. **L4 이중화 대응**: 다중 서버 환경 기준 개발 (세션/상태 공유 문제 고려)
4. **정적 파일 버전**: .js, .css 파일에 자동 버전 파라미터 추가 (`?v=timestamp`)
5. **보안**: ID/패스워드 암호화 처리, JWT 인증
6. **MongoDB 인덱스**: 데이터 추가/검색 시 필요한 인덱스 자동 등록
7. **requirements.txt**: 패키지 추가 시 자동 업데이트

### JWT 인증 체계
- URL에 JWT 포함: `https://xxx.xxx.com/{jwt}/api/...`
- 모든 API 호출 시 JWT에서 사용자 정보 확인
- 관리자 판별: 조직도 user_info의 `role` 필드가 `"admin"`인 경우

### 프론트엔드 규칙
- HTML, CSS, JS, jQuery만 사용 (프레임워크 없음)
- Tailwind CSS로 스타일링
- Chart.js 차트 라이브러리
- 사용자 검색: 동명이인 선택창, 키보드 ↑↓ + Enter 선택

## 주요 기능

### 관리자 모듈
- 템플릿 관리 (목록, 생성, 편집, 삭제)
- 공통 배경이미지 업로드
- 슬라이드 캔버스 편집 (이미지/텍스트 오브젝트 드래그&드롭)
- 텍스트 오브젝트: 폰트/색상/사이즈 지정 (서버 등록 폰트)
- 슬라이드 메타정보 관리 (제목/거버넌스/서브텍스트 수 등 → 추천 활용)
- 레이아웃: 좌측 슬라이드 목록 + 우측 편집 캔버스

### 사용자 모듈
- 프로젝트 CRUD
- 리소스 등록 (파일: Word/Excel/PPT/PDF/Text, 텍스트 붙여넣기)
- 웹 검색 리소스 추가 (Perplexity API)
- 사용자 지침 입력 + 템플릿 선택
- PPT 자동 생성 (타이핑 효과로 실시간 표시)
- 프레젠테이션 미리보기
- PDF/PPTX 다운로드
- 미리보기 URL 공유 링크

## MongoDB 컬렉션 구조 (PPTMaker DB)
- `templates`: 템플릿 관리
- `slides`: 슬라이드 데이터 (오브젝트 위치/속성)
- `projects`: 사용자 프로젝트
- `resources`: 프로젝트 리소스
- `generated_slides`: 생성된 PPT 슬라이드 데이터
- `fonts`: 등록된 폰트 정보

## 빌드 & 실행
```bash
cd server
python -m venv venv
source venv/Scripts/activate  # Windows
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```
