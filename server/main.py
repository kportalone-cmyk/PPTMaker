import os
import sys
import time
import hashlib
import logging
from pathlib import Path
from contextlib import asynccontextmanager

# Windows asyncio ConnectionResetError 로그 억제
logging.getLogger("asyncio").setLevel(logging.CRITICAL)

from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.middleware.cors import CORSMiddleware

from config import settings
from services.mongo_service import init_indexes, seed_demo_accounts, seed_slide_styles, close_connection, close_connection_sync
from services.redis_service import init_redis, close_redis, close_redis_sync
from routers import auth, template, project, resource, generate, font, prompt, collaboration, onlyoffice
from utils.versioning import get_file_version


class StaticCacheMiddleware:
    """업로드 이미지에 브라우저 캐시 헤더 추가 (순수 ASGI 미들웨어 - shutdown hang 방지)"""

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return

        path = scope.get("path", "")
        if not path.startswith("/uploads/"):
            await self.app(scope, receive, send)
            return

        async def send_with_cache(message):
            if message["type"] == "http.response.start":
                headers = list(message.get("headers", []))
                headers.append((b"cache-control", b"public, max-age=604800, immutable"))
                message["headers"] = headers
            await send(message)

        await self.app(scope, receive, send_with_cache)


@asynccontextmanager
async def lifespan(app: FastAPI):
    """서버 시작/종료 이벤트"""
    import asyncio

    # 시작 시 인덱스 초기화
    await init_indexes()
    await seed_demo_accounts()
    await seed_slide_styles()
    print("MongoDB 인덱스 초기화 완료")

    # Redis 초기화
    redis_ok = await init_redis()
    if redis_ok:
        print("Redis 연결 성공")
    else:
        print("Redis 연결 실패 - MongoDB 폴백 모드로 동작합니다")

    # 기본 프롬프트 초기화
    from routers.prompt import ensure_default_prompts
    await ensure_default_prompts()
    print("기본 프롬프트 초기화 완료")

    # 업로드 디렉토리 생성
    for sub in ["backgrounds", "images", "resources", "generated", "documents", "custom_templates"]:
        os.makedirs(os.path.join(settings.UPLOAD_DIR, sub), exist_ok=True)

    try:
        yield
    except asyncio.CancelledError:
        pass
    finally:
        # 종료 시 연결 해제 (CancelledError 포함 모든 상황에서 정리)
        try:
            await close_redis()
            print("Redis 연결 해제")
        except Exception:
            pass
        try:
            close_connection_sync()
            print("MongoDB 연결 해제")
        except Exception:
            pass


app = FastAPI(
    title=f"{settings.SOLUTION_NAME} API",
    description="기업용 파워포인트 자동 생성 솔루션",
    version="1.0.0",
    lifespan=lifespan,
)

# 업로드 파일 캐시 헤더 (순수 ASGI 미들웨어)
app.add_middleware(StaticCacheMiddleware)

# CORS 설정 (.env CORS_ORIGINS 기반)
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 라우터 등록
app.include_router(auth.router)
app.include_router(template.router)
app.include_router(project.router)
app.include_router(resource.router)
app.include_router(generate.router)
app.include_router(font.router)
app.include_router(prompt.router)
app.include_router(collaboration.router)
app.include_router(onlyoffice.router)

# 정적 파일 서빙
project_root = Path(__file__).resolve().parent.parent

# 업로드 파일
uploads_path = Path(settings.UPLOAD_DIR)
if not uploads_path.is_absolute():
    uploads_path = Path(__file__).resolve().parent / settings.UPLOAD_DIR
uploads_path.mkdir(parents=True, exist_ok=True)
app.mount("/uploads", StaticFiles(directory=str(uploads_path)), name="uploads")

# 관리자 정적 파일
admin_path = project_root / "admin"
if admin_path.exists():
    app.mount("/admin/css", StaticFiles(directory=str(admin_path / "css")), name="admin_css")
    app.mount("/admin/js", StaticFiles(directory=str(admin_path / "js")), name="admin_js")

# 사용자 정적 파일
front_path = project_root / "front"
if front_path.exists():
    app.mount("/front/css", StaticFiles(directory=str(front_path / "css")), name="front_css")
    app.mount("/front/js", StaticFiles(directory=str(front_path / "js")), name="front_js")
    if (front_path / "img").exists():
        app.mount("/front/img", StaticFiles(directory=str(front_path / "img")), name="front_img")
    if (front_path / "static").exists():
        app.mount("/front/static", StaticFiles(directory=str(front_path / "static")), name="front_static")


# ============ HTML 페이지 라우트 ============

def inject_version(html_content: str, base_dir: str) -> str:
    """HTML 내 CSS/JS 파일 참조에 버전 파라미터 추가 + 솔루션명 주입"""
    import re

    def add_version(match):
        tag = match.group(0)
        url = match.group(1)
        if "?" in url:
            return tag
        file_path = os.path.join(base_dir, url.lstrip("/"))
        version = get_file_version(file_path)
        new_url = f"{url}?v={version}"
        return tag.replace(url, new_url)

    # CSS href
    html_content = re.sub(
        r'href="([^"]+\.css)"',
        add_version,
        html_content
    )
    # JS src
    html_content = re.sub(
        r'src="([^"]+\.js)"',
        add_version,
        html_content
    )

    # 솔루션명 주입 (window.__SOLUTION_NAME__ 전역 변수)
    solution_script = f'<script>window.__SOLUTION_NAME__="{settings.SOLUTION_NAME}";</script>'
    html_content = html_content.replace("</head>", f"{solution_script}\n</head>", 1)

    return html_content


@app.get("/main", response_class=HTMLResponse)
async def serve_main_landing():
    """제품 홍보 랜딩 페이지"""
    html_path = project_root / "front" / "static" / "main.html"
    if not html_path.exists():
        return HTMLResponse("<h1>Page not found</h1>", status_code=404)
    content = html_path.read_text(encoding="utf-8")
    # 솔루션명 주입
    solution_script = f'<script>window.__SOLUTION_NAME__="{settings.SOLUTION_NAME}";</script>'
    content = content.replace("</head>", f"{solution_script}\n</head>", 1)
    return HTMLResponse(content)


@app.get("/admin", response_class=HTMLResponse)
@app.get("/admin/{path:path}", response_class=HTMLResponse)
async def serve_admin(path: str = ""):
    """관리자 페이지 서빙"""
    html_path = project_root / "admin" / "index.html"
    if not html_path.exists():
        return HTMLResponse("<h1>Admin page not found</h1>", status_code=404)
    content = html_path.read_text(encoding="utf-8")
    content = inject_version(content, str(project_root))
    return HTMLResponse(content)


@app.get("/")
async def serve_root():
    """루트 접속 시 홍보 페이지로 리다이렉트"""
    return RedirectResponse(url="/main", status_code=302)


@app.get("/app", response_class=HTMLResponse)
@app.get("/app/{path:path}", response_class=HTMLResponse)
async def serve_front(path: str = ""):
    """사용자 페이지 서빙"""
    html_path = project_root / "front" / "index.html"
    if not html_path.exists():
        return HTMLResponse("<h1>Front page not found</h1>", status_code=404)
    content = html_path.read_text(encoding="utf-8")
    content = inject_version(content, str(project_root))
    return HTMLResponse(content)


@app.get("/shared/{share_token}", response_class=HTMLResponse)
async def serve_shared(share_token: str):
    """공유 프레젠테이션 페이지 서빙"""
    html_path = project_root / "front" / "index.html"
    if not html_path.exists():
        return HTMLResponse("<h1>Page not found</h1>", status_code=404)
    content = html_path.read_text(encoding="utf-8")
    content = inject_version(content, str(project_root))
    return HTMLResponse(content)


@app.get("/{jwt_token}/{lang}", response_class=HTMLResponse)
async def serve_front_with_lang(jwt_token: str, lang: str):
    """JWT + 언어 설정 URL: /{jwt_token}/{lang}
    예: /eyJ0eXA.../ko
    """
    # API 경로와 충돌 방지
    if jwt_token in ("api", "admin", "front", "app", "shared", "uploads", "main"):
        return HTMLResponse("<h1>Not Found</h1>", status_code=404)
    html_path = project_root / "front" / "index.html"
    if not html_path.exists():
        return HTMLResponse("<h1>Front page not found</h1>", status_code=404)
    content = html_path.read_text(encoding="utf-8")
    content = inject_version(content, str(project_root))
    return HTMLResponse(content)


@app.get("/{jwt_token}", response_class=HTMLResponse)
async def serve_front_with_jwt(jwt_token: str):
    """JWT만 포함된 URL: /{jwt_token}
    예: /eyJ0eXA... (언어 미지정 → 프론트엔드 기본 언어 사용)
    """
    # 예약어/정적파일 경로 충돌 방지
    if jwt_token in ("api", "admin", "front", "app", "shared", "uploads", "main", "favicon.ico"):
        return HTMLResponse("<h1>Not Found</h1>", status_code=404)
    html_path = project_root / "front" / "index.html"
    if not html_path.exists():
        return HTMLResponse("<h1>Front page not found</h1>", status_code=404)
    content = html_path.read_text(encoding="utf-8")
    content = inject_version(content, str(project_root))
    return HTMLResponse(content)


# 지원 언어 정보 API
@app.get("/api/config/langs")
async def get_supported_langs():
    """지원 언어 목록 반환"""
    lang_labels = {
        "ko": "한국어",
        "en": "English",
        "ja": "日本語",
        "zh": "中文",
    }
    return {
        "langs": [
            {"code": lang, "label": lang_labels.get(lang, lang)}
            for lang in settings.SUPPORTED_LANGS
        ],
        "default": settings.SUPPORTED_LANGS[0] if settings.SUPPORTED_LANGS else "ko",
    }


# 사용자 템플릿 목록 API (JWT 인증 필요)
@app.get("/{jwt_token}/api/templates")
async def list_templates_for_user(jwt_token: str):
    """사용자용 템플릿 목록 (선택용)"""
    from services.auth_service import decode_jwt_token
    from services.template_service import get_all_templates_summary
    payload = decode_jwt_token(jwt_token)
    if not payload:
        from fastapi import HTTPException
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")
    templates = await get_all_templates_summary()
    return {"templates": templates}


@app.get("/{jwt_token}/api/html-skills")
async def list_html_skills_for_user(jwt_token: str):
    """사용자용 HTML 스킬 목록 (발행된 스킬만)"""
    from services.auth_service import decode_jwt_token
    from services.mongo_service import get_db
    payload = decode_jwt_token(jwt_token)
    if not payload:
        from fastapi import HTTPException
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")
    db = get_db()
    cursor = db.html_skills.find({"is_published": True}).sort("title", 1)
    skills = []
    async for s in cursor:
        s["_id"] = str(s["_id"])
        skills.append(s)
    return {"skills": skills}


@app.get("/{jwt_token}/api/templates/{template_id}/slides")
async def get_template_slides_for_user(jwt_token: str, template_id: str):
    """사용자용 템플릿 슬라이드 목록 (수동 모드 슬라이드 선택용)"""
    from services.auth_service import decode_jwt_token
    from services.template_service import get_template_slides
    payload = decode_jwt_token(jwt_token)
    if not payload:
        from fastapi import HTTPException
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")
    slides = await get_template_slides(template_id)
    return {"slides": slides}


if __name__ == "__main__":
    import signal
    import uvicorn
    import platform

    # Windows에서 Ctrl+C 시 깔끔하게 종료되도록 시그널 핸들러 설정
    def _force_exit(signum, frame):
        """Ctrl+C 시 Redis/MongoDB 연결 정리 후 강제 종료"""
        try:
            close_redis_sync()
            print("\nRedis 연결 해제")
        except Exception:
            pass
        try:
            close_connection_sync()
            print("MongoDB 연결 해제")
        except Exception:
            pass
        print("서버 종료")
        os._exit(0)

    signal.signal(signal.SIGINT, _force_exit)
    if hasattr(signal, "SIGTERM"):
        signal.signal(signal.SIGTERM, _force_exit)

    # 콘솔 타이틀 설정
    _proto = "HTTPS" if (settings.SSL_CERTFILE and settings.SSL_KEYFILE) else "HTTP"
    _title = f"{settings.SOLUTION_NAME} Server | {_proto} :{settings.SERVER_PORT} | DB:{settings.PPTMAKER_DB} | Redis:{settings.REDIS_HOST}:{settings.REDIS_PORT} | {settings.ANTHROPIC_MODEL}"
    if platform.system() == "Windows":
        import ctypes
        ctypes.windll.kernel32.SetConsoleTitleW(_title)
    else:
        sys.stdout.write(f"\033]0;{_title}\007")
        sys.stdout.flush()

    uvicorn_kwargs = {
        "host": settings.SERVER_HOST,
        "port": settings.SERVER_PORT,
        "reload": settings.AUTO_RELOAD,
    }

    # SSL 설정
    use_https = False
    if settings.SSL_CERTFILE and settings.SSL_KEYFILE:
        if os.path.exists(settings.SSL_CERTFILE) and os.path.exists(settings.SSL_KEYFILE):
            uvicorn_kwargs["ssl_certfile"] = settings.SSL_CERTFILE
            uvicorn_kwargs["ssl_keyfile"] = settings.SSL_KEYFILE
            if settings.SSL_KEYFILE_PASSWORD:
                uvicorn_kwargs["ssl_keyfile_password"] = settings.SSL_KEYFILE_PASSWORD
            use_https = True

    # 서버 기본 정보 출력
    protocol = "https" if use_https else "http"
    print("")
    print("=" * 60)
    print(f"  {settings.SOLUTION_NAME} Server")
    print("=" * 60)
    print(f"  Protocol     : {protocol.upper()}")
    print(f"  Host         : {settings.SERVER_HOST}")
    print(f"  Port         : {settings.SERVER_PORT}")
    print(f"  URL          : {protocol}://{settings.SERVER_HOST}:{settings.SERVER_PORT}")
    print(f"  Auto Reload  : {'ON' if settings.AUTO_RELOAD else 'OFF'}")
    print("-" * 60)
    print(f"  MongoDB      : {settings.MONGO_URI.split('@')[-1].split('?')[0] if '@' in settings.MONGO_URI else settings.MONGO_URI}")
    print(f"  {settings.SOLUTION_NAME} DB : {settings.PPTMAKER_DB}")
    print(f"  Org DB       : {settings.ORG_DB} / {settings.ORG_COLLECTION}")
    _redis_pw = "(비밀번호 설정)" if settings.REDIS_PASSWORD else "(비밀번호 없음)"
    print(f"  Redis        : {settings.REDIS_HOST}:{settings.REDIS_PORT} DB={settings.REDIS_DB} {_redis_pw}")
    print("-" * 60)
    print(f"  CORS Origins : {', '.join(settings.CORS_ORIGINS)}")
    print(f"  Languages    : {', '.join(settings.SUPPORTED_LANGS)}")
    print(f"  Upload Dir   : {os.path.abspath(settings.UPLOAD_DIR)}")
    print(f"  Max Upload   : {settings.MAX_UPLOAD_SIZE // (1024*1024)} MB")
    print("-" * 60)
    print(f"  Claude Model : {settings.ANTHROPIC_MODEL}")
    print(f"  Claude Key   : {'***' + settings.ANTHROPIC_API_KEY[-8:] if settings.ANTHROPIC_API_KEY else '(미설정)'}")
    print(f"  Perplexity   : {'***' + settings.PERPLEXITY_API_KEY[-8:] if settings.PERPLEXITY_API_KEY else '(미설정)'}")
    print("-" * 60)
    if use_https:
        print(f"  SSL Cert     : {settings.SSL_CERTFILE}")
        print(f"  SSL Key      : {settings.SSL_KEYFILE}")
    else:
        print(f"  SSL          : 비활성 (HTTP 모드)")
    print("-" * 60)
    print(f"  Python       : {platform.python_version()}")
    print(f"  OS           : {platform.system()} {platform.release()}")
    print("=" * 60)
    print("")

    uvicorn.run("main:app", **uvicorn_kwargs)
