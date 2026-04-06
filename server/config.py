import os
from dotenv import load_dotenv
from pathlib import Path

# .env 파일 로드 (프로젝트 루트)
env_path = Path(__file__).resolve().parent.parent / ".env"
load_dotenv(dotenv_path=env_path)


class Settings:
    # Solution
    SOLUTION_NAME: str = os.getenv("SOLUTION_NAME", "K-Portal OfficeMaker")

    # MongoDB
    MONGO_URI: str = os.getenv("MONGO_URI", "")
    PPTMAKER_DB: str = os.getenv("PPTMAKER_DB", "PPTMaker")
    ORG_DB: str = os.getenv("ORG_DB", "im_org_info")
    ORG_COLLECTION: str = os.getenv("ORG_COLLECTION", "user_info")

    # JWT
    JWT_SECRET: str = os.getenv("JWT_SECRET", "")
    JWT_ALGORITHM: str = os.getenv("JWT_ALGORITHM", "HS256")
    JWT_EXPIRE_HOURS: int = int(os.getenv("JWT_EXPIRE_HOURS", "24"))

    # Claude API (콤마 구분으로 여러 키 등록 가능 → 라운드 로빈)
    ANTHROPIC_API_KEY: str = os.getenv("ANTHROPIC_API_KEY", "")
    ANTHROPIC_API_KEYS: list = [
        k.strip() for k in os.getenv("ANTHROPIC_API_KEY", "").split(",") if k.strip()
    ]
    ANTHROPIC_MODEL: str = os.getenv("ANTHROPIC_MODEL", "claude-opus-4-6")
    ANTHROPIC_OUTLINE_MODEL: str = os.getenv("ANTHROPIC_OUTLINE_MODEL", "claude-sonnet-4-6")
    ANTHROPIC_MAX_TOKENS: int = int(os.getenv("ANTHROPIC_MAX_TOKENS", "0"))  # 0이면 API 기본값 사용
    ANTHROPIC_OUTLINE_MAX_TOKENS: int = int(os.getenv("ANTHROPIC_OUTLINE_MAX_TOKENS", "0"))  # Outline(Sonnet) 전용

    # Google AI (콤마 구분으로 여러 키 등록 가능 → 라운드 로빈)
    GOOGLE_API_KEY: str = os.getenv("GOOGLE_API_KEY", "")
    GOOGLE_API_KEYS: list = [
        k.strip() for k in os.getenv("GOOGLE_API_KEY", "").split(",") if k.strip()
    ]
    GOOGLE_IMAGE_MODEL: str = os.getenv("GOOGLE_IMAGE_MODEL", "gemini-3.1-flash-image-preview")
    GOOGLE_IMAGE_THINKING: str = os.getenv("GOOGLE_IMAGE_THINKING", "HIGH")  # none, LOW, MEDIUM, HIGH

    # Perplexity
    PERPLEXITY_API_KEY: str = os.getenv("PERPLEXITY_API_KEY", "")

    # Server
    SERVER_HOST: str = os.getenv("SERVER_HOST", "0.0.0.0")
    SERVER_PORT: int = int(os.getenv("SERVER_PORT", "8000"))
    AUTO_RELOAD: bool = os.getenv("AUTO_RELOAD", "true").lower() in ("true", "1", "yes")

    # Upload (절대경로 또는 프로젝트 루트 기준 상대경로)
    _raw_upload_dir: str = os.getenv("UPLOAD_DIR", "./uploads")
    UPLOAD_DIR: str = str(
        Path(_raw_upload_dir).resolve()
        if Path(_raw_upload_dir).is_absolute()
        else (Path(__file__).resolve().parent.parent / _raw_upload_dir).resolve()
    )
    MAX_UPLOAD_SIZE: int = int(os.getenv("MAX_UPLOAD_SIZE", "52428800"))

    # CORS
    CORS_ORIGINS: list = [
        o.strip() for o in os.getenv("CORS_ORIGINS", "*").split(",") if o.strip()
    ]

    # 지원 언어
    SUPPORTED_LANGS: list = [
        l.strip() for l in os.getenv("SUPPORTED_LANGS", "ko").split(",") if l.strip()
    ]

    # Redis
    REDIS_HOST: str = os.getenv("REDIS_HOST", "127.0.0.1")
    REDIS_PORT: int = int(os.getenv("REDIS_PORT", "6379"))
    REDIS_DB: int = int(os.getenv("REDIS_DB", "0"))
    REDIS_PASSWORD: str = os.getenv("REDIS_PASSWORD", "")

    # OnlyOffice
    ONLYOFFICE_URL: str = os.getenv("ONLYOFFICE_URL", "")
    ONLYOFFICE_JWT_SECRET: str = os.getenv("ONLYOFFICE_JWT_SECRET", "")
    SERVER_BASE_URL: str = os.getenv("SERVER_BASE_URL", "")

    # Infographic
    DEFAULT_INFOGRAPHIC_RATIO: int = int(os.getenv("DEFAULT_INFOGRAPHIC_RATIO", "60"))

    # SSL
    SSL_CERTFILE: str = os.getenv("SSL_CERTFILE", "")
    SSL_KEYFILE: str = os.getenv("SSL_KEYFILE", "")
    SSL_KEYFILE_PASSWORD: str = os.getenv("SSL_KEYFILE_PASSWORD", "")


settings = Settings()


# ── API 키 라운드 로빈 ──
import itertools
import threading

class _KeyRotator:
    """Thread-safe 라운드 로빈 키 로테이터"""
    def __init__(self, keys: list):
        self._keys = keys or [""]
        self._cycle = itertools.cycle(self._keys)
        self._lock = threading.Lock()

    def next(self) -> str:
        with self._lock:
            return next(self._cycle)

    @property
    def count(self) -> int:
        return len(self._keys)


anthropic_key_rotator = _KeyRotator(settings.ANTHROPIC_API_KEYS)
google_key_rotator = _KeyRotator(settings.GOOGLE_API_KEYS)
