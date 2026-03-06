import os
from dotenv import load_dotenv
from pathlib import Path

# .env 파일 로드 (프로젝트 루트)
env_path = Path(__file__).resolve().parent.parent / ".env"
load_dotenv(dotenv_path=env_path)


class Settings:
    # MongoDB
    MONGO_URI: str = os.getenv("MONGO_URI", "")
    PPTMAKER_DB: str = os.getenv("PPTMAKER_DB", "PPTMaker")
    ORG_DB: str = os.getenv("ORG_DB", "im_org_info")
    ORG_COLLECTION: str = os.getenv("ORG_COLLECTION", "user_info")

    # JWT
    JWT_SECRET: str = os.getenv("JWT_SECRET", "")
    JWT_ALGORITHM: str = os.getenv("JWT_ALGORITHM", "HS256")
    JWT_EXPIRE_HOURS: int = int(os.getenv("JWT_EXPIRE_HOURS", "24"))

    # Claude API
    ANTHROPIC_API_KEY: str = os.getenv("ANTHROPIC_API_KEY", "")
    ANTHROPIC_MODEL: str = os.getenv("ANTHROPIC_MODEL", "claude-opus-4-6")
    ANTHROPIC_OUTLINE_MODEL: str = os.getenv("ANTHROPIC_OUTLINE_MODEL", "claude-sonnet-4-6")
    ANTHROPIC_MAX_TOKENS: int = int(os.getenv("ANTHROPIC_MAX_TOKENS", "0"))  # 0이면 API 기본값 사용
    ANTHROPIC_OUTLINE_MAX_TOKENS: int = int(os.getenv("ANTHROPIC_OUTLINE_MAX_TOKENS", "0"))  # Outline(Sonnet) 전용

    # Perplexity
    PERPLEXITY_API_KEY: str = os.getenv("PERPLEXITY_API_KEY", "")

    # Server
    SERVER_HOST: str = os.getenv("SERVER_HOST", "0.0.0.0")
    SERVER_PORT: int = int(os.getenv("SERVER_PORT", "8000"))

    # Upload
    UPLOAD_DIR: str = os.getenv("UPLOAD_DIR", "./uploads")
    MAX_UPLOAD_SIZE: int = int(os.getenv("MAX_UPLOAD_SIZE", "52428800"))

    # CORS
    CORS_ORIGINS: list = [
        o.strip() for o in os.getenv("CORS_ORIGINS", "*").split(",") if o.strip()
    ]

    # 지원 언어
    SUPPORTED_LANGS: list = [
        l.strip() for l in os.getenv("SUPPORTED_LANGS", "ko").split(",") if l.strip()
    ]

    # SSL
    SSL_CERTFILE: str = os.getenv("SSL_CERTFILE", "")
    SSL_KEYFILE: str = os.getenv("SSL_KEYFILE", "")
    SSL_KEYFILE_PASSWORD: str = os.getenv("SSL_KEYFILE_PASSWORD", "")


settings = Settings()
