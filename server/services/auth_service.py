import base64
from datetime import datetime, timedelta
from jose import jwt, JWTError
from passlib.context import CryptContext
from config import settings
from services.mongo_service import get_org_db
from services import redis_service

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")


def hash_password(password: str) -> str:
    return pwd_context.hash(password)


def verify_password(plain: str, hashed: str) -> bool:
    return pwd_context.verify(plain, hashed)


def create_jwt_token(data: dict) -> str:
    to_encode = data.copy()
    expire = datetime.utcnow() + timedelta(hours=settings.JWT_EXPIRE_HOURS)
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, settings.JWT_SECRET, algorithm=settings.JWT_ALGORITHM)


def decode_jwt_token(token: str) -> dict | None:
    """JWT 복호화 (raw / base64 인코딩 시크릿 모두 시도)"""
    secrets = [
        settings.JWT_SECRET,
        base64.b64encode(settings.JWT_SECRET.encode("utf-8")),
    ]
    for secret in secrets:
        try:
            return jwt.decode(token, secret, algorithms=[settings.JWT_ALGORITHM])
        except JWTError:
            continue
    return None


def extract_user_key(payload: dict) -> str:
    """JWT 페이로드에서 사용자 키 추출 (내부/외부 JWT 모두 지원)
    - 내부 JWT: user_key 필드
    - 외부 JWT: email 필드 (조직도 ky에 해당)
    """
    return payload.get("user_key") or payload.get("email") or ""


async def get_user_by_key(user_key: str) -> dict | None:
    """조직도에서 사용자 조회 (ky 기준, Redis 캐시 우선)"""
    # Redis 캐시 확인
    cached = await redis_service.get_cached_user(user_key)
    if cached:
        return cached

    # MongoDB 조회
    org_db = get_org_db()
    col = org_db[settings.ORG_COLLECTION]
    user = await col.find_one({"ky": user_key})
    if user:
        user["_id"] = str(user["_id"])
        # Redis에 캐시 (15분)
        await redis_service.cache_user(user_key, user)
    return user


async def get_user_flexible(payload: dict) -> dict | None:
    """다양한 JWT 형식에서 사용자 조회 (내부/외부 JWT 모두 지원)
    1) user_key 또는 email → ky 매칭
    2) userid → em(메일) 매칭
    """
    # 1차: ky 기준 조회 (Redis 캐시 포함)
    user_key = extract_user_key(payload)
    if user_key:
        user = await get_user_by_key(user_key)
        if user:
            return user

    # 2차: userid → em(메일) 기준 조회
    userid = payload.get("userid") or ""
    if userid:
        # Redis 캐시 확인 (em 기준)
        cached = await redis_service.get_cached_user(f"em:{userid}")
        if cached:
            return cached

        org_db = get_org_db()
        col = org_db[settings.ORG_COLLECTION]
        user = await col.find_one({"em": userid})
        if user:
            user["_id"] = str(user["_id"])
            # ky 기준과 em 기준 양쪽으로 캐시
            await redis_service.cache_user(user.get("ky", userid), user)
            await redis_service.cache_user(f"em:{userid}", user)
            return user

    return None


async def search_users_by_name(name: str) -> list:
    """이름으로 사용자 검색 (동명이인 대응)"""
    org_db = get_org_db()
    col = org_db[settings.ORG_COLLECTION]
    cursor = col.find({"nm": {"$regex": name, "$options": "i"}}).limit(20)
    users = []
    async for user in cursor:
        user["_id"] = str(user["_id"])
        users.append(user)
    return users


def is_admin(user: dict) -> bool:
    """관리자 여부 확인"""
    return user.get("role", "").lower() == "admin"
