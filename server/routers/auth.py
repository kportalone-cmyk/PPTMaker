from fastapi import APIRouter, HTTPException, Request
from models.user import LoginRequest, UserSearchRequest
from services.auth_service import (
    create_jwt_token, decode_jwt_token, extract_user_key,
    get_user_by_key, get_user_flexible, search_users_by_name, is_admin,
    verify_password
)
from services.mongo_service import get_db

router = APIRouter(prefix="/api/auth", tags=["auth"])


async def get_current_user(request: Request) -> dict:
    """요청에서 JWT 토큰을 추출하고 사용자 정보 반환 (내부/외부 JWT 모두 지원)"""
    token = request.path_params.get("jwt_token", "")
    if not token:
        raise HTTPException(status_code=401, detail="인증 토큰이 필요합니다")

    payload = decode_jwt_token(token)
    if not payload:
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")

    user = await get_user_flexible(payload)
    if not user:
        raise HTTPException(status_code=404, detail="사용자를 찾을 수 없습니다")

    return user


@router.post("/login")
async def login(req: LoginRequest):
    """로그인 후 JWT 토큰 발급"""
    db = get_db()
    # 계정 정보 확인
    account = await db.accounts.find_one({"user_key": req.user_key})

    if not account:
        # 조직도에서 사용자 확인
        user = await get_user_by_key(req.user_key)
        if not user:
            raise HTTPException(status_code=401, detail="사용자를 찾을 수 없습니다")
        raise HTTPException(status_code=401, detail="등록되지 않은 계정입니다")

    if not verify_password(req.password, account.get("password", "")):
        raise HTTPException(status_code=401, detail="비밀번호가 일치하지 않습니다")

    user = await get_user_by_key(req.user_key)
    token = create_jwt_token({
        "user_key": req.user_key,
        "nm": user.get("nm", ""),
        "role": user.get("role", ""),
    })

    return {
        "token": token,
        "user": {
            "nm": user.get("nm"),
            "dp": user.get("dp"),
            "em": user.get("em"),
            "role": user.get("role"),
        }
    }


@router.post("/search-users")
async def search_users(req: UserSearchRequest):
    """이름으로 사용자 검색 (동명이인 대응)"""
    users = await search_users_by_name(req.name)
    return {
        "users": [
            {
                "nm": u.get("nm"),
                "dp": u.get("dp"),
                "em": u.get("em"),
                "ky": u.get("ky"),
            }
            for u in users
        ]
    }


@router.get("/verify/{jwt_token}")
async def verify_token(jwt_token: str):
    """JWT 토큰 유효성 검증 (내부/외부 JWT 모두 지원)"""
    payload = decode_jwt_token(jwt_token)
    if not payload:
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")

    user = await get_user_flexible(payload)
    if not user:
        raise HTTPException(status_code=404, detail="사용자를 찾을 수 없습니다")

    return {
        "valid": True,
        "user": {
            "nm": user.get("nm"),
            "dp": user.get("dp"),
            "em": user.get("em"),
            "role": user.get("role"),
            "ky": user.get("ky"),
        },
        "is_admin": is_admin(user)
    }
