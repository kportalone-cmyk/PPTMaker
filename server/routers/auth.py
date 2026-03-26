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
    """요청에서 JWT 토큰을 추출하고 사용자 정보 반환 (내부/외부 JWT 모두 지원, 데모 계정 포함)"""
    token = request.path_params.get("jwt_token", "")
    if not token:
        raise HTTPException(status_code=401, detail="인증 토큰이 필요합니다")

    payload = decode_jwt_token(token)
    if not payload:
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")

    user = await get_user_flexible(payload)
    if not user:
        # 조직도에 없으면 demo_accounts에서 조회
        user_key = extract_user_key(payload)
        if user_key:
            db = get_db()
            user = await db.demo_accounts.find_one({"user_key": user_key})
    if not user:
        raise HTTPException(status_code=404, detail="사용자를 찾을 수 없습니다")

    return user


@router.post("/login")
async def login(req: LoginRequest):
    """로그인 후 JWT 토큰 발급 (user_key 또는 user_name 으로 로그인)"""
    db = get_db()

    user_key = req.user_key.strip() if req.user_key else ""
    user_name = req.user_name.strip() if req.user_name else ""

    # user_name 으로 로그인 시 사용자 조회
    if not user_key and user_name:
        import re
        name_regex = re.compile(f"^{re.escape(user_name)}$", re.IGNORECASE)
        # 1) 조직도에서 검색 (대소문자 무시)
        users = await search_users_by_name(user_name)
        exact = [u for u in users if name_regex.match(u.get("nm", ""))]
        if len(exact) == 1:
            user_key = exact[0].get("ky", "")
        elif len(exact) > 1:
            raise HTTPException(status_code=401, detail="동명이인이 존재합니다. 관리자에게 문의하세요.")
        else:
            # 2) 조직도에 없으면 demo_accounts 컬렉션에서 이름으로 검색 (대소문자 무시)
            demo = await db.demo_accounts.find_one({"nm": name_regex})
            if demo:
                user_key = demo.get("user_key", "")
            else:
                raise HTTPException(status_code=401, detail="사용자를 찾을 수 없습니다")

    if not user_key:
        raise HTTPException(status_code=401, detail="사용자 이름을 입력하세요")

    # 계정 정보 확인 (accounts → demo_accounts 순서)
    account = await db.accounts.find_one({"user_key": user_key})
    is_demo = False
    if not account:
        account = await db.demo_accounts.find_one({"user_key": user_key})
        is_demo = True

    if not account:
        raise HTTPException(status_code=401, detail="등록되지 않은 계정입니다")

    if not verify_password(req.password, account.get("password", "")):
        raise HTTPException(status_code=401, detail="비밀번호가 일치하지 않습니다")

    # 사용자 정보: 데모 계정은 demo_accounts 자체 프로필 사용
    if is_demo:
        user = account
    else:
        user = await get_user_by_key(user_key)
        if not user:
            raise HTTPException(status_code=401, detail="사용자를 찾을 수 없습니다")

    token = create_jwt_token({
        "user_key": user_key,
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
        # 조직도에 없으면 demo_accounts에서 조회
        user_key = extract_user_key(payload)
        if user_key:
            db = get_db()
            user = await db.demo_accounts.find_one({"user_key": user_key})
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
