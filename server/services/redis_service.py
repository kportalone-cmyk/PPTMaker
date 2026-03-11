"""
Redis 서비스 모듈
- 슬라이드 잠금, 온라인 상태, 생성 취소, 사용자 캐시
- 모든 함수는 Redis 장애 시 None 반환 → 호출측에서 MongoDB 폴백
"""
import json
import asyncio
from datetime import datetime
from redis.asyncio import Redis
from config import settings

_redis: Redis | None = None
_available: bool = False

# ─── 키 접두사 ───
_PREFIX = "officemaker"


def _key(*parts: str) -> str:
    return f"{_PREFIX}:{':'.join(parts)}"


# ──────────────────────────────────────────────
# 연결 관리
# ──────────────────────────────────────────────

async def init_redis() -> bool:
    """Redis 연결 초기화 및 ping 확인"""
    global _redis, _available
    try:
        _redis = Redis(
            host=settings.REDIS_HOST,
            port=settings.REDIS_PORT,
            db=settings.REDIS_DB,
            password=settings.REDIS_PASSWORD or None,
            decode_responses=True,
            socket_connect_timeout=3,
            socket_timeout=3,
        )
        await _redis.ping()
        _available = True
        return True
    except Exception as e:
        print(f"[Redis] 연결 실패: {e}")
        _available = False
        _redis = None
        return False


async def close_redis():
    """Redis 연결 종료 (2초 타임아웃)"""
    global _redis, _available
    _available = False
    if _redis:
        try:
            await asyncio.wait_for(_redis.close(), timeout=2.0)
        except (asyncio.TimeoutError, Exception):
            pass
        _redis = None


def close_redis_sync():
    """동기 버전 - 시그널 핸들러용 (강제 종료)"""
    global _redis, _available
    _available = False
    if _redis:
        try:
            # 연결 풀 강제 종료
            _redis.connection_pool.disconnect()
        except Exception:
            pass
        _redis = None


def is_available() -> bool:
    return _available and _redis is not None


def get_redis() -> Redis | None:
    return _redis if _available else None


# ──────────────────────────────────────────────
# 슬라이드 잠금 (Lock)
# ──────────────────────────────────────────────

async def acquire_slide_lock(
    project_id: str, slide_id: str,
    user_key: str, user_name: str,
    ttl: int = 300,
) -> dict | None:
    """
    슬라이드 잠금 획득.
    Returns:
      {"acquired": True} - 잠금 성공
      {"acquired": False, "holder": {...}} - 다른 사용자가 잠금 중
      None - Redis 불가 (MongoDB 폴백 필요)
    """
    r = get_redis()
    if not r:
        return None
    try:
        key = _key("lock", project_id, slide_id)
        existing = await r.get(key)

        if existing:
            data = json.loads(existing)
            if data.get("user_key") == user_key:
                # 본인 잠금 갱신
                data["locked_at"] = datetime.utcnow().isoformat()
                await r.set(key, json.dumps(data), ex=ttl)
                return {"acquired": True}
            else:
                return {"acquired": False, "holder": data}

        # 새 잠금 (SET NX)
        lock_data = json.dumps({
            "user_key": user_key,
            "user_name": user_name,
            "locked_at": datetime.utcnow().isoformat(),
        })
        result = await r.set(key, lock_data, ex=ttl, nx=True)
        if result:
            return {"acquired": True}

        # 다른 사용자가 먼저 획득 (race condition)
        existing = await r.get(key)
        if existing:
            return {"acquired": False, "holder": json.loads(existing)}
        return {"acquired": True}
    except Exception:
        return None


async def release_slide_lock(
    project_id: str, slide_id: str, user_key: str
) -> bool | None:
    """
    슬라이드 잠금 해제. 본인 잠금만 해제.
    Returns: True(해제됨), False(권한없음), None(Redis 불가)
    """
    r = get_redis()
    if not r:
        return None
    try:
        key = _key("lock", project_id, slide_id)
        existing = await r.get(key)
        if not existing:
            return True
        data = json.loads(existing)
        if data.get("user_key") != user_key:
            return False
        await r.delete(key)
        return True
    except Exception:
        return None


async def renew_slide_lock(
    project_id: str, user_key: str, ttl: int = 300
) -> str | None:
    """
    프로젝트 내 해당 사용자의 모든 잠금 TTL 갱신.
    Returns: 갱신된 slide_id 또는 None
    """
    r = get_redis()
    if not r:
        return None
    try:
        pattern = _key("lock", project_id, "*")
        renewed_slide_id = None
        async for key in r.scan_iter(match=pattern, count=100):
            val = await r.get(key)
            if val:
                data = json.loads(val)
                if data.get("user_key") == user_key:
                    await r.expire(key, ttl)
                    # 키에서 slide_id 추출
                    parts = key.split(":")
                    if len(parts) >= 4:
                        renewed_slide_id = parts[-1]
        return renewed_slide_id
    except Exception:
        return None


async def renew_specific_lock(
    project_id: str, slide_id: str, user_key: str, ttl: int = 300
) -> bool | None:
    """
    특정 슬라이드 Lock만 갱신하고, 해당 사용자의 다른 Lock은 삭제.
    Returns: True(갱신됨), False(Lock 없음), None(Redis 불가)
    """
    r = get_redis()
    if not r:
        return None
    try:
        target_key = _key("lock", project_id, slide_id)
        # 대상 Lock 갱신
        val = await r.get(target_key)
        renewed = False
        if val:
            data = json.loads(val)
            if data.get("user_key") == user_key:
                await r.expire(target_key, ttl)
                renewed = True

        # 다른 좀비 Lock 정리
        pattern = _key("lock", project_id, "*")
        async for key in r.scan_iter(match=pattern, count=100):
            if key == target_key:
                continue
            v = await r.get(key)
            if v:
                d = json.loads(v)
                if d.get("user_key") == user_key:
                    await r.delete(key)

        return renewed
    except Exception:
        return None


async def get_project_locks(project_id: str) -> list | None:
    """
    프로젝트의 모든 활성 잠금 조회.
    Returns: [{"slide_id", "user_key", "user_name", "locked_at"}] 또는 None
    """
    r = get_redis()
    if not r:
        return None
    try:
        pattern = _key("lock", project_id, "*")
        locks = []
        async for key in r.scan_iter(match=pattern, count=100):
            val = await r.get(key)
            if val:
                data = json.loads(val)
                parts = key.split(":")
                slide_id = parts[-1] if len(parts) >= 4 else ""
                ttl_remaining = await r.ttl(key)
                locks.append({
                    "slide_id": slide_id,
                    "user_key": data.get("user_key", ""),
                    "user_name": data.get("user_name", ""),
                    "locked_at": data.get("locked_at", ""),
                    "ttl": ttl_remaining,
                })
        return locks
    except Exception:
        return None


async def delete_project_locks(project_id: str, user_key: str = None) -> bool | None:
    """프로젝트의 잠금 삭제 (user_key 지정 시 해당 사용자 잠금만)"""
    r = get_redis()
    if not r:
        return None
    try:
        pattern = _key("lock", project_id, "*")
        async for key in r.scan_iter(match=pattern, count=100):
            if user_key:
                val = await r.get(key)
                if val:
                    data = json.loads(val)
                    if data.get("user_key") == user_key:
                        await r.delete(key)
            else:
                await r.delete(key)
        return True
    except Exception:
        return None


# ──────────────────────────────────────────────
# 온라인 상태 (Presence)
# ──────────────────────────────────────────────

async def update_presence(
    project_id: str, user_key: str, user_name: str, ttl: int = 90
) -> bool | None:
    """온라인 상태 갱신"""
    r = get_redis()
    if not r:
        return None
    try:
        key = _key("presence", project_id)
        data = json.dumps({
            "user_name": user_name,
            "last_seen": datetime.utcnow().isoformat(),
        })
        await r.hset(key, user_key, data)
        await r.expire(key, ttl)
        return True
    except Exception:
        return None


async def get_online_users(project_id: str) -> list | None:
    """프로젝트의 온라인 사용자 목록"""
    r = get_redis()
    if not r:
        return None
    try:
        key = _key("presence", project_id)
        all_data = await r.hgetall(key)
        users = []
        now = datetime.utcnow()
        for uk, val in all_data.items():
            data = json.loads(val)
            last_seen_str = data.get("last_seen", "")
            if last_seen_str:
                last_seen = datetime.fromisoformat(last_seen_str)
                # 90초 이내만 온라인으로 표시
                diff = (now - last_seen).total_seconds()
                if diff <= 90:
                    users.append({
                        "user_key": uk,
                        "user_name": data.get("user_name", ""),
                        "last_seen": last_seen_str,
                    })
        return users
    except Exception:
        return None


async def remove_presence(project_id: str, user_key: str) -> bool | None:
    """온라인 상태 제거"""
    r = get_redis()
    if not r:
        return None
    try:
        key = _key("presence", project_id)
        await r.hdel(key, user_key)
        return True
    except Exception:
        return None


# ──────────────────────────────────────────────
# 생성 취소 플래그
# ──────────────────────────────────────────────

async def set_generation_cancel(project_id: str, ttl: int = 300) -> bool | None:
    """생성 취소 플래그 설정"""
    r = get_redis()
    if not r:
        return None
    try:
        key = _key("cancel", project_id)
        await r.set(key, "1", ex=ttl)
        return True
    except Exception:
        return None


async def check_generation_cancel(project_id: str) -> bool | None:
    """
    생성 취소 여부 확인.
    Returns: True(취소됨), False(진행중), None(Redis 불가→MongoDB 폴백)
    """
    r = get_redis()
    if not r:
        return None
    try:
        key = _key("cancel", project_id)
        val = await r.get(key)
        return val is not None
    except Exception:
        return None


async def clear_generation_cancel(project_id: str) -> bool | None:
    """생성 취소 플래그 제거"""
    r = get_redis()
    if not r:
        return None
    try:
        key = _key("cancel", project_id)
        await r.delete(key)
        return True
    except Exception:
        return None


# ──────────────────────────────────────────────
# 사용자 캐시
# ──────────────────────────────────────────────

async def cache_user(user_key: str, user_data: dict, ttl: int = 900) -> bool | None:
    """사용자 정보 캐시 (15분)"""
    r = get_redis()
    if not r:
        return None
    try:
        key = _key("user", user_key)
        # _id 등 ObjectId 제거
        safe = {k: v for k, v in user_data.items() if isinstance(v, (str, int, float, bool))}
        await r.set(key, json.dumps(safe), ex=ttl)
        return True
    except Exception:
        return None


async def get_cached_user(user_key: str) -> dict | None:
    """캐시된 사용자 정보 조회. 없으면 None."""
    r = get_redis()
    if not r:
        return None
    try:
        key = _key("user", user_key)
        val = await r.get(key)
        if val:
            return json.loads(val)
        return None
    except Exception:
        return None


# ──────────────────────────────────────────────
# 범용 데이터 캐시 (템플릿/폰트/프롬프트 등)
# ──────────────────────────────────────────────

async def cache_set(cache_key: str, data, ttl: int = 3600) -> bool | None:
    """범용 캐시 저장"""
    r = get_redis()
    if not r:
        return None
    try:
        key = _key("cache", cache_key)
        await r.set(key, json.dumps(data, ensure_ascii=False, default=str), ex=ttl)
        return True
    except Exception:
        return None


async def cache_get(cache_key: str):
    """범용 캐시 조회. 없거나 Redis 불가 시 None."""
    r = get_redis()
    if not r:
        return None
    try:
        key = _key("cache", cache_key)
        val = await r.get(key)
        if val:
            return json.loads(val)
        return None
    except Exception:
        return None


async def cache_delete(*cache_keys: str) -> bool | None:
    """캐시 키 삭제 (여러 키 동시 삭제 가능)"""
    r = get_redis()
    if not r:
        return None
    try:
        keys = [_key("cache", k) for k in cache_keys]
        await r.delete(*keys)
        return True
    except Exception:
        return None


async def cache_delete_pattern(pattern: str) -> bool | None:
    """패턴 기반 캐시 삭제 (예: 'template:*')"""
    r = get_redis()
    if not r:
        return None
    try:
        full_pattern = _key("cache", pattern)
        async for key in r.scan_iter(match=full_pattern, count=100):
            await r.delete(key)
        return True
    except Exception:
        return None
