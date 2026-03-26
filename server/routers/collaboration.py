import json
import asyncio
from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse
from bson import ObjectId
from datetime import datetime, timedelta
from models.collaboration import CollaboratorAdd, CollaboratorUpdate
from services.mongo_service import get_db
from services.auth_service import (
    decode_jwt_token, extract_user_key, get_user_flexible, get_user_by_key
)
from services import redis_service

router = APIRouter(tags=["collaboration"])

_PREFIX = "officemaker"
LOCK_TTL_MINUTES = 5
LOCK_TTL_SECONDS = LOCK_TTL_MINUTES * 60


async def get_user_key(jwt_token: str) -> str:
    """JWT에서 user_key 추출 (내부/외부 JWT 모두 지원)"""
    payload = decode_jwt_token(jwt_token)
    if not payload:
        raise HTTPException(status_code=401, detail="유효하지 않은 토큰입니다")
    user_key = extract_user_key(payload)
    if user_key:
        return user_key
    user = await get_user_flexible(payload)
    if user:
        return user.get("ky", "")
    raise HTTPException(status_code=401, detail="사용자를 확인할 수 없습니다")


async def check_project_access(
    db, project_id: str, user_key: str, required_role: str = "viewer"
) -> dict:
    """
    프로젝트 접근 권한 확인.
    Returns {"role": "owner"|"editor"|"viewer", "project": doc}
    required_role: "viewer" (모든 접근), "editor" (편집 가능), "owner" (소유자만)
    """
    project = await db.projects.find_one({"_id": ObjectId(project_id)})
    if not project:
        raise HTTPException(status_code=404, detail="프로젝트를 찾을 수 없습니다")

    project["_id"] = str(project["_id"])

    # 소유자 확인
    if project.get("user_key") == user_key:
        return {"role": "owner", "project": project}

    if required_role == "owner":
        raise HTTPException(status_code=403, detail="프로젝트 소유자만 가능합니다")

    # 협업자 확인
    collab = await db.collaborators.find_one({
        "project_id": project_id,
        "user_key": user_key
    })
    if not collab:
        raise HTTPException(status_code=403, detail="접근 권한이 없습니다")

    role = collab.get("role", "viewer")
    if required_role == "editor" and role == "viewer":
        raise HTTPException(status_code=403, detail="편집 권한이 없습니다")

    return {"role": role, "project": project}


# ──────────────────────────────────────────────
# 협업자 관리
# ──────────────────────────────────────────────

@router.get("/{jwt_token}/api/projects/{project_id}/collaborators")
async def list_collaborators(jwt_token: str, project_id: str):
    """협업자 목록 조회"""
    user_key = await get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "viewer")

    cursor = db.collaborators.find({"project_id": project_id})
    result = []
    async for c in cursor:
        c["_id"] = str(c["_id"])
        result.append(c)
    return {"collaborators": result}


@router.post("/{jwt_token}/api/projects/{project_id}/collaborators")
async def add_collaborator(jwt_token: str, project_id: str, data: CollaboratorAdd):
    """협업자 추가 (소유자 전용)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "owner")

    if data.user_key == user_key:
        raise HTTPException(status_code=400, detail="자신은 추가할 수 없습니다")

    if data.role not in ("editor", "viewer"):
        raise HTTPException(status_code=400, detail="역할은 editor 또는 viewer만 가능합니다")

    target_user = await get_user_by_key(data.user_key)
    if not target_user:
        raise HTTPException(status_code=404, detail="사용자를 찾을 수 없습니다")

    await db.collaborators.update_one(
        {"project_id": project_id, "user_key": data.user_key},
        {"$set": {
            "project_id": project_id,
            "user_key": data.user_key,
            "user_name": target_user.get("nm", ""),
            "user_dept": target_user.get("dp", ""),
            "role": data.role,
            "invited_by": user_key,
            "created_at": datetime.utcnow(),
        }},
        upsert=True
    )
    return {"success": True}


@router.put("/{jwt_token}/api/projects/{project_id}/collaborators/{target_user_key}")
async def update_collaborator_role(
    jwt_token: str, project_id: str, target_user_key: str, data: CollaboratorUpdate
):
    """협업자 역할 변경 (소유자 전용)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "owner")

    if data.role not in ("editor", "viewer"):
        raise HTTPException(status_code=400, detail="역할은 editor 또는 viewer만 가능합니다")

    result = await db.collaborators.update_one(
        {"project_id": project_id, "user_key": target_user_key},
        {"$set": {"role": data.role}}
    )
    if result.matched_count == 0:
        raise HTTPException(status_code=404, detail="협업자를 찾을 수 없습니다")
    return {"success": True}


@router.delete("/{jwt_token}/api/projects/{project_id}/collaborators/{target_user_key}")
async def remove_collaborator(
    jwt_token: str, project_id: str, target_user_key: str
):
    """협업자 제거 (소유자 또는 본인 탈퇴)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()
    access = await check_project_access(db, project_id, user_key, "viewer")
    if access["role"] != "owner" and user_key != target_user_key:
        raise HTTPException(status_code=403, detail="권한이 없습니다")

    await db.collaborators.delete_one(
        {"project_id": project_id, "user_key": target_user_key}
    )
    # 제거된 사용자의 활성 Lock도 해제 (Redis + MongoDB)
    await redis_service.delete_project_locks(project_id, target_user_key)
    await db.slide_locks.delete_many(
        {"project_id": project_id, "user_key": target_user_key}
    )
    return {"success": True}


# ──────────────────────────────────────────────
# 슬라이드 Lock 관리 (Redis 우선, MongoDB 폴백)
# ──────────────────────────────────────────────

@router.post("/{jwt_token}/api/projects/{project_id}/slides/{slide_id}/lock")
async def acquire_lock(jwt_token: str, project_id: str, slide_id: str):
    """슬라이드 Lock 획득 (Redis 우선)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "editor")

    user = await get_user_by_key(user_key)
    user_name = user.get("nm", user_key) if user else user_key

    now = datetime.utcnow()
    expires_at = now + timedelta(minutes=LOCK_TTL_MINUTES)

    # Redis 우선 시도
    redis_result = await redis_service.acquire_slide_lock(
        project_id, slide_id, user_key, user_name, ttl=LOCK_TTL_SECONDS
    )

    if redis_result is not None:
        # Redis 응답 성공
        if not redis_result.get("acquired"):
            holder = redis_result.get("holder", {})
            raise HTTPException(
                status_code=409,
                detail=f"{holder.get('user_name', '다른 사용자')}이(가) 편집 중입니다"
            )
        # MongoDB에도 동기화 (이력/폴백용)
        await db.slide_locks.update_one(
            {"project_id": project_id, "slide_id": slide_id},
            {"$set": {
                "project_id": project_id,
                "slide_id": slide_id,
                "user_key": user_key,
                "user_name": user_name,
                "locked_at": now,
                "expires_at": expires_at,
            }},
            upsert=True
        )
        # SSE 이벤트 발행: lock_changed (acquired)
        try:
            await redis_service.publish_collab_event(project_id, "lock_changed", {
                "slide_id": slide_id,
                "action": "acquired",
                "user_key": user_key,
                "user_name": user_name,
            })
        except Exception:
            pass
        return {"success": True, "expires_at": expires_at.isoformat()}

    # Redis 불가 → MongoDB 폴백
    existing = await db.slide_locks.find_one(
        {"project_id": project_id, "slide_id": slide_id}
    )
    if existing and existing.get("user_key") != user_key:
        if existing.get("expires_at") and existing["expires_at"] > now:
            raise HTTPException(
                status_code=409,
                detail=f"{existing.get('user_name', '다른 사용자')}이(가) 편집 중입니다"
            )

    await db.slide_locks.update_one(
        {"project_id": project_id, "slide_id": slide_id},
        {"$set": {
            "project_id": project_id,
            "slide_id": slide_id,
            "user_key": user_key,
            "user_name": user_name,
            "locked_at": now,
            "expires_at": expires_at,
        }},
        upsert=True
    )
    return {"success": True, "expires_at": expires_at.isoformat()}


@router.delete("/{jwt_token}/api/projects/{project_id}/slides/{slide_id}/lock")
async def release_lock(jwt_token: str, project_id: str, slide_id: str):
    """슬라이드 Lock 해제 (Redis + MongoDB)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "viewer")

    # Redis 해제 시도
    redis_result = await redis_service.release_slide_lock(project_id, slide_id, user_key)
    if redis_result is False:
        raise HTTPException(status_code=403, detail="본인의 잠금만 해제할 수 있습니다")

    # MongoDB도 해제
    lock = await db.slide_locks.find_one(
        {"project_id": project_id, "slide_id": slide_id}
    )
    if lock and lock.get("user_key") != user_key:
        raise HTTPException(status_code=403, detail="본인의 잠금만 해제할 수 있습니다")

    await db.slide_locks.delete_one(
        {"project_id": project_id, "slide_id": slide_id, "user_key": user_key}
    )

    # SSE 이벤트 발행: lock_changed (released)
    try:
        user = await get_user_by_key(user_key)
        u_name = user.get("nm", user_key) if user else user_key
        await redis_service.publish_collab_event(project_id, "lock_changed", {
            "slide_id": slide_id,
            "action": "released",
            "user_key": user_key,
            "user_name": u_name,
        })
    except Exception:
        pass

    return {"success": True}


# ──────────────────────────────────────────────
# Heartbeat (Redis 우선, MongoDB 폴백)
# ──────────────────────────────────────────────

@router.post("/{jwt_token}/api/projects/{project_id}/heartbeat")
async def heartbeat(jwt_token: str, project_id: str, body: dict = None):
    """Lock TTL 갱신 및 온라인 상태 업데이트 (Redis 우선)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "viewer")

    now = datetime.utcnow()
    new_expires = now + timedelta(minutes=LOCK_TTL_MINUTES)

    # 프론트엔드가 전송한 현재 편집 중인 슬라이드 ID
    editing_slide_id = (body or {}).get("editing_slide_id")

    # Redis 우선: Lock TTL 갱신 + 온라인 상태
    if redis_service.is_available():
        if editing_slide_id:
            # 편집 중인 슬라이드 Lock만 갱신, 나머지는 정리
            await redis_service.renew_specific_lock(
                project_id, editing_slide_id, user_key, ttl=LOCK_TTL_SECONDS
            )
            # MongoDB 동기화: 해당 Lock만 갱신
            await db.slide_locks.update_one(
                {"project_id": project_id, "slide_id": editing_slide_id, "user_key": user_key},
                {"$set": {"expires_at": new_expires}}
            )
            # 다른 좀비 Lock 정리
            await db.slide_locks.delete_many({
                "project_id": project_id,
                "user_key": user_key,
                "slide_id": {"$ne": editing_slide_id},
            })
        else:
            # 편집 모드가 아님 → 이 사용자의 모든 좀비 Lock 정리
            await redis_service.delete_project_locks(project_id, user_key)
            await db.slide_locks.delete_many({
                "project_id": project_id,
                "user_key": user_key,
            })

        # 온라인 상태 갱신
        user = await get_user_by_key(user_key)
        user_name = user.get("nm", user_key) if user else user_key
        await redis_service.update_presence(project_id, user_key, user_name, ttl=90)

        # 첫 heartbeat 감지 → user_joined 이벤트 발행
        try:
            r = redis_service.get_redis()
            if r:
                join_key = f"{_PREFIX}:joined:{project_id}:{user_key}"
                is_new = await r.set(join_key, "1", ex=90, nx=True)
                if is_new:
                    await redis_service.publish_collab_event(project_id, "user_joined", {
                        "user_key": user_key,
                        "user_name": user_name,
                    })
                else:
                    # 갱신만 (TTL 리셋)
                    await r.expire(join_key, 90)
        except Exception:
            pass

        # MongoDB 온라인 상태도 갱신
        await db.online_presence.update_one(
            {"project_id": project_id, "user_key": user_key},
            {"$set": {
                "project_id": project_id,
                "user_key": user_key,
                "user_name": user_name,
                "last_seen": now,
            }},
            upsert=True
        )
        return {"success": True}

    # MongoDB 폴백
    if editing_slide_id:
        # 편집 중인 슬라이드 Lock만 갱신
        await db.slide_locks.update_one(
            {"project_id": project_id, "slide_id": editing_slide_id, "user_key": user_key},
            {"$set": {"expires_at": new_expires}}
        )
        # 다른 좀비 Lock 정리
        await db.slide_locks.delete_many({
            "project_id": project_id,
            "user_key": user_key,
            "slide_id": {"$ne": editing_slide_id},
        })
    else:
        # 편집 모드가 아님 → 이 사용자의 모든 좀비 Lock 정리
        await db.slide_locks.delete_many({
            "project_id": project_id,
            "user_key": user_key,
        })

    user = await get_user_by_key(user_key)
    user_name = user.get("nm", user_key) if user else user_key
    await db.online_presence.update_one(
        {"project_id": project_id, "user_key": user_key},
        {"$set": {
            "project_id": project_id,
            "user_key": user_key,
            "user_name": user_name,
            "last_seen": now,
        }},
        upsert=True
    )
    return {"success": True}


# ──────────────────────────────────────────────
# 협업 상태 폴링 (Redis 우선, MongoDB 폴백)
# ──────────────────────────────────────────────

@router.get("/{jwt_token}/api/projects/{project_id}/collab-status")
async def get_collab_status(jwt_token: str, project_id: str):
    """협업 상태 조회: locks, online_users, slide_timestamps"""
    user_key = await get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "viewer")

    now = datetime.utcnow()

    # 잠금 조회 (Redis 우선)
    locks = []
    redis_locks = await redis_service.get_project_locks(project_id)
    if redis_locks is not None:
        for rl in redis_locks:
            ttl = rl.get("ttl", 0)
            if ttl > 0:
                expires_at = now + timedelta(seconds=ttl)
                locks.append({
                    "slide_id": rl["slide_id"],
                    "user_key": rl["user_key"],
                    "user_name": rl.get("user_name", ""),
                    "expires_at": expires_at.isoformat(),
                })
    else:
        # MongoDB 폴백
        async for lock in db.slide_locks.find(
            {"project_id": project_id, "expires_at": {"$gt": now}},
            {"slide_id": 1, "user_key": 1, "user_name": 1, "expires_at": 1}
        ):
            locks.append({
                "slide_id": lock["slide_id"],
                "user_key": lock["user_key"],
                "user_name": lock.get("user_name", ""),
                "expires_at": lock["expires_at"].isoformat(),
            })

    # 온라인 사용자 (Redis 우선)
    online_users = []
    redis_users = await redis_service.get_online_users(project_id)
    if redis_users is not None:
        online_users = redis_users
    else:
        # MongoDB 폴백
        cutoff = now - timedelta(seconds=90)
        async for p in db.online_presence.find(
            {"project_id": project_id, "last_seen": {"$gte": cutoff}},
            {"user_key": 1, "user_name": 1, "last_seen": 1}
        ):
            online_users.append({
                "user_key": p["user_key"],
                "user_name": p.get("user_name", ""),
                "last_seen": p["last_seen"].isoformat(),
            })

    # 슬라이드 updated_at (MongoDB - 영구 데이터)
    slide_timestamps = {}
    async for s in db.generated_slides.find(
        {"project_id": project_id},
        {"_id": 1, "updated_at": 1}
    ):
        sid = str(s["_id"])
        ts = s.get("updated_at")
        slide_timestamps[sid] = ts.isoformat() if ts else None

    collab_count = await db.collaborators.count_documents({"project_id": project_id})
    slide_count = len(slide_timestamps)

    return {
        "locks": locks,
        "online_users": online_users,
        "slide_timestamps": slide_timestamps,
        "collaborator_count": collab_count,
        "slide_count": slide_count,
        "server_time": now.isoformat(),
    }


# ──────────────────────────────────────────────
# SSE 실시간 협업 스트림
# ──────────────────────────────────────────────

async def _build_init_state(db, project_id: str) -> dict:
    """SSE init 이벤트에 보낼 초기 상태 구성"""
    now = datetime.utcnow()

    # 잠금 조회 (Redis 우선)
    locks = []
    redis_locks = await redis_service.get_project_locks(project_id)
    if redis_locks is not None:
        for rl in redis_locks:
            ttl = rl.get("ttl", 0)
            if ttl > 0:
                expires_at = now + timedelta(seconds=ttl)
                locks.append({
                    "slide_id": rl["slide_id"],
                    "user_key": rl["user_key"],
                    "user_name": rl.get("user_name", ""),
                    "expires_at": expires_at.isoformat(),
                })
    else:
        async for lock in db.slide_locks.find(
            {"project_id": project_id, "expires_at": {"$gt": now}},
            {"slide_id": 1, "user_key": 1, "user_name": 1, "expires_at": 1}
        ):
            locks.append({
                "slide_id": lock["slide_id"],
                "user_key": lock["user_key"],
                "user_name": lock.get("user_name", ""),
                "expires_at": lock["expires_at"].isoformat(),
            })

    # 온라인 사용자 (Redis 우선)
    online_users = []
    redis_users = await redis_service.get_online_users(project_id)
    if redis_users is not None:
        online_users = redis_users
    else:
        cutoff = now - timedelta(seconds=90)
        async for p in db.online_presence.find(
            {"project_id": project_id, "last_seen": {"$gte": cutoff}},
            {"user_key": 1, "user_name": 1, "last_seen": 1}
        ):
            online_users.append({
                "user_key": p["user_key"],
                "user_name": p.get("user_name", ""),
                "last_seen": p["last_seen"].isoformat(),
            })

    # 슬라이드 타임스탬프
    slide_timestamps = {}
    async for s in db.generated_slides.find(
        {"project_id": project_id},
        {"_id": 1, "updated_at": 1}
    ):
        sid = str(s["_id"])
        ts = s.get("updated_at")
        slide_timestamps[sid] = ts.isoformat() if ts else None

    return {
        "locks": locks,
        "online_users": online_users,
        "slide_timestamps": slide_timestamps,
    }


def _sse_event(event: str, data: dict) -> str:
    """SSE 형식 문자열 생성"""
    return f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False, default=str)}\n\n"


@router.get("/{jwt_token}/api/collab/{project_id}/stream")
async def collab_stream(jwt_token: str, project_id: str):
    """SSE 기반 실시간 협업 이벤트 스트림"""
    # 인증 및 권한 확인 (스트림 시작 전 검증)
    user_key = await get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "viewer")

    user = await get_user_by_key(user_key)
    user_name = user.get("nm", user_key) if user else user_key

    async def event_generator():
        # 1. 초기 상태 전송
        init_state = await _build_init_state(db, project_id)
        yield _sse_event("init", init_state)

        # 2. 온라인 상태 갱신
        await redis_service.update_presence(project_id, user_key, user_name, ttl=90)
        await db.online_presence.update_one(
            {"project_id": project_id, "user_key": user_key},
            {"$set": {
                "project_id": project_id,
                "user_key": user_key,
                "user_name": user_name,
                "last_seen": datetime.utcnow(),
            }},
            upsert=True
        )

        # 3. Redis Pub/Sub 구독 시도
        if redis_service.is_available():
            heartbeat_counter = 0
            try:
                async for msg in redis_service.subscribe_collab(project_id):
                    if msg is not None:
                        event_type, data = msg
                        if event_type:
                            yield _sse_event(event_type, data)
                        heartbeat_counter = 0
                    else:
                        # None = timeout (1초마다), heartbeat 카운터 증가
                        heartbeat_counter += 1
                        if heartbeat_counter >= 30:
                            # 30초마다 ping
                            yield _sse_event("ping", {})
                            heartbeat_counter = 0
                            # presence 갱신
                            await redis_service.update_presence(
                                project_id, user_key, user_name, ttl=90
                            )
            except asyncio.CancelledError:
                pass
            finally:
                # 연결 종료: presence 제거 + user_left 이벤트
                await redis_service.remove_presence(project_id, user_key)
                await db.online_presence.delete_one(
                    {"project_id": project_id, "user_key": user_key}
                )
                try:
                    # joined 키 삭제 (다음 접속 시 다시 user_joined 발행되도록)
                    r = redis_service.get_redis()
                    if r:
                        await r.delete(f"{_PREFIX}:joined:{project_id}:{user_key}")
                    await redis_service.publish_collab_event(project_id, "user_left", {
                        "user_key": user_key,
                        "user_name": user_name,
                    })
                except Exception:
                    pass
        else:
            # Redis 불가 → MongoDB 폴링 폴백 (5초 간격)
            heartbeat_counter = 0
            prev_state = init_state
            try:
                while True:
                    await asyncio.sleep(5)
                    heartbeat_counter += 5
                    cur_state = await _build_init_state(db, project_id)

                    # 변경된 부분만 이벤트 발행
                    if cur_state["locks"] != prev_state["locks"]:
                        yield _sse_event("lock_changed", {"locks": cur_state["locks"]})
                    if cur_state["online_users"] != prev_state["online_users"]:
                        yield _sse_event("users_changed", {"online_users": cur_state["online_users"]})
                    if cur_state["slide_timestamps"] != prev_state["slide_timestamps"]:
                        yield _sse_event("slides_changed", {"slide_timestamps": cur_state["slide_timestamps"]})

                    prev_state = cur_state

                    if heartbeat_counter >= 30:
                        yield _sse_event("ping", {})
                        heartbeat_counter = 0

                    # presence 갱신
                    await db.online_presence.update_one(
                        {"project_id": project_id, "user_key": user_key},
                        {"$set": {"last_seen": datetime.utcnow()}},
                    )
            except asyncio.CancelledError:
                pass
            finally:
                await db.online_presence.delete_one(
                    {"project_id": project_id, "user_key": user_key}
                )
                try:
                    await redis_service.publish_collab_event(project_id, "user_left", {
                        "user_key": user_key,
                        "user_name": user_name,
                    })
                except Exception:
                    pass

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        },
    )


# ──────────────────────────────────────────────
# 변경 이력
# ──────────────────────────────────────────────

@router.get("/{jwt_token}/api/projects/{project_id}/history")
async def get_history(jwt_token: str, project_id: str, limit: int = 50):
    """변경 이력 조회"""
    user_key = await get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "viewer")

    cursor = db.slide_history.find(
        {"project_id": project_id}
    ).sort("created_at", -1).limit(limit)

    history = []
    async for h in cursor:
        h["_id"] = str(h["_id"])
        # 스냅샷은 목록에서 제외 (용량 절약)
        h.pop("before_snapshot", None)
        h.pop("after_snapshot", None)
        history.append(h)
    return {"history": history}


@router.post("/{jwt_token}/api/projects/{project_id}/history/{history_id}/revert")
async def revert_history(jwt_token: str, project_id: str, history_id: str):
    """히스토리 되돌리기 (소유자 전용)"""
    user_key = await get_user_key(jwt_token)
    db = get_db()
    await check_project_access(db, project_id, user_key, "owner")

    entry = await db.slide_history.find_one({"_id": ObjectId(history_id)})
    if not entry:
        raise HTTPException(status_code=404, detail="히스토리를 찾을 수 없습니다")
    if entry.get("project_id") != project_id:
        raise HTTPException(status_code=403, detail="권한이 없습니다")

    snapshot = entry.get("before_snapshot")
    if not snapshot:
        raise HTTPException(status_code=400, detail="되돌릴 스냅샷이 없습니다")

    slide_id = entry.get("slide_id")
    now = datetime.utcnow()

    # Lock 확인 (Lock된 슬라이드는 되돌리기 불가)
    lock = await db.slide_locks.find_one(
        {"project_id": project_id, "slide_id": slide_id}
    )
    if lock and lock.get("expires_at") and lock["expires_at"] > now:
        raise HTTPException(
            status_code=423,
            detail=f"{lock.get('user_name', '다른 사용자')}이(가) 편집 중이므로 되돌릴 수 없습니다"
        )

    # 현재 상태 스냅샷 (되돌리기 자체도 이력에 기록)
    current = await db.generated_slides.find_one({"_id": ObjectId(slide_id)})
    if current:
        user = await get_user_by_key(user_key)
        user_name = user.get("nm", user_key) if user else user_key

        await db.slide_history.insert_one({
            "project_id": project_id,
            "slide_id": slide_id,
            "action": "revert",
            "user_key": user_key,
            "user_name": user_name,
            "before_snapshot": {
                "objects": current.get("objects", []),
                "items": current.get("items", []),
            },
            "after_snapshot": snapshot,
            "description": f"히스토리 되돌리기 (#{history_id[:8]})",
            "created_at": now,
        })

    # 스냅샷 적용
    await db.generated_slides.update_one(
        {"_id": ObjectId(slide_id)},
        {"$set": {
            "objects": snapshot.get("objects", []),
            "items": snapshot.get("items", []),
            "updated_at": now,
        }}
    )

    return {"success": True, "slide_id": slide_id}
