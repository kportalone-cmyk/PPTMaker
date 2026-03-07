from motor.motor_asyncio import AsyncIOMotorClient
from config import settings

_client = None


def get_client() -> AsyncIOMotorClient:
    global _client
    if _client is None:
        _client = AsyncIOMotorClient(settings.MONGO_URI)
    return _client


def get_db():
    """PPTMaker DB 반환"""
    return get_client()[settings.PPTMAKER_DB]


def get_org_db():
    """조직도 DB 반환"""
    return get_client()[settings.ORG_DB]


async def init_indexes():
    """MongoDB 인덱스 초기화"""
    db = get_db()

    # templates 컬렉션
    await db.templates.create_index("created_at")
    await db.templates.create_index("name")

    # slides 컬렉션
    await db.slides.create_index("template_id")
    await db.slides.create_index([("template_id", 1), ("order", 1)])

    # projects 컬렉션
    await db.projects.create_index("user_key")
    await db.projects.create_index("created_at")

    # resources 컬렉션
    await db.resources.create_index("project_id")

    # generated_slides 컬렉션
    await db.generated_slides.create_index("project_id")
    await db.generated_slides.create_index([("project_id", 1), ("order", 1)])

    # generated_excel 컬렉션
    await db.generated_excel.create_index("project_id", unique=True)

    # onlyoffice_documents 컬렉션
    await db.onlyoffice_documents.create_index("project_id", unique=True)
    await db.onlyoffice_documents.create_index("document_key")

    # generated_docx 컬렉션
    await db.generated_docx.create_index("project_id", unique=True)

    # fonts 컬렉션
    await db.fonts.create_index("name", unique=True)

    # prompts 컬렉션
    await db.prompts.create_index("key", unique=True)

    # collaborators 컬렉션
    await db.collaborators.create_index(
        [("project_id", 1), ("user_key", 1)], unique=True
    )
    await db.collaborators.create_index("user_key")
    await db.collaborators.create_index("project_id")

    # slide_locks 컬렉션
    await db.slide_locks.create_index(
        "expires_at", expireAfterSeconds=0
    )
    await db.slide_locks.create_index(
        [("project_id", 1), ("slide_id", 1)], unique=True
    )
    await db.slide_locks.create_index([("project_id", 1), ("user_key", 1)])
    # collab-status에서 expires_at 필터링용 복합 인덱스
    await db.slide_locks.create_index([("project_id", 1), ("expires_at", 1)])

    # slide_history 컬렉션
    await db.slide_history.create_index([("project_id", 1), ("created_at", -1)])
    await db.slide_history.create_index([("project_id", 1), ("slide_id", 1)])

    # online_presence 컬렉션
    await db.online_presence.create_index(
        "last_seen", expireAfterSeconds=60
    )
    await db.online_presence.create_index(
        [("project_id", 1), ("user_key", 1)], unique=True
    )
    # collab-status에서 project_id + last_seen 필터링용
    await db.online_presence.create_index([("project_id", 1), ("last_seen", -1)])

    # 조직도 인덱스
    org_db = get_org_db()
    org_col = org_db[settings.ORG_COLLECTION]
    await org_col.create_index("nm")
    await org_col.create_index("ky", unique=True)
    await org_col.create_index("em")


async def close_connection():
    global _client
    if _client:
        _client.close()
        _client = None


def close_connection_sync():
    """동기 버전 - 서버 종료 시 CancelledError 상황에서 사용"""
    global _client
    if _client:
        _client.close()
        _client = None
