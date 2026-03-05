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

    # fonts 컬렉션
    await db.fonts.create_index("name", unique=True)

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
