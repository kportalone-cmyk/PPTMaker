from motor.motor_asyncio import AsyncIOMotorClient
from config import settings

_client = None


def get_client() -> AsyncIOMotorClient:
    global _client
    if _client is None:
        _client = AsyncIOMotorClient(settings.MONGO_URI)
    return _client


def get_db():
    """앱 DB 반환"""
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
    await db.templates.create_index("is_published")

    # slides 컬렉션
    await db.slides.create_index("template_id")
    await db.slides.create_index([("template_id", 1), ("order", 1)])

    # accounts 컬렉션
    await db.accounts.create_index("user_key", unique=True)

    # demo_accounts 컬렉션 (데모/체험 계정 - 조직도 외 별도 관리)
    await db.demo_accounts.create_index("user_key", unique=True)
    await db.demo_accounts.create_index("nm")

    # projects 컬렉션
    await db.projects.create_index("user_key")
    await db.projects.create_index("created_at")
    await db.projects.create_index("status")
    await db.projects.create_index("project_type")
    await db.projects.create_index("share_token", sparse=True)
    await db.projects.create_index("name")
    await db.projects.create_index([("user_key", 1), ("created_at", -1)])
    await db.projects.create_index([("user_key", 1), ("updated_at", -1)])
    await db.projects.create_index([("status", 1), ("updated_at", -1)])

    # resources 컬렉션
    await db.resources.create_index("project_id")
    await db.resources.create_index("resource_type")
    await db.resources.create_index([("project_id", 1), ("created_at", -1)])

    # generated_slides 컬렉션
    await db.generated_slides.create_index("project_id")
    await db.generated_slides.create_index([("project_id", 1), ("order", 1)])

    # generated_excel 컬렉션
    await db.generated_excel.create_index("project_id", unique=True)

    # onlyoffice_documents 컬렉션
    await db.onlyoffice_documents.create_index("project_id", unique=True)
    await db.onlyoffice_documents.create_index("document_key")

    # docx_templates 컬렉션
    await db.docx_templates.create_index("project_id")
    await db.docx_templates.create_index("created_at")

    # generated_docx 컬렉션
    await db.generated_docx.create_index("project_id", unique=True)

    # project_folders 컬렉션
    await db.project_folders.create_index([("user_key", 1), ("order", 1)])
    await db.project_folders.create_index([("user_key", 1), ("name", 1)], unique=True)

    # file_text_cache 컬렉션 (SHA-256 해시 기반 파일 텍스트 캐시)
    await db.file_text_cache.create_index("file_hash", unique=True)

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
    await db.slide_history.create_index([("slide_id", 1), ("created_at", -1)])

    # online_presence 컬렉션
    await db.online_presence.create_index(
        "last_seen", expireAfterSeconds=60
    )
    await db.online_presence.create_index(
        [("project_id", 1), ("user_key", 1)], unique=True
    )
    # collab-status에서 project_id + last_seen 필터링용
    await db.online_presence.create_index([("project_id", 1), ("last_seen", -1)])

    # custom_templates 컬렉션
    await db.custom_templates.create_index("project_id", unique=True)

    # html_skills 컬렉션
    await db.html_skills.create_index("is_published")
    await db.html_skills.create_index("created_at")

    # generated_html 컬렉션
    await db.generated_html.create_index("project_id", unique=True)

    # slide_styles 컬렉션
    await db.slide_styles.create_index("style_id", unique=True)
    await db.slide_styles.create_index("category")
    await db.slide_styles.create_index("is_active")

    # 조직도 인덱스
    org_db = get_org_db()
    org_col = org_db[settings.ORG_COLLECTION]
    await org_col.create_index("nm")
    await org_col.create_index("ky", unique=True)
    await org_col.create_index("em")
    await org_col.create_index("dp")


async def seed_demo_accounts():
    """데모 계정 시드 (demo_accounts 컬렉션에 없으면 생성)"""
    from services.auth_service import hash_password
    db = get_db()

    demo_user = await db.demo_accounts.find_one({"user_key": "demo"})
    new_hash = hash_password("kmslab1234")
    if not demo_user:
        await db.demo_accounts.insert_one({
            "user_key": "demo",
            "nm": "Demo",
            "dp": "Demo",
            "em": "demo@example.com",
            "ky": "demo",
            "role": "",
            "password": new_hash,
        })
    else:
        # 패스워드 갱신
        await db.demo_accounts.update_one(
            {"user_key": "demo"},
            {"$set": {"password": new_hash}},
        )

    # 기존 accounts 컬렉션에서 Demo 계정 정리 (마이그레이션)
    old_account = await db.accounts.find_one({"user_key": "demo"})
    if old_account:
        await db.accounts.delete_one({"user_key": "demo"})


async def seed_slide_styles():
    """슬라이드 스타일 시드 데이터 초기화 (없으면 삽입, 있으면 업데이트)"""
    from services.slide_styles_seed import SLIDE_STYLES
    db = get_db()

    existing_count = await db.slide_styles.count_documents({})
    if existing_count >= len(SLIDE_STYLES):
        return  # 이미 충분히 있으면 스킵

    from datetime import datetime
    now = datetime.utcnow()
    for style in SLIDE_STYLES:
        await db.slide_styles.update_one(
            {"style_id": style["style_id"]},
            {"$setOnInsert": {**style, "created_at": now}},
            upsert=True,
        )
    print(f"[DB] 슬라이드 스타일 {len(SLIDE_STYLES)}개 시드 완료")


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
