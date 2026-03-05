from perplexity import AsyncPerplexity
from config import settings


async def search_web(query: str) -> dict:
    """Perplexity Search API를 사용한 웹 검색 (SDK 기반)"""
    if not settings.PERPLEXITY_API_KEY:
        return {
            "pages": [],
            "error": "Perplexity API 키가 설정되지 않았습니다."
        }

    try:
        async with AsyncPerplexity(api_key=settings.PERPLEXITY_API_KEY) as client:
            search = await client.search.create(
                query=query,
                max_results=10,
                max_tokens=50000,
                max_tokens_per_page=4096,
            )

        pages = []
        for result in search.results:
            content = getattr(result, "snippet", "") or getattr(result, "content", "")
            if not content or len(content.strip()) < 50:
                continue
            pages.append({
                "url": getattr(result, "url", ""),
                "title": getattr(result, "title", ""),
                "content": content.strip(),
            })

        return {"pages": pages}

    except Exception as e:
        return {
            "pages": [],
            "error": f"검색 중 오류가 발생했습니다: {str(e)}"
        }
