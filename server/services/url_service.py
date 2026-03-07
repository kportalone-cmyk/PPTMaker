"""URL 콘텐츠 수집 서비스 - 일반 웹페이지 텍스트 추출 및 YouTube 자막 추출"""

import re
import httpx
from urllib.parse import urlparse, parse_qs
from bs4 import BeautifulSoup


def is_youtube_url(url: str) -> bool:
    """YouTube URL 여부 판별"""
    parsed = urlparse(url)
    host = parsed.hostname or ""
    return any(h in host for h in ["youtube.com", "youtu.be", "youtube-nocookie.com"])


def extract_youtube_id(url: str) -> str | None:
    """YouTube URL에서 비디오 ID 추출"""
    parsed = urlparse(url)
    host = parsed.hostname or ""

    # youtu.be 단축 URL
    if "youtu.be" in host:
        vid = parsed.path.lstrip("/").split("/")[0]
        if vid and len(vid) == 11:
            return vid

    # youtube.com 계열
    if "youtube" in host:
        # /watch?v=ID
        qs = parse_qs(parsed.query)
        if "v" in qs:
            return qs["v"][0]
        # /embed/ID, /v/ID, /shorts/ID
        m = re.match(r"^/(?:embed|v|shorts)/([a-zA-Z0-9_-]{11})", parsed.path)
        if m:
            return m.group(1)

    return None


async def fetch_youtube_subtitles(video_id: str) -> dict:
    """YouTube 자막 추출 (우선순위: ko → en → 자동생성 → 아무거나)"""
    try:
        from youtube_transcript_api import YouTubeTranscriptApi

        ytt_api = YouTubeTranscriptApi()

        # 선호 언어 순서로 시도
        preferred = ["ko", "en", "ja", "zh-Hans", "zh-Hant"]
        try:
            transcript = ytt_api.fetch(video_id, languages=preferred)
            text = " ".join(snippet.text for snippet in transcript)
            return {"title": f"YouTube: {video_id}", "content": text}
        except Exception:
            pass

        # 아무 언어나 시도
        try:
            transcript_list = ytt_api.list(video_id)
            available = transcript_list.transcript_entries
            if available:
                first_lang = available[0].language_code
                transcript = ytt_api.fetch(video_id, languages=[first_lang])
                text = " ".join(snippet.text for snippet in transcript)
                return {"title": f"YouTube: {video_id}", "content": text}
        except Exception:
            pass

        return {"title": f"YouTube: {video_id}", "content": "", "error": "자막을 찾을 수 없습니다"}

    except Exception as e:
        return {"title": f"YouTube: {video_id}", "content": "", "error": str(e)}


async def fetch_youtube_title(video_id: str) -> str:
    """YouTube 비디오 제목을 oEmbed API로 가져오기"""
    try:
        async with httpx.AsyncClient(timeout=10) as client:
            resp = await client.get(
                f"https://www.youtube.com/oembed?url=https://www.youtube.com/watch?v={video_id}&format=json"
            )
            if resp.status_code == 200:
                data = resp.json()
                return data.get("title", "")
    except Exception:
        pass
    return ""


async def fetch_url_content(url: str) -> dict:
    """일반 웹페이지의 텍스트 콘텐츠 추출"""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        }
        async with httpx.AsyncClient(timeout=30, follow_redirects=True, max_redirects=5) as client:
            resp = await client.get(url, headers=headers)
            resp.raise_for_status()

        soup = BeautifulSoup(resp.text, "html.parser")

        # 제목 추출
        title = ""
        if soup.title and soup.title.string:
            title = soup.title.string.strip()

        # 불필요한 태그 제거
        for tag in soup(["script", "style", "nav", "footer", "header", "aside", "noscript", "iframe"]):
            tag.decompose()

        # 본문 텍스트 추출 (article 우선, 없으면 body)
        article = soup.find("article") or soup.find("main") or soup.find("body")
        if article:
            text = article.get_text(separator="\n", strip=True)
        else:
            text = soup.get_text(separator="\n", strip=True)

        # 과도한 빈 줄 정리
        lines = [line.strip() for line in text.splitlines()]
        lines = [line for line in lines if line]
        content = "\n".join(lines)

        # 너무 긴 콘텐츠 제한 (약 50KB)
        if len(content) > 50000:
            content = content[:50000] + "\n\n...(이하 생략)"

        return {"title": title or url, "content": content}

    except Exception as e:
        return {"title": url, "content": "", "error": str(e)}


async def process_url(url: str) -> dict:
    """URL을 분석하여 콘텐츠 수집 (YouTube/일반 URL 자동 분기)"""
    url = url.strip()
    if not url:
        return None

    # http:// 또는 https:// 없으면 추가
    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    if is_youtube_url(url):
        video_id = extract_youtube_id(url)
        if not video_id:
            return {"title": url, "content": "", "source_url": url, "resource_type": "youtube", "error": "유효하지 않은 YouTube URL입니다"}

        result = await fetch_youtube_subtitles(video_id)
        # 제목 가져오기
        yt_title = await fetch_youtube_title(video_id)
        if yt_title:
            result["title"] = yt_title

        result["source_url"] = url
        result["resource_type"] = "youtube"
        return result
    else:
        result = await fetch_url_content(url)
        result["source_url"] = url
        result["resource_type"] = "url"
        return result
