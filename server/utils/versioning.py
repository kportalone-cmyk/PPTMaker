import os
import time
import hashlib


def get_file_version(file_path: str) -> str:
    """파일의 수정 시간 기반 버전 해시 생성"""
    try:
        mtime = os.path.getmtime(file_path)
        return hashlib.md5(str(mtime).encode()).hexdigest()[:8]
    except OSError:
        return hashlib.md5(str(time.time()).encode()).hexdigest()[:8]


def versioned_url(url: str, base_dir: str) -> str:
    """URL에 버전 파라미터 추가"""
    file_path = os.path.join(base_dir, url.lstrip("/"))
    version = get_file_version(file_path)
    separator = "&" if "?" in url else "?"
    return f"{url}{separator}v={version}"
