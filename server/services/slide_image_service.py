"""
LibreOffice headless + PyMuPDF 를 이용한 PPTX → 슬라이드별 PNG 변환 서비스.

설치 의존성:
  - LibreOffice (soffice/soffice.exe) — 시스템에 설치되어 있어야 함.
    Windows 기본 경로: C:\\Program Files\\LibreOffice\\program\\soffice.exe
    환경변수 LIBREOFFICE_PATH 로 override 가능.
  - PyMuPDF — pip install PyMuPDF (자급자족 Python 패키지, 별도 시스템 바이너리 불필요).

설치되어 있지 않거나 변환 실패 시 빈 리스트를 반환하므로, 호출 측은 "이미지 미리보기는
선택적 기능 — 실패해도 .pptx 다운로드는 정상 동작" 으로 처리해야 한다.
"""

import asyncio
import os
import shutil
import sys
from pathlib import Path
from typing import Optional

from config import settings


def _find_soffice() -> Optional[str]:
    """LibreOffice 바이너리 경로 탐색.

    1) settings.LIBREOFFICE_PATH (환경변수)
    2) Windows 기본 설치 경로
    3) Linux/Mac 기본 명령 (PATH 검색)
    """
    env_path = getattr(settings, "LIBREOFFICE_PATH", None)
    if env_path:
        if os.path.isfile(env_path):
            return env_path
        print(f"[SlideImage] LIBREOFFICE_PATH 가 지정되었으나 파일이 없음: {env_path}")

    if sys.platform == "win32":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for c in candidates:
            if os.path.isfile(c):
                return c

    found = shutil.which("soffice") or shutil.which("soffice.exe") or shutil.which("libreoffice")
    return found


async def convert_pptx_to_images(
    pptx_path: str,
    out_dir: str,
    dpi: int = 220,
    timeout: int = 240,
) -> list[str]:
    """PPTX 를 슬라이드별 PNG 로 변환.

    1) LibreOffice headless 로 PPTX → PDF
    2) PyMuPDF 로 PDF → 페이지별 PNG (slide_001.png, slide_002.png, ...)

    Args:
      pptx_path: 입력 .pptx 절대 경로
      out_dir:   출력 디렉토리 (없으면 생성). PNG 들은 여기에 저장됨.
      dpi:       PNG 해상도 (72=원본, 144=2배, 220=권장 기본 — HD/Retina 디스플레이 대응,
                 300=프린트 품질). 10" × 5.625" 슬라이드 기준 픽셀 크기:
                 144 → 1440×810, 200 → 2000×1125, 220 → 2200×1238, 300 → 3000×1688.
                 settings.SLIDE_IMAGE_DPI 로 환경변수 override 가능.
      timeout:   soffice 변환 타임아웃 (초)

    Returns:
      생성된 PNG 파일들의 절대 경로 리스트. 변환 실패 시 빈 리스트.
    """
    pptx = Path(pptx_path)
    if not pptx.is_file():
        print(f"[SlideImage] PPTX 파일 없음: {pptx_path}")
        return []

    soffice = _find_soffice()
    if not soffice:
        print("[SlideImage] LibreOffice (soffice) 가 설치되지 않았거나 PATH 에 없습니다. "
              "LIBREOFFICE_PATH 환경변수로 경로를 지정하거나 LibreOffice 를 설치하세요.")
        return []

    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    # UNO 사용자 프로필을 작업 디렉토리에 격리 — 같은 머신에서 LibreOffice 다른
    # 프로세스와 락 충돌을 막는다.
    user_profile = out / "_uno_profile"
    user_profile.mkdir(exist_ok=True)
    profile_uri = user_profile.absolute().as_uri()

    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--nofirststartwizard",
        "--nologo",
        f"-env:UserInstallation={profile_uri}",
        "--convert-to", "pdf",
        "--outdir", str(out.absolute()),
        str(pptx.absolute()),
    ]

    try:
        proc = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        try:
            stdout, stderr = await asyncio.wait_for(proc.communicate(), timeout=timeout)
        except asyncio.TimeoutError:
            try:
                proc.kill()
            except Exception:
                pass
            print(f"[SlideImage] soffice 변환 타임아웃 ({timeout}s)")
            return []

        if proc.returncode != 0:
            err = (stderr or b"").decode("utf-8", "ignore")[:300]
            print(f"[SlideImage] soffice 변환 실패 (exit {proc.returncode}): {err}")
            return []
    except FileNotFoundError:
        print(f"[SlideImage] soffice 바이너리를 실행할 수 없음: {soffice}")
        return []
    except Exception as e:
        print(f"[SlideImage] soffice 실행 실패: {e}")
        return []

    pdf_path = out / (pptx.stem + ".pdf")
    if not pdf_path.is_file():
        print(f"[SlideImage] PDF 생성 결과를 찾을 수 없음: {pdf_path}")
        return []

    try:
        import fitz  # PyMuPDF
    except ImportError:
        print("[SlideImage] PyMuPDF 가 설치되지 않았습니다. "
              "requirements.txt 의 PyMuPDF 를 설치하세요: pip install PyMuPDF")
        return []

    image_paths: list[str] = []
    try:
        doc = fitz.open(str(pdf_path))
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img_path = out / f"slide_{i:03d}.png"
            pix.save(str(img_path))
            image_paths.append(str(img_path))
        doc.close()
    except Exception as e:
        print(f"[SlideImage] PDF→PNG 변환 실패: {e}")
        return []

    try:
        pdf_path.unlink()
    except Exception:
        pass
    try:
        shutil.rmtree(user_profile, ignore_errors=True)
    except Exception:
        pass

    print(f"[SlideImage] {len(image_paths)} 슬라이드 이미지 변환 완료 → {out}")
    return image_paths


def to_relative_urls(image_paths: list[str]) -> list[str]:
    """절대 파일 경로를 /uploads/... 형태의 정적 서빙 URL 로 변환.

    image_paths 가 settings.UPLOAD_DIR 하위에 있을 때만 변환되며, 그 외 경로는
    파일명만 잘라 /uploads/ 뒤에 붙인다 (안전한 폴백).
    """
    upload_root = os.path.abspath(settings.UPLOAD_DIR)
    out = []
    for p in image_paths:
        abs_p = os.path.abspath(p)
        if abs_p.lower().startswith(upload_root.lower()):
            rel = abs_p[len(upload_root):].replace("\\", "/").lstrip("/")
            out.append(f"/uploads/{rel}")
        else:
            out.append("/uploads/" + os.path.basename(abs_p))
    return out
