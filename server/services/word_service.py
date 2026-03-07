"""
DOCX 생성 서비스 - python-docx 기반 Word 문서 생성

generated_docx 컬렉션의 데이터를 .docx 파일로 변환합니다.
"""

import os
import re
import uuid
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from services.mongo_service import get_db
from config import settings


async def generate_docx(project_id: str) -> str:
    """generated_docx 데이터를 .docx 파일로 변환

    Returns:
        생성된 파일의 상대 경로 (예: /uploads/documents/xxx.docx)
    """
    db = get_db()
    docx_data = await db.generated_docx.find_one({"project_id": project_id})
    if not docx_data:
        raise ValueError("생성된 문서 데이터가 없습니다")

    sections = docx_data.get("sections", [])
    if not sections:
        raise ValueError("섹션 데이터가 없습니다")

    meta = docx_data.get("meta", {})
    doc = Document()

    # 기본 스타일 설정
    style = doc.styles["Normal"]
    font = style.font
    font.name = "맑은 고딕"
    font.size = Pt(11)

    # 문서 제목
    title = meta.get("title", "")
    if title:
        heading = doc.add_heading(title, level=0)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 문서 설명
    description = meta.get("description", "")
    if description:
        p = doc.add_paragraph(description)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.style.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        doc.add_paragraph("")  # 빈 줄

    # 섹션 처리
    for section in sections:
        section_title = section.get("title", "")
        level = section.get("level", 1)
        content = section.get("content", "")

        # 제목 추가
        if section_title:
            doc.add_heading(section_title, level=min(level, 4))

        # 본문 처리 (Markdown → docx)
        if content:
            _render_markdown_content(doc, content)

    # 파일 저장
    output_dir = os.path.join(settings.UPLOAD_DIR, "documents")
    os.makedirs(output_dir, exist_ok=True)
    filename = f"{uuid.uuid4().hex}.docx"
    output_path = os.path.join(output_dir, filename)
    doc.save(output_path)

    return f"/uploads/documents/{filename}"


def _render_markdown_content(doc: Document, content: str):
    """마크다운 텍스트를 docx 요소로 변환"""
    lines = content.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # 빈 줄
        if not stripped:
            i += 1
            continue

        # 테이블 감지 (| col1 | col2 |)
        if stripped.startswith("|") and "|" in stripped[1:]:
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            _add_table(doc, table_lines)
            continue

        # 불릿 리스트
        if stripped.startswith("- ") or stripped.startswith("* "):
            text = stripped[2:]
            p = doc.add_paragraph(style="List Bullet")
            _add_inline_formatting(p, text)
            i += 1
            continue

        # 번호 리스트
        num_match = re.match(r"^\d+\.\s+(.+)$", stripped)
        if num_match:
            text = num_match.group(1)
            p = doc.add_paragraph(style="List Number")
            _add_inline_formatting(p, text)
            i += 1
            continue

        # 인용문
        if stripped.startswith("> "):
            text = stripped[2:]
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1)
            run = p.add_run(text)
            run.italic = True
            run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
            i += 1
            continue

        # 일반 텍스트
        p = doc.add_paragraph()
        _add_inline_formatting(p, stripped)
        i += 1


def _add_inline_formatting(paragraph, text: str):
    """인라인 마크다운 서식 처리 (**굵게**, *기울임*, ~~취소선~~)"""
    # 패턴: **bold**, *italic*, ~~strikethrough~~
    pattern = re.compile(r"(\*\*(.+?)\*\*|\*(.+?)\*|~~(.+?)~~)")

    last_end = 0
    for match in pattern.finditer(text):
        # 매치 앞 텍스트
        if match.start() > last_end:
            paragraph.add_run(text[last_end:match.start()])

        if match.group(2):  # **bold**
            run = paragraph.add_run(match.group(2))
            run.bold = True
        elif match.group(3):  # *italic*
            run = paragraph.add_run(match.group(3))
            run.italic = True
        elif match.group(4):  # ~~strikethrough~~
            run = paragraph.add_run(match.group(4))
            run.font.strike = True

        last_end = match.end()

    # 나머지 텍스트
    if last_end < len(text):
        paragraph.add_run(text[last_end:])


def _add_table(doc: Document, table_lines: list):
    """마크다운 테이블을 docx 테이블로 변환"""
    if len(table_lines) < 2:
        return

    # 구분선(---|---) 행 제거
    data_lines = []
    for line in table_lines:
        cells = [c.strip() for c in line.strip("|").split("|")]
        # 구분선 체크
        if all(re.match(r"^[-:]+$", c) for c in cells if c):
            continue
        data_lines.append(cells)

    if not data_lines:
        return

    num_cols = max(len(row) for row in data_lines)
    num_rows = len(data_lines)

    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for row_idx, row_data in enumerate(data_lines):
        for col_idx, cell_text in enumerate(row_data):
            if col_idx < num_cols:
                cell = table.cell(row_idx, col_idx)
                cell.text = cell_text.strip()
                # 헤더 행 굵게
                if row_idx == 0:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.bold = True

    doc.add_paragraph("")  # 테이블 후 빈 줄
