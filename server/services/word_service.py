"""
DOCX 생성 서비스 - python-docx 기반 Word 문서 생성

generated_docx 컬렉션의 데이터를 .docx 파일로 변환합니다.
웹 프리뷰(CKEditor/스트리밍)와 최대한 동일한 스타일을 적용합니다.
"""

import os
import re
import json
import uuid
import tempfile
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from services.mongo_service import get_db
from config import settings

# matplotlib 차트 렌더링
try:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import numpy as np
    _HAS_MATPLOTLIB = True
except ImportError:
    _HAS_MATPLOTLIB = False

# 웹 프리뷰와 동일한 색상 팔레트
_COLORS = {
    "h1": RGBColor(0x1A, 0x1A, 0x2E),
    "h2": RGBColor(0x1E, 0x29, 0x3B),
    "h3": RGBColor(0x33, 0x41, 0x55),
    "h4": RGBColor(0x47, 0x55, 0x69),
    "body": RGBColor(0x33, 0x41, 0x55),
    "description": RGBColor(0x64, 0x74, 0x8B),
    "blockquote": RGBColor(0x43, 0x38, 0xCA),
    "table_header_text": RGBColor(0x1E, 0x29, 0x3B),
    "table_body_text": RGBColor(0x33, 0x41, 0x55),
    "accent": RGBColor(0x63, 0x66, 0xF1),
}

# 제목 폰트 크기 (px→pt 변환: CKEditor 14px 기준)
_HEADING_SIZES = {
    0: Pt(22),  # 문서 타이틀
    1: Pt(20),  # h1: 26px → ~20pt
    2: Pt(15),  # h2: 20px → ~15pt
    3: Pt(13),  # h3: 17px → ~13pt
    4: Pt(11),  # h4: 15px → ~11pt
}

# 제목 색상
_HEADING_COLORS = {
    0: _COLORS["h1"],
    1: _COLORS["h1"],
    2: _COLORS["h2"],
    3: _COLORS["h3"],
    4: _COLORS["h4"],
}


def _set_cell_shading(cell, color_hex: str):
    """테이블 셀 배경색 설정"""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def _set_paragraph_spacing(paragraph, line_spacing=1.6, space_before=0, space_after=Pt(6)):
    """단락 줄간격 및 전후 간격 설정"""
    pf = paragraph.paragraph_format
    pf.line_spacing = line_spacing
    if space_before is not None:
        pf.space_before = space_before
    if space_after is not None:
        pf.space_after = space_after


def _set_run_font(run, font_name="맑은 고딕", size=None, color=None, bold=None, italic=None):
    """run에 폰트 속성 설정"""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    if size:
        run.font.size = size
    if color:
        run.font.color.rgb = color
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def _add_styled_heading(doc, text: str, level: int):
    """웹 프리뷰 스타일과 일치하는 제목 추가"""
    level = min(level, 4)
    p = doc.add_paragraph()

    # 제목 텍스트 run
    run = p.add_run(text)
    _set_run_font(
        run,
        size=_HEADING_SIZES.get(level, Pt(11)),
        color=_HEADING_COLORS.get(level, _COLORS["body"]),
        bold=True,
    )

    # 줄간격 설정
    pf = p.paragraph_format
    pf.line_spacing = 1.4

    if level == 0:
        # 문서 타이틀: 가운데 정렬, 아래 여백
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.space_before = Pt(4)
        pf.space_after = Pt(20)
        # 하단 보더 (accent 색상)
        p_element = p._element
        pPr = p_element.get_or_add_pPr()
        borders = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:bottom w:val="single" w:sz="12" w:space="4" w:color="6366F1"/>'
            f'</w:pBdr>'
        )
        pPr.append(borders)

    elif level == 1:
        # h1: 큰 여백
        pf.space_before = Pt(36)
        pf.space_after = Pt(14)
        # 하단 보더
        p_element = p._element
        pPr = p_element.get_or_add_pPr()
        borders = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:bottom w:val="single" w:sz="12" w:space="4" w:color="6366F1"/>'
            f'</w:pBdr>'
        )
        pPr.append(borders)

    elif level == 2:
        # h2: 좌측 보더 + 배경
        pf.space_before = Pt(30)
        pf.space_after = Pt(12)
        pf.left_indent = Pt(10)
        # 좌측 보더
        p_element = p._element
        pPr = p_element.get_or_add_pPr()
        borders = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:left w:val="single" w:sz="18" w:space="8" w:color="6366F1"/>'
            f'</w:pBdr>'
        )
        pPr.append(borders)
        # 배경 shading
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F0FF" w:val="clear"/>')
        pPr.append(shading)

    elif level == 3:
        # h3: 얇은 좌측 보더
        pf.space_before = Pt(24)
        pf.space_after = Pt(10)
        pf.left_indent = Pt(8)
        p_element = p._element
        pPr = p_element.get_or_add_pPr()
        borders = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:left w:val="single" w:sz="12" w:space="6" w:color="A5B4FC"/>'
            f'</w:pBdr>'
        )
        pPr.append(borders)

    elif level == 4:
        # h4: 간단한 여백만
        pf.space_before = Pt(20)
        pf.space_after = Pt(8)

    return p


# ── 차트 렌더링 ──────────────────────────────────────────

# 차트 색상 팔레트 (웹 프리뷰와 동일)
_CHART_PALETTE = [
    "#6366F1", "#EC4899", "#F59E0B", "#10B981",
    "#3B82F6", "#8B5CF6", "#EF4444", "#14B8A6",
]


def _is_chart_json(data: dict) -> bool:
    """dict가 Chart.js 설정인지 확인"""
    return (
        isinstance(data, dict)
        and "type" in data
        and "data" in data
        and isinstance(data.get("data"), dict)
        and "labels" in data["data"]
        and "datasets" in data["data"]
    )


def _render_chart_image(chart_config: dict, output_path: str) -> bool:
    """Chart.js JSON 설정을 matplotlib 이미지로 변환"""
    if not _HAS_MATPLOTLIB:
        return False

    try:
        # 한글 폰트 설정
        plt.rcParams['font.family'] = 'Malgun Gothic'
        plt.rcParams['axes.unicode_minus'] = False

        chart_type = chart_config.get("type", "bar")
        data = chart_config["data"]
        labels = data.get("labels", [])
        datasets = data.get("datasets", [])
        options = chart_config.get("options", {})

        if not labels or not datasets:
            return False

        fig, ax = plt.subplots(figsize=(10, 5.5))

        if chart_type == 'bar':
            x = np.arange(len(labels))
            n = len(datasets)
            width = 0.7 / max(n, 1)
            for idx, ds in enumerate(datasets):
                offset = (idx - n / 2 + 0.5) * width
                colors = ds.get("backgroundColor", _CHART_PALETTE[idx % len(_CHART_PALETTE)])
                ax.bar(x + offset, ds["data"], width,
                       label=ds.get("label", ""), color=colors,
                       edgecolor='white', linewidth=0.5)
            ax.set_xticks(x)
            max_label_len = max((len(str(l)) for l in labels), default=0)
            rotation = 30 if max_label_len > 6 else 0
            ha = 'right' if rotation else 'center'
            ax.set_xticklabels(labels, rotation=rotation, ha=ha, fontsize=10)

        elif chart_type == 'line':
            for idx, ds in enumerate(datasets):
                color = ds.get("borderColor", _CHART_PALETTE[idx % len(_CHART_PALETTE)])
                ax.plot(range(len(labels)), ds["data"],
                        label=ds.get("label", ""), marker='o',
                        linewidth=2.5, color=color, markersize=6)
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels, fontsize=10)

        elif chart_type in ('pie', 'doughnut'):
            ds = datasets[0]
            colors = ds.get("backgroundColor", _CHART_PALETTE[:len(labels)])
            if chart_type == 'doughnut':
                wedges, texts, autotexts = ax.pie(
                    ds["data"], labels=labels, autopct='%1.1f%%',
                    colors=colors, pctdistance=0.75,
                    textprops={'fontsize': 10})
                centre = plt.Circle((0, 0), 0.50, fc='white')
                ax.add_patch(centre)
            else:
                ax.pie(ds["data"], labels=labels, autopct='%1.1f%%',
                       colors=colors, textprops={'fontsize': 10})
            ax.set_aspect('equal')

        elif chart_type == 'area':
            for idx, ds in enumerate(datasets):
                color = ds.get("backgroundColor", _CHART_PALETTE[idx % len(_CHART_PALETTE)])
                ax.fill_between(range(len(labels)), ds["data"],
                                alpha=0.4, color=color, label=ds.get("label", ""))
                ax.plot(range(len(labels)), ds["data"], linewidth=2, color=color)
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels, fontsize=10)

        else:
            # 기본 bar
            x = np.arange(len(labels))
            for idx, ds in enumerate(datasets):
                colors = ds.get("backgroundColor", _CHART_PALETTE[idx % len(_CHART_PALETTE)])
                ax.bar(x, ds["data"], label=ds.get("label", ""), color=colors)
            ax.set_xticks(x)
            ax.set_xticklabels(labels, fontsize=10)

        # 제목
        title = options.get("title", "")
        if isinstance(title, str) and title:
            ax.set_title(title, fontsize=14, fontweight='bold', pad=15)
        elif isinstance(title, dict):
            ax.set_title(title.get("text", title.get("display", "")),
                         fontsize=14, fontweight='bold', pad=15)

        # 축 레이블
        scales = options.get("scales", {})
        if isinstance(scales, dict):
            y_cfg = scales.get("y", {})
            if isinstance(y_cfg, dict):
                y_title = y_cfg.get("title", "")
                if isinstance(y_title, str) and y_title:
                    ax.set_ylabel(y_title, fontsize=11)
                elif isinstance(y_title, dict):
                    ax.set_ylabel(y_title.get("text", ""), fontsize=11)
            x_cfg = scales.get("x", {})
            if isinstance(x_cfg, dict):
                x_title = x_cfg.get("title", "")
                if isinstance(x_title, str) and x_title:
                    ax.set_xlabel(x_title, fontsize=11)
                elif isinstance(x_title, dict):
                    ax.set_xlabel(x_title.get("text", ""), fontsize=11)

        # 범례 및 그리드
        if chart_type not in ('pie', 'doughnut'):
            if any(ds.get("label") for ds in datasets):
                ax.legend(fontsize=10, framealpha=0.9)
            ax.grid(axis='y', alpha=0.3, linestyle='--')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)

        plt.tight_layout()
        fig.savefig(output_path, dpi=150, bbox_inches='tight', facecolor='white')
        plt.close(fig)
        return True

    except Exception as e:
        print(f"[WordService] Chart render error: {e}")
        try:
            plt.close('all')
        except:
            pass
        return False


def _add_chart_to_doc(doc: Document, chart_config: dict):
    """Chart.js 설정을 이미지로 렌더링하여 문서에 삽입"""
    tmp_path = os.path.join(tempfile.gettempdir(), f"chart_{uuid.uuid4().hex}.png")
    try:
        if _render_chart_image(chart_config, tmp_path):
            doc.add_picture(tmp_path, width=Cm(15))
            last_p = doc.paragraphs[-1]
            last_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_paragraph_spacing(last_p, space_before=Pt(10), space_after=Pt(14))
    finally:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except:
            pass


def _add_code_block(doc: Document, code_text: str):
    """코드 블록을 배경색 있는 단락으로 렌더링"""
    p = doc.add_paragraph()
    run = p.add_run(code_text)
    _set_run_font(run, font_name="Consolas", size=Pt(9), color=_COLORS["body"])
    _set_paragraph_spacing(p, line_spacing=1.3, space_before=Pt(6), space_after=Pt(6))
    # 배경색
    p_element = p._element
    pPr = p_element.get_or_add_pPr()
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F1F5F9" w:val="clear"/>')
    pPr.append(shading)


async def generate_docx(project_id: str) -> str:
    """generated_docx 데이터를 .docx 파일로 변환"""
    db = get_db()
    docx_data = await db.generated_docx.find_one({"project_id": project_id})
    if not docx_data:
        raise ValueError("생성된 문서 데이터가 없습니다")

    sections = docx_data.get("sections", [])
    if not sections:
        raise ValueError("섹션 데이터가 없습니다")

    meta = docx_data.get("meta", {})
    doc = Document()

    # 기본 스타일 설정 - 웹 프리뷰와 동일한 폰트/크기
    style = doc.styles["Normal"]
    font = style.font
    font.name = "맑은 고딕"
    font.size = Pt(11)
    font.color.rgb = _COLORS["body"]
    # eastAsia 폰트 설정
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    # 기본 줄간격/단락간격
    pf = style.paragraph_format
    pf.line_spacing = 1.6
    pf.space_after = Pt(6)

    # 문서 제목
    title = meta.get("title", "")
    if title:
        _add_styled_heading(doc, title, level=0)

    # 문서 설명
    description = meta.get("description", "")
    if description:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(description)
        _set_run_font(run, color=_COLORS["description"], size=Pt(10), italic=True)
        _set_paragraph_spacing(p, space_before=0, space_after=Pt(4))
        # 구분선
        hr = doc.add_paragraph()
        hr_pf = hr.paragraph_format
        hr_pf.space_before = Pt(8)
        hr_pf.space_after = Pt(12)
        hr_element = hr._element
        hrPr = hr_element.get_or_add_pPr()
        hr_border = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:bottom w:val="single" w:sz="4" w:space="1" w:color="CBD5E1"/>'
            f'</w:pBdr>'
        )
        hrPr.append(hr_border)

    # 섹션 처리
    for section in sections:
        section_title = section.get("title", "")
        level = section.get("level", 1)
        content = section.get("content", "")

        # 제목 추가
        if section_title:
            _add_styled_heading(doc, section_title, level=min(level, 4))

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

        # 코드 블록 (```...```) - 차트 JSON 감지 포함
        if stripped.startswith("```"):
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            if i < len(lines):
                i += 1  # skip closing ```

            code_text = "\n".join(code_lines).strip()
            if code_text:
                # Chart.js JSON인지 확인
                try:
                    parsed = json.loads(code_text)
                    if _is_chart_json(parsed):
                        _add_chart_to_doc(doc, parsed)
                        continue
                except (json.JSONDecodeError, ValueError):
                    pass
                # 일반 코드 블록 렌더링
                _add_code_block(doc, code_text)
            continue

        # Raw JSON 차트 감지 (``` 없이 { 로 시작하는 JSON)
        if stripped == "{":
            json_lines = [lines[i]]
            brace_count = stripped.count("{") - stripped.count("}")
            j = i + 1
            while j < len(lines) and brace_count > 0:
                json_lines.append(lines[j])
                brace_count += lines[j].count("{") - lines[j].count("}")
                j += 1
            json_text = "\n".join(json_lines).strip()
            try:
                parsed = json.loads(json_text)
                if _is_chart_json(parsed):
                    _add_chart_to_doc(doc, parsed)
                    i = j
                    continue
            except (json.JSONDecodeError, ValueError):
                pass
            # 차트가 아니면 일반 텍스트로 처리 (fall through)

        # 마크다운 헤딩 (# ~ ####)
        heading_match = re.match(r"^(#{1,4})\s+(.+)$", stripped)
        if heading_match:
            level = len(heading_match.group(1))
            heading_text = heading_match.group(2)
            _add_styled_heading(doc, heading_text, level=level)
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
            _set_paragraph_spacing(p, space_before=Pt(2), space_after=Pt(2))
            i += 1
            continue

        # 번호 리스트
        num_match = re.match(r"^\d+\.\s+(.+)$", stripped)
        if num_match:
            text = num_match.group(1)
            p = doc.add_paragraph(style="List Number")
            _add_inline_formatting(p, text)
            _set_paragraph_spacing(p, space_before=Pt(2), space_after=Pt(2))
            i += 1
            continue

        # 인용문 (웹 프리뷰 인디고 테마와 동일)
        if stripped.startswith("> "):
            text = stripped[2:]
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.left_indent = Cm(1)
            _set_paragraph_spacing(p, line_spacing=1.5, space_before=Pt(6), space_after=Pt(6))
            # 좌측 보더 (accent 색상)
            p_element = p._element
            pPr = p_element.get_or_add_pPr()
            borders = parse_xml(
                f'<w:pBdr {nsdecls("w")}>'
                f'  <w:left w:val="single" w:sz="18" w:space="8" w:color="6366F1"/>'
                f'</w:pBdr>'
            )
            pPr.append(borders)
            # 배경 shading
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F0FF" w:val="clear"/>')
            pPr.append(shading)
            # 텍스트
            run = p.add_run(text)
            _set_run_font(run, color=_COLORS["blockquote"], size=Pt(10))
            i += 1
            continue

        # 일반 텍스트
        p = doc.add_paragraph()
        _add_inline_formatting(p, stripped)
        _set_paragraph_spacing(p, space_before=Pt(3), space_after=Pt(6))
        i += 1


def _add_inline_formatting(paragraph, text: str):
    """인라인 마크다운 서식 처리 (**굵게**, *기울임*, ~~취소선~~, `코드`)"""
    pattern = re.compile(r"(\*\*(.+?)\*\*|\*(.+?)\*|~~(.+?)~~|`([^`]+)`)")

    last_end = 0
    for match in pattern.finditer(text):
        # 매치 앞 텍스트
        if match.start() > last_end:
            run = paragraph.add_run(text[last_end:match.start()])
            run.font.color.rgb = _COLORS["body"]

        if match.group(2):  # **bold**
            run = paragraph.add_run(match.group(2))
            run.bold = True
            run.font.color.rgb = _COLORS["h2"]
        elif match.group(3):  # *italic*
            run = paragraph.add_run(match.group(3))
            run.italic = True
            run.font.color.rgb = _COLORS["h4"]
        elif match.group(4):  # ~~strikethrough~~
            run = paragraph.add_run(match.group(4))
            run.font.strike = True
        elif match.group(5):  # `code`
            run = paragraph.add_run(match.group(5))
            run.font.color.rgb = _COLORS["accent"]
            run.font.size = Pt(10)
            # 코드 배경 shading
            rPr = run._element.get_or_add_rPr()
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F1F5F9" w:val="clear"/>')
            rPr.append(shading)

        last_end = match.end()

    # 나머지 텍스트
    if last_end < len(text):
        run = paragraph.add_run(text[last_end:])
        run.font.color.rgb = _COLORS["body"]


def _add_table(doc: Document, table_lines: list):
    """마크다운 테이블을 웹 프리뷰 스타일과 동일한 docx 테이블로 변환"""
    if len(table_lines) < 2:
        return

    # 구분선(---|---) 행 제거
    data_lines = []
    for line in table_lines:
        cells = [c.strip() for c in line.strip("|").split("|")]
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
        row = table.rows[row_idx]
        for col_idx, cell_text in enumerate(row_data):
            if col_idx < num_cols:
                cell = table.cell(row_idx, col_idx)
                # 기존 텍스트 클리어 후 run으로 추가
                cell.text = ""
                p = cell.paragraphs[0]
                run = p.add_run(cell_text.strip())
                run.font.name = "맑은 고딕"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")

                # 셀 패딩
                _set_paragraph_spacing(p, line_spacing=1.4, space_before=Pt(3), space_after=Pt(3))

                if row_idx == 0:
                    # 헤더 행: 굵게, 배경색, 텍스트 색상
                    run.bold = True
                    run.font.size = Pt(10)
                    run.font.color.rgb = _COLORS["table_header_text"]
                    _set_cell_shading(cell, "F1F5F9")
                    # 헤더 하단 보더 강조
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    borders = parse_xml(
                        f'<w:tcBorders {nsdecls("w")}>'
                        f'  <w:bottom w:val="single" w:sz="8" w:space="0" w:color="6366F1"/>'
                        f'</w:tcBorders>'
                    )
                    tcPr.append(borders)
                else:
                    # 데이터 행
                    run.font.size = Pt(10)
                    run.font.color.rgb = _COLORS["table_body_text"]
                    # 짝수 행 배경색 (교차 행, 0-based에서 row_idx 짝수 = 데이터 홀수행)
                    if row_idx % 2 == 0:
                        _set_cell_shading(cell, "F8FAFC")

    # 테이블 후 여백
    spacing_p = doc.add_paragraph("")
    spacing_p.paragraph_format.space_before = Pt(4)
    spacing_p.paragraph_format.space_after = Pt(4)
