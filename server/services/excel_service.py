"""
XLSX 생성 서비스 - openpyxl 기반 스프레드시트 생성

generated_excel 컬렉션의 데이터를 .xlsx 파일로 변환합니다.
차트는 matplotlib로 PNG 이미지를 생성하여 시트에 삽입합니다.
"""

import os
import uuid
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XlImage
from services.mongo_service import get_db
from bson import ObjectId
from config import settings

# matplotlib 서버 환경 설정 (GUI 불필요)
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np


async def generate_xlsx(project_id: str) -> str:
    """generated_excel 데이터를 .xlsx 파일로 변환

    Returns:
        생성된 파일의 상대 경로 (예: /uploads/generated/xxx.xlsx)
    """
    db = get_db()
    excel_data = await db.generated_excel.find_one({"project_id": project_id})
    if not excel_data:
        raise ValueError("생성된 엑셀 데이터가 없습니다")

    wb = Workbook()
    # 기본 시트 제거
    wb.remove(wb.active)

    sheets = excel_data.get("sheets", [])
    if not sheets:
        raise ValueError("시트 데이터가 없습니다")

    # 스타일 정의
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_alignment = Alignment(vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    alt_fill = PatternFill(start_color="F2F7FB", end_color="F2F7FB", fill_type="solid")

    for sheet_data in sheets:
        sheet_name = sheet_data.get("name", "Sheet")[:31]  # Excel 시트명 최대 31자
        ws = wb.create_sheet(title=sheet_name)

        columns = sheet_data.get("columns", [])
        rows = sheet_data.get("rows", [])

        if not columns and not rows:
            continue

        # 헤더 행
        for col_idx, col_name in enumerate(columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border

        # 데이터 행
        for row_idx, row_data in enumerate(rows, 2):
            for col_idx, cell_value in enumerate(row_data, 1):
                if col_idx > len(columns):
                    break
                # 숫자 타입 변환 시도
                value = _try_convert_value(cell_value)
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.alignment = cell_alignment
                cell.border = thin_border
                # 짝수 행 배경색
                if row_idx % 2 == 0:
                    cell.fill = alt_fill

        # 자동 열 너비 계산
        for col_idx in range(1, len(columns) + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row_idx in range(1, len(rows) + 2):
                cell_val = ws.cell(row=row_idx, column=col_idx).value
                if cell_val is not None:
                    # 한글 문자는 약 2배 너비
                    val_str = str(cell_val)
                    char_len = sum(2 if ord(c) > 127 else 1 for c in val_str)
                    max_len = max(max_len, char_len)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 10), 60)

        # 헤더 행 고정
        ws.freeze_panes = "A2"

        # 차트 이미지 추가
        _add_chart_images_to_sheet(ws, sheet_data)

    # 파일 저장
    output_dir = os.path.join(settings.UPLOAD_DIR, "generated")
    os.makedirs(output_dir, exist_ok=True)
    filename = f"{uuid.uuid4().hex}.xlsx"
    output_path = os.path.join(output_dir, filename)
    wb.save(output_path)

    return f"/uploads/generated/{filename}"


def _try_convert_value(value):
    """문자열을 적절한 타입으로 변환 시도"""
    if value is None or value == "":
        return value
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        # 정수 시도
        try:
            return int(value)
        except (ValueError, TypeError):
            pass
        # 실수 시도
        try:
            return float(value)
        except (ValueError, TypeError):
            pass
        # 퍼센트 (예: "85%")
        if value.endswith("%"):
            try:
                return float(value[:-1]) / 100
            except (ValueError, TypeError):
                pass
    return value


# ============ 차트 자동 생성 ============

def auto_generate_chart_definition(sheet_data: dict, chart_type: str = "bar", title: str = None) -> dict | None:
    """시트 데이터에서 자동으로 차트 정의 생성 (LLM 없이)

    Args:
        sheet_data: {"name", "columns", "rows"} 형식의 시트 데이터
        chart_type: bar, line, pie, area, scatter, doughnut, radar
        title: 차트 제목 (None이면 자동 생성)

    Returns:
        차트 정의 dict 또는 None (데이터 부족 시)
    """
    columns = sheet_data.get("columns", [])
    rows = sheet_data.get("rows", [])

    if not columns or not rows:
        return None

    # 라벨 컬럼: 첫 번째 컬럼
    labels_column = 0

    # 숫자 데이터가 있는 컬럼을 시리즈로 선택
    series = []
    for i in range(len(columns)):
        if i == labels_column:
            continue
        # 해당 컬럼에 숫자 데이터가 있는지 확인
        numeric_count = 0
        for row in rows:
            if i < len(row):
                val = row[i]
                if isinstance(val, (int, float)):
                    numeric_count += 1
                elif isinstance(val, str):
                    try:
                        float(val.replace(",", ""))
                        numeric_count += 1
                    except (ValueError, TypeError):
                        pass
        if numeric_count > 0:
            series.append({"name": columns[i], "column": i})

    if not series:
        return None

    # pie/doughnut은 시리즈 1개만
    if chart_type in ("pie", "doughnut"):
        series = series[:1]

    sheet_name = sheet_data.get("name", "Sheet")
    chart_title = title or sheet_name

    return {
        "type": chart_type,
        "title": chart_title,
        "data_range": {
            "labels_column": labels_column,
            "series": series,
            "row_start": 0,
            "row_end": len(rows) - 1,
        },
        "options": {
            "stacked": False,
            "show_legend": len(series) > 1,
        },
    }


def render_chart_to_file(chart_def: dict, sheet_data: dict) -> str | None:
    """차트를 matplotlib로 이미지 렌더링하여 파일로 저장

    Returns:
        이미지 상대 경로 (예: /uploads/charts/xxx.png) 또는 None
    """
    rows = sheet_data.get("rows", [])
    columns = sheet_data.get("columns", [])

    img_bytes = _render_chart_image(chart_def, rows, columns)
    if not img_bytes:
        return None

    output_dir = os.path.join(settings.UPLOAD_DIR, "charts")
    os.makedirs(output_dir, exist_ok=True)
    filename = f"{uuid.uuid4().hex}.png"
    filepath = os.path.join(output_dir, filename)

    with open(filepath, "wb") as f:
        f.write(img_bytes.getvalue())

    return f"/uploads/charts/{filename}"


# ============ 차트 이미지 생성 (matplotlib) ============

# 색상 팔레트 (Chart.js 프론트엔드와 동일)
_CHART_COLORS = [
    "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
    "#5B9BD5", "#70AD47", "#264478", "#9B59B6",
]


def _add_chart_images_to_sheet(ws, sheet_data: dict):
    """matplotlib로 차트 PNG를 생성하여 시트에 이미지로 삽입"""
    charts = sheet_data.get("charts", [])
    if not charts:
        return

    rows = sheet_data.get("rows", [])
    columns = sheet_data.get("columns", [])
    num_data_rows = len(rows)
    num_cols = len(columns)

    if num_data_rows == 0 or num_cols == 0:
        return

    for chart_idx, chart_def in enumerate(charts[:3]):
        chart_type = chart_def.get("type", "bar")
        if chart_type not in ("bar", "line", "pie", "area", "scatter", "doughnut", "radar"):
            continue

        try:
            img_bytes = _render_chart_image(chart_def, rows, columns)
            if img_bytes:
                img = XlImage(img_bytes)
                # 앵커: 데이터 아래 3행, 차트 간 20행 간격
                anchor_row = num_data_rows + 3 + (chart_idx * 20)
                ws.add_image(img, f"A{anchor_row}")
        except Exception as e:
            print(f"[Excel] 차트 이미지 생성 실패 (chart_idx={chart_idx}): {e}")
            continue


def _render_chart_image(chart_def: dict, rows: list, columns: list) -> BytesIO | None:
    """단일 차트 정의 → matplotlib PNG BytesIO"""
    chart_type = chart_def.get("type", "bar")
    title = chart_def.get("title", "")
    data_range = chart_def.get("data_range", {})
    series_defs = data_range.get("series", [])
    labels_col = data_range.get("labels_column", 0)
    row_start = data_range.get("row_start", 0)
    row_end = data_range.get("row_end")
    options = chart_def.get("options", {})

    num_rows = len(rows)
    num_cols = len(columns)
    if row_end is None:
        row_end = num_rows - 1
    row_end = min(row_end, num_rows - 1)

    if row_start > row_end or not series_defs:
        return None

    # 라벨 추출
    labels = []
    for i in range(row_start, row_end + 1):
        row = rows[i] if i < num_rows else []
        val = row[labels_col] if labels_col < len(row) else ""
        labels.append(str(val) if val is not None else "")

    # 시리즈 데이터 추출
    series_list = []
    for s_def in series_defs:
        col_idx = s_def.get("column", 1)
        if col_idx < 0 or col_idx >= num_cols:
            continue
        name = s_def.get("name", columns[col_idx] if col_idx < len(columns) else f"Series {col_idx}")
        values = []
        for i in range(row_start, row_end + 1):
            row = rows[i] if i < num_rows else []
            val = row[col_idx] if col_idx < len(row) else 0
            try:
                values.append(float(val))
            except (ValueError, TypeError):
                values.append(0)
        series_list.append({"name": name, "values": values})

    if not series_list:
        return None

    # 한글 폰트 설정
    plt.rcParams["font.family"] = "Malgun Gothic"
    plt.rcParams["axes.unicode_minus"] = False

    # 차트 타입별 렌더링
    if chart_type in ("pie", "doughnut"):
        fig, ax = plt.subplots(figsize=(8, 6))
        _render_pie(ax, labels, series_list[0]["values"], title, chart_type == "doughnut")
    elif chart_type == "radar":
        fig, ax = plt.subplots(figsize=(8, 6), subplot_kw={"projection": "polar"})
        _render_radar(ax, labels, series_list, title, options)
    else:
        fig, ax = plt.subplots(figsize=(10, 6))
        if chart_type == "bar":
            _render_bar(ax, labels, series_list, title, options)
        elif chart_type == "line":
            _render_line(ax, labels, series_list, title, options)
        elif chart_type == "area":
            _render_area(ax, labels, series_list, title, options)
        elif chart_type == "scatter":
            _render_scatter(ax, labels, series_list, title, options)

    fig.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


# ---- 차트 유형별 렌더링 ----

def _render_bar(ax, labels, series_list, title, options):
    """세로 막대 차트"""
    x = np.arange(len(labels))
    n = len(series_list)
    width = 0.7 / max(n, 1)
    stacked = options.get("stacked", False)

    for i, s in enumerate(series_list):
        color = _CHART_COLORS[i % len(_CHART_COLORS)]
        if stacked:
            bottom = np.zeros(len(labels))
            for j in range(i):
                bottom += np.array(series_list[j]["values"])
            ax.bar(x, s["values"], width=0.7, bottom=bottom, label=s["name"], color=color)
        else:
            offset = (i - n / 2 + 0.5) * width
            ax.bar(x + offset, s["values"], width=width, label=s["name"], color=color)

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45 if len(labels) > 6 else 0, ha="right" if len(labels) > 6 else "center")
    ax.set_title(title, fontsize=14, fontweight="bold", pad=12)
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    if len(series_list) > 1 or options.get("show_legend", True):
        ax.legend(loc="best")
    ax.grid(axis="y", alpha=0.3)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def _render_line(ax, labels, series_list, title, options):
    """꺾은선 차트"""
    x = np.arange(len(labels))
    for i, s in enumerate(series_list):
        color = _CHART_COLORS[i % len(_CHART_COLORS)]
        ax.plot(x, s["values"], marker="o", markersize=5, label=s["name"], color=color, linewidth=2)

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45 if len(labels) > 6 else 0, ha="right" if len(labels) > 6 else "center")
    ax.set_title(title, fontsize=14, fontweight="bold", pad=12)
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    if len(series_list) > 1 or options.get("show_legend", True):
        ax.legend(loc="best")
    ax.grid(alpha=0.3)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def _render_area(ax, labels, series_list, title, options):
    """영역 차트"""
    x = np.arange(len(labels))
    for i, s in enumerate(series_list):
        color = _CHART_COLORS[i % len(_CHART_COLORS)]
        ax.fill_between(x, s["values"], alpha=0.4, label=s["name"], color=color)
        ax.plot(x, s["values"], color=color, linewidth=1.5)

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45 if len(labels) > 6 else 0, ha="right" if len(labels) > 6 else "center")
    ax.set_title(title, fontsize=14, fontweight="bold", pad=12)
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    if len(series_list) > 1 or options.get("show_legend", True):
        ax.legend(loc="best")
    ax.grid(alpha=0.3)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def _render_scatter(ax, labels, series_list, title, options):
    """산점도"""
    try:
        x_values = [float(v) for v in labels]
    except (ValueError, TypeError):
        x_values = list(range(len(labels)))

    for i, s in enumerate(series_list):
        color = _CHART_COLORS[i % len(_CHART_COLORS)]
        ax.scatter(x_values, s["values"], label=s["name"], color=color, s=60, alpha=0.8)

    ax.set_title(title, fontsize=14, fontweight="bold", pad=12)
    ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    if len(series_list) > 1 or options.get("show_legend", True):
        ax.legend(loc="best")
    ax.grid(alpha=0.3)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def _render_pie(ax, labels, values, title, is_doughnut=False):
    """파이/도넛 차트"""
    colors = [_CHART_COLORS[i % len(_CHART_COLORS)] for i in range(len(labels))]
    wedges, texts, autotexts = ax.pie(
        values,
        labels=labels,
        autopct="%1.1f%%",
        colors=colors,
        startangle=90,
        pctdistance=0.75 if is_doughnut else 0.6,
    )
    for t in autotexts:
        t.set_fontsize(9)

    if is_doughnut:
        centre_circle = plt.Circle((0, 0), 0.50, fc="white")
        ax.add_artist(centre_circle)

    ax.set_title(title, fontsize=14, fontweight="bold", pad=12)


def _render_radar(ax, labels, series_list, title, options):
    """레이더 차트"""
    n = len(labels)
    if n < 3:
        return
    angles = np.linspace(0, 2 * np.pi, n, endpoint=False).tolist()
    angles += angles[:1]  # 닫기

    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_thetagrids(np.degrees(angles[:-1]), labels)

    for i, s in enumerate(series_list):
        color = _CHART_COLORS[i % len(_CHART_COLORS)]
        vals = s["values"] + s["values"][:1]
        ax.plot(angles, vals, linewidth=2, label=s["name"], color=color)
        ax.fill(angles, vals, alpha=0.15, color=color)

    ax.set_title(title, fontsize=14, fontweight="bold", pad=20)
    if len(series_list) > 1 or options.get("show_legend", True):
        ax.legend(loc="upper right", bbox_to_anchor=(1.3, 1.1))
