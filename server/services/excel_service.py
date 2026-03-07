"""
XLSX 생성 서비스 - openpyxl 기반 스프레드시트 생성

generated_excel 컬렉션의 데이터를 .xlsx 파일로 변환합니다.
"""

import os
import uuid
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import (
    BarChart, LineChart, PieChart, AreaChart,
    ScatterChart, DoughnutChart, RadarChart, Reference,
)
from services.mongo_service import get_db
from bson import ObjectId
from config import settings


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

        # 차트 추가
        _add_charts_to_sheet(ws, sheet_data)

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


# ============ 차트 생성 ============

_CHART_CLASS_MAP = {
    "bar": BarChart,
    "line": LineChart,
    "pie": PieChart,
    "area": AreaChart,
    "scatter": ScatterChart,
    "doughnut": DoughnutChart,
    "radar": RadarChart,
}


def _add_charts_to_sheet(ws, sheet_data: dict):
    """시트에 차트 추가 (openpyxl)"""
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
        ChartClass = _CHART_CLASS_MAP.get(chart_type)
        if not ChartClass:
            continue

        chart = ChartClass()
        chart.title = chart_def.get("title", "")
        chart.width = 18
        chart.height = 10

        data_range = chart_def.get("data_range", {})
        series_defs = data_range.get("series", [])
        labels_col = data_range.get("labels_column", 0)
        row_start = data_range.get("row_start", 0)
        row_end = data_range.get("row_end")
        if row_end is None:
            row_end = num_data_rows - 1

        # openpyxl Reference: 1-based, 헤더=row1, 데이터=row2~
        min_row = row_start + 2
        max_row = row_end + 2

        if min_row > max_row or max_row > num_data_rows + 1:
            continue

        # 카테고리 라벨
        cats = Reference(
            ws,
            min_col=labels_col + 1,
            min_row=min_row,
            max_row=max_row,
        )

        # scatter 차트: X/Y 참조 방식 별도 처리
        if chart_type == "scatter":
            from openpyxl.chart import Series as ChartSeries
            x_ref = Reference(ws, min_col=labels_col + 1, min_row=min_row, max_row=max_row)
            for s_def in series_defs:
                col_idx = s_def.get("column", 1)
                if col_idx < 0 or col_idx >= num_cols:
                    continue
                y_ref = Reference(ws, min_col=col_idx + 1, min_row=min_row, max_row=max_row)
                series = ChartSeries(y_ref, xvalues=x_ref, title=s_def.get("name", ""))
                chart.series.append(series)
        else:
            for s_def in series_defs:
                col_idx = s_def.get("column", 1)
                if col_idx < 0 or col_idx >= num_cols:
                    continue
                data_ref = Reference(
                    ws,
                    min_col=col_idx + 1,
                    min_row=min_row - 1,  # 헤더 행 포함 (titles_from_data)
                    max_row=max_row,
                )
                chart.add_data(data_ref, titles_from_data=True)

            chart.set_categories(cats)

        # 차트 옵션
        options = chart_def.get("options", {})
        if options.get("show_legend") is False:
            chart.legend = None

        if chart_type == "bar" and options.get("stacked"):
            chart.grouping = "stacked"

        if chart_type == "area":
            chart.grouping = "standard"

        chart.style = 10

        # 앵커: 데이터 아래 3행, 차트 간 16행 간격
        anchor_row = num_data_rows + 3 + (chart_idx * 16)
        ws.add_chart(chart, f"A{anchor_row}")
