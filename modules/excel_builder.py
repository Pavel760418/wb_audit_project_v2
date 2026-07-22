"""
Низкоуровневые примитивы сборки итогового Excel-файла:
светофор, статусы блоков, условное форматирование, базовые таблицы.
"""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

STATUS_COLORS = {
    "ok": "C6EFCE",
    "full": "C6EFCE",
    "warning": "FFEB9C",
    "partial": "FFEB9C",
    "missing": "FFC7CE",
    "error": "FFC7CE",
    "none": "FFC7CE",
    "insufficient_data": "FFC7CE",
}

STATUS_LABELS = {
    "ok": "Полный расчёт",
    "full": "Полный расчёт",
    "warning": "Частичный расчёт",
    "partial": "Частичный расчёт",
    "missing": "Нет данных для расчёта",
    "error": "Нет данных для расчёта",
    "none": "Нет данных для расчёта",
    "insufficient_data": "Нет данных для расчёта",
}


def write_sheet_status(ws, status: str, cell: str = "A1"):
    label = STATUS_LABELS.get(status, "Нет данных для расчёта")
    color = STATUS_COLORS.get(status, "FFC7CE")
    ws[cell] = f"Статус листа: {label}"
    ws[cell].font = Font(bold=True, size=12)
    ws[cell].fill = PatternFill(start_color=color, end_color=color, fill_type="solid")


def write_availability_table(ws, summary_rows: list, start_row: int = 3):
    headers = ["Отчёт", "Обязателен", "Статус", "Комментарий"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=col_idx, value=header)
        cell.font = Font(bold=True)

    for row_offset, row in enumerate(summary_rows, start=1):
        r = start_row + row_offset
        ws.cell(row=r, column=1, value=row["Отчёт"])
        ws.cell(row=r, column=2, value=row["Обязателен"])
        status_cell = ws.cell(row=r, column=3, value=STATUS_LABELS.get(row["Статус"], row["Статус"]))
        color = STATUS_COLORS.get(row["Статус"], "FFC7CE")
        status_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        ws.cell(row=r, column=4, value=row["Комментарий"])

    for col_idx in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 32


def autosize_columns(ws, max_width: int = 60):
    for column_cells in ws.columns:
        length = max((len(str(cell.value)) if cell.value is not None else 0) for cell in column_cells)
        col_letter = get_column_letter(column_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max_width, max(12, length + 2))
