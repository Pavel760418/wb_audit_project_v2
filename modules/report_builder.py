"""
Сборка итогового Excel-файла аудита кабинета Wildberries.

Лист «Дашборд» содержит сводку: период, состав предоставленных данных,
диагностику недельных финансовых отчётов и ограничения анализа. Профильные
листы получают явный статус полноты (полный / частичный / нет данных).

Внешний вид и состав книги сохранены: светофор статусов, названия листов и
аккуратное форматирование не ухудшаются — блоки только дополняются.
"""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Font

from modules import analytics
from modules.data_map import (
    build_block_limitations,
    build_block_status_map,
    build_summary_table,
)
from modules.excel_builder import (
    STATUS_LABELS,
    autosize_columns,
    write_availability_table,
    write_sheet_status,
)


def _write_finance_files_section(ws, load_results, start_row):
    """Пишет диагностику недельных финансовых отчётов, если они загружались."""
    finance = load_results.get("finance_weekly")
    if not finance or not getattr(finance, "file_details", None):
        return start_row

    header_cell = ws.cell(row=start_row, column=1, value="Загруженные недельные финансовые отчёты")
    header_cell.font = Font(bold=True, size=12)

    summary_cell = ws.cell(
        row=start_row + 1,
        column=1,
        value=(
            f"Всего файлов: {finance.files_total} · "
            f"успешно: {finance.files_ok} · "
            f"с предупреждениями: {finance.files_warning} · "
            f"не вошло в расчёт: {finance.files_failed}"
        ),
    )
    summary_cell.font = Font(italic=True)

    headers = ["Файл", "Статус", "Строк", "Комментарий"]
    header_row = start_row + 2
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = Font(bold=True)

    for offset, detail in enumerate(finance.file_details, start=1):
        r = header_row + offset
        ws.cell(row=r, column=1, value=detail.get("name", "—"))
        ws.cell(row=r, column=2, value=STATUS_LABELS.get(detail.get("status", ""), detail.get("status", "")))
        ws.cell(row=r, column=3, value=detail.get("rows", 0))
        ws.cell(row=r, column=4, value=detail.get("message", ""))

    return header_row + len(finance.file_details) + 2


def _write_limitations_section(ws, load_results, start_row):
    """Пишет блок «Ограничения анализа» — какие блоки посчитаны не полностью и почему."""
    limitations = build_block_limitations(load_results)

    header_cell = ws.cell(row=start_row, column=1, value="Ограничения анализа")
    header_cell.font = Font(bold=True, size=12)

    if not limitations:
        ws.cell(
            row=start_row + 1,
            column=1,
            value="Ограничений нет: все аналитические блоки рассчитаны в полном объёме.",
        )
        return start_row + 3

    headers = ["Блок анализа", "Статус", "Каких данных не хватает"]
    header_row = start_row + 1
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = Font(bold=True)

    for offset, item in enumerate(limitations, start=1):
        r = header_row + offset
        ws.cell(row=r, column=1, value=item["block"])
        ws.cell(row=r, column=2, value=STATUS_LABELS.get(item["status"], item["status"]))
        missing = ", ".join(item["missing"]) if item["missing"] else "—"
        ws.cell(row=r, column=3, value=missing)

    return header_row + len(limitations) + 2


def build_report(load_results: dict, period_start: str | None, period_end: str | None) -> Workbook:
    wb = Workbook()

    dashboard_ws = wb.active
    dashboard_ws.title = "Дашборд"
    dashboard_ws["A1"] = "Аудит кабинета Wildberries"
    dashboard_ws["A1"].font = Font(bold=True, size=14)
    period_label = f"Период анализа: {period_start or 'не указан'} — {period_end or 'не указан'}"
    dashboard_ws["A2"] = period_label

    finance = load_results.get("finance_weekly")
    if finance and getattr(finance, "files_total", 0):
        dashboard_ws["A3"] = (
            f"Недельных финансовых отчётов в расчёте: {finance.files_ok + finance.files_warning} "
            f"из {finance.files_total}"
        )

    summary_rows = build_summary_table(load_results)
    dashboard_ws["A5"] = "Состав предоставленных данных"
    dashboard_ws["A5"].font = Font(bold=True, size=12)
    write_availability_table(dashboard_ws, summary_rows, start_row=7)

    next_row = 7 + len(summary_rows) + 3
    next_row = _write_finance_files_section(dashboard_ws, load_results, next_row)
    _write_limitations_section(dashboard_ws, load_results, next_row)
    autosize_columns(dashboard_ws)

    block_status = build_block_status_map(load_results)

    sales_ws = wb.create_sheet("Продажи")
    sales_summary = analytics.calc_sales_summary(load_results)
    write_sheet_status(sales_ws, block_status.get("sales", "none"))
    sales_ws["A3"] = str(sales_summary)
    autosize_columns(sales_ws)

    finance_ws = wb.create_sheet("Финансы")
    finance_summary = analytics.calc_unit_economics_cabinet(load_results)
    write_sheet_status(finance_ws, block_status.get("finance", "none"))
    finance_ws["A3"] = str(finance_summary)
    autosize_columns(finance_ws)

    unit_sku_ws = wb.create_sheet("Юнит-экономика SKU")
    unit_sku_summary = analytics.calc_unit_economics_sku(load_results)
    write_sheet_status(unit_sku_ws, block_status.get("unit_economics_sku", "none"))
    unit_sku_ws["A3"] = str(unit_sku_summary)
    autosize_columns(unit_sku_ws)

    unit_cabinet_ws = wb.create_sheet("Юнит-экономика кабинета")
    unit_cabinet_summary = analytics.calc_unit_economics_cabinet(load_results)
    write_sheet_status(unit_cabinet_ws, block_status.get("unit_economics_cabinet", "none"))
    unit_cabinet_ws["A3"] = str(unit_cabinet_summary)
    autosize_columns(unit_cabinet_ws)

    stocks_ws = wb.create_sheet("Остатки и доступность")
    stock_risks = analytics.calc_stock_risks(load_results)
    write_sheet_status(stocks_ws, block_status.get("stocks", "none"))
    stocks_ws["A3"] = str(stock_risks)
    autosize_columns(stocks_ws)

    risks_ws = wb.create_sheet("Риски и проблемы")
    write_sheet_status(risks_ws, block_status.get("risks", "none"))
    autosize_columns(risks_ws)

    actions_ws = wb.create_sheet("План действий")
    actions_ws["A1"] = "План действий формируется на основе реально рассчитанных блоков."
    autosize_columns(actions_ws)

    instructions_ws = wb.create_sheet("Инструкция")
    instructions_ws["A1"] = (
        "Обязателен еженедельный финансовый отчёт (можно загрузить несколько недель). "
        "Остальные отчёты — по мере доступности."
    )
    autosize_columns(instructions_ws)

    return wb
