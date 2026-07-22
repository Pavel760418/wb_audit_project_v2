"""
Сборка итогового Excel-файла аудита кабинета WB.
Собирает лист "Дашборд" со сводкой доступности данных и профильные листы
с реальными статусами (полный/частичный/нет данных).
"""
from __future__ import annotations

from openpyxl import Workbook

from modules.data_map import build_summary_table, build_block_status_map
from modules.excel_builder import write_sheet_status, write_availability_table, autosize_columns
from modules import analytics


def build_report(load_results: dict, period_start: str | None, period_end: str | None) -> Workbook:
    wb = Workbook()

    dashboard_ws = wb.active
    dashboard_ws.title = "Дашборд"
    dashboard_ws["A1"] = "Аудит кабинета Wildberries"
    dashboard_ws["A1"].font = dashboard_ws["A1"].font.copy(bold=True, size=14)
    period_label = f"Период: {period_start or 'не указан'} - {period_end or 'не указан'}"
    dashboard_ws["A2"] = period_label

    summary_rows = build_summary_table(load_results)
    dashboard_ws["A4"] = "Состав предоставленных данных"
    dashboard_ws["A4"].font = dashboard_ws["A4"].font.copy(bold=True, size=12)
    write_availability_table(dashboard_ws, summary_rows, start_row=6)
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
    actions_ws["A1"] = "План действий формируется на основе реально рассчитанных блоков"
    autosize_columns(actions_ws)

    instructions_ws = wb.create_sheet("Инструкция")
    instructions_ws["A1"] = "Обязателен: еженедельный финансовый отчёт. Остальные отчёты — по мере доступности."
    autosize_columns(instructions_ws)

    return wb
