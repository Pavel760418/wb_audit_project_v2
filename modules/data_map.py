"""
Карта доступности входных данных и статусов аналитических блоков.
Строится на основе результатов loader.py и реестра MIN_DATA_FOR_BLOCK из config.py.
"""
from __future__ import annotations

from config import MIN_DATA_FOR_BLOCK, REPORTS_REGISTRY


def build_data_availability_map(load_results: dict) -> dict:
    """
    Возвращает сводную таблицу по каждому отчёту: статус, заголовок, обязательность.
    """
    availability = {}
    for report_key, meta in REPORTS_REGISTRY.items():
        result = load_results.get(report_key)
        status = result.status if result else "missing"
        availability[report_key] = {
            "title": meta["title"],
            "required": meta["required"],
            "status": status,
            "message": result.message if result else "Отчёт не запрашивался",
        }
    return availability


def build_block_status_map(load_results: dict) -> dict:
    """
    Для каждого аналитического блока определяет статус:
    "full" - все нужные источники загружены успешно,
    "partial" - часть источников доступна,
    "none" - ни одного источника нет.
    """
    block_status = {}
    for block, required_reports in MIN_DATA_FOR_BLOCK.items():
        statuses = [load_results.get(r).status if load_results.get(r) else "missing" for r in required_reports]
        ok_count = sum(1 for s in statuses if s in ("ok", "warning"))
        if ok_count == len(required_reports):
            block_status[block] = "full"
        elif ok_count > 0:
            block_status[block] = "partial"
        else:
            block_status[block] = "none"
    return block_status


def build_summary_table(load_results: dict) -> list:
    """
    Формирует построчную summary-таблицу полноты данных для главной страницы отчёта.
    """
    rows = []
    availability = build_data_availability_map(load_results)
    for report_key, info in availability.items():
        rows.append(
            {
                "Отчёт": info["title"],
                "Обязателен": "Да" if info["required"] else "Нет",
                "Статус": info["status"],
                "Комментарий": info["message"],
            }
        )
    return rows
