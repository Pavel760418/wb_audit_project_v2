"""
CLI-вход для запуска аудита кабинета WB локально в PyCharm.
Поэтапно запрашивает пути к Excel-отчётам (Enter — пропустить) и период анализа.
"""
from __future__ import annotations

import os
from datetime import datetime

from config import REPORTS_REGISTRY
from modules.loader import load_all_reports
from modules.report_builder import build_report


def run_audit_pipeline(file_paths: dict, period_start: str | None, period_end: str | None):
    """
    Единая pipeline-функция. Используется и из main.py, и из streamlit_app.py,
    чтобы CLI и веб-версия считали строго одинаково.
    """
    load_results = load_all_reports(file_paths)
    workbook = build_report(load_results, period_start, period_end)
    return workbook, load_results


def ask_file_path(report_key: str, meta: dict) -> str | None:
    prompt = f"\n[{ 'ОБЯЗАТЕЛЕН' if meta['required'] else 'опционально' }] {meta['title']}\nГде искать: {meta['where']}\nВведите путь к файлу или нажмите Enter, чтобы пропустить: "
    path = input(prompt).strip()
    if not path:
        return None
    if not os.path.isfile(path):
        print(f"Файл не найден по пути: {path}. Пропускаю этот отчёт.")
        return None
    return path


def main():
    print("=== Аудит кабинета Wildberries ===")
    file_paths = {}
    for report_key, meta in REPORTS_REGISTRY.items():
        file_paths[report_key] = ask_file_path(report_key, meta)

    period_start = input("\nНачало периода анализа (ГГГГ-ММ-ДД) или Enter, чтобы пропустить: ").strip() or None
    period_end = input("Конец периода анализа (ГГГГ-ММ-ДД) или Enter, чтобы пропустить: ").strip() or None

    workbook, _ = run_audit_pipeline(file_paths, period_start, period_end)

    os.makedirs("output", exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join("output", f"WB_Audit_Report_{timestamp}.xlsx")
    workbook.save(output_path)

    print(f"\nГотово. Итоговый файл сохранён: {output_path}")


if __name__ == "__main__":
    main()
