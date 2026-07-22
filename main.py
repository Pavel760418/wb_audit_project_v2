"""
CLI-вход для запуска аудита кабинета Wildberries локально (в том числе в PyCharm).

Поэтапно запрашивает пути к Excel-отчётам (Enter — пропустить) и период анализа.
Еженедельный финансовый отчёт можно указать несколькими файлами (пути через
запятую) — они будут объединены в единый анализируемый период.
"""
from __future__ import annotations

import os
from datetime import datetime

from config import REPORTS_REGISTRY
from modules.loader import load_all_reports
from modules.report_builder import build_report


def run_audit_pipeline(file_paths: dict, period_start: str | None, period_end: str | None):
    """Единая pipeline-функция для CLI и веб-версии.

    Значение в file_paths может быть None, одним файлом/путём или списком файлов.
    Так CLI (main.py) и Streamlit (streamlit_app.py) считают строго одинаково.
    """
    load_results = load_all_reports(file_paths)
    workbook = build_report(load_results, period_start, period_end)
    return workbook, load_results


def _ask_single_path(meta: dict) -> str | None:
    prompt = (
        f"\n[{'ОБЯЗАТЕЛЕН' if meta['required'] else 'опционально'}] {meta['title']}\n"
        f"Где взять: {meta['where']}\n"
        "Введите путь к файлу или нажмите Enter, чтобы пропустить: "
    )
    path = input(prompt).strip()
    if not path:
        return None
    if not os.path.isfile(path):
        print(f"Файл не найден по пути: {path}. Пропускаю этот отчёт.")
        return None
    return path


def _ask_multiple_paths(meta: dict) -> list:
    prompt = (
        f"\n[{'ОБЯЗАТЕЛЕН' if meta['required'] else 'опционально'}] {meta['title']}\n"
        f"Где взять: {meta['where']}\n"
        "Можно указать несколько недельных файлов через запятую.\n"
        "Введите путь(и) к файлам или нажмите Enter, чтобы пропустить: "
    )
    raw = input(prompt).strip()
    if not raw:
        return []
    paths = []
    for part in raw.split(","):
        candidate = part.strip().strip('"')
        if not candidate:
            continue
        if os.path.isfile(candidate):
            paths.append(candidate)
        else:
            print(f"Файл не найден по пути: {candidate}. Пропускаю его.")
    return paths


def ask_report_files(report_key: str, meta: dict):
    if meta.get("supports_multiple"):
        return _ask_multiple_paths(meta)
    return _ask_single_path(meta)


def main():
    print("=== Аудит кабинета Wildberries ===")
    file_paths = {}
    for report_key, meta in REPORTS_REGISTRY.items():
        file_paths[report_key] = ask_report_files(report_key, meta)

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
