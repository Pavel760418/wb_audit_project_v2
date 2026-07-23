"""
Диагностика распознавания колонок входных отчётов.

Показывает, какие поля ожидались, какие реальные колонки есть в файле и что с чем
сопоставлено. Используется и в Streamlit, и на служебном листе итогового Excel —
чтобы было понятно, почему тот или иной блок посчитан частично или revenue = 0.
"""
from __future__ import annotations

from wb_config import COLUMN_MAPPING
from modules.columns import find_all_columns

# Подписи живут здесь (а не только в wb_config), чтобы диагностика не падала
# из‑за устаревшего кэша модуля настроек на Streamlit Cloud.
_FINANCE_FIELD_LABELS = {
    "sku": "Артикул",
    "date": "Дата",
    "amount": "Сумма к перечислению",
    "logistics": "Логистика",
    "storage": "Хранение",
    "penalty": "Штрафы и удержания",
}


def diagnose_mapping(df, mapping: dict, labels: dict | None = None) -> list:
    """Возвращает список строк диагностики для одного отчёта.

    Каждая строка: поле, человеко-понятная подпись, сопоставленные колонки (может
    быть несколько, например для удержаний), список допустимых вариантов названий.
    """
    rows = []
    for field, aliases in mapping.items():
        matched_cols = find_all_columns(df, aliases) if df is not None else []
        rows.append(
            {
                "field": field,
                "label": (labels or {}).get(field, field),
                "matched": ", ".join(matched_cols) if matched_cols else "",
                "candidates": list(aliases),
            }
        )
    return rows


def finance_column_diagnostics(load_results: dict) -> dict | None:
    """Диагностика колонок объединённого финансового отчёта (или None, если данных нет)."""
    finance = load_results.get("finance_weekly")
    if not finance or finance.dataframe is None:
        return None

    df = finance.dataframe
    return {
        "available_columns": [str(c) for c in df.columns],
        "rows": int(len(df)),
        "fields": diagnose_mapping(df, COLUMN_MAPPING["finance_weekly"], _FINANCE_FIELD_LABELS),
    }
