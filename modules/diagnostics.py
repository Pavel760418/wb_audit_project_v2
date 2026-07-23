"""
Диагностика распознавания колонок входных отчётов.

Показывает, какие поля ожидались, какие реальные колонки есть в файле и что с чем
сопоставлено. Используется и в Streamlit, и на служебном листе итогового Excel —
чтобы было понятно, почему тот или иной блок посчитан частично или revenue = 0.
"""
from __future__ import annotations

from config import COLUMN_MAPPING, FINANCE_FIELD_LABELS
from modules.columns import find_column


def diagnose_mapping(df, mapping: dict, labels: dict | None = None) -> list:
    """Возвращает список строк диагностики для одного отчёта.

    Каждая строка: поле, человеко-понятная подпись, сопоставленная колонка,
    список допустимых вариантов названий.
    """
    rows = []
    for field, aliases in mapping.items():
        matched = find_column(df, aliases) if df is not None else None
        rows.append(
            {
                "field": field,
                "label": (labels or {}).get(field, field),
                "matched": matched,
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
        "fields": diagnose_mapping(df, COLUMN_MAPPING["finance_weekly"], FINANCE_FIELD_LABELS),
    }
