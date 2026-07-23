"""
Единые утилиты для распознавания и нормализации колонок Excel-отчётов Wildberries.

Выгрузки WB нередко отличаются подписями колонок: лишние пробелы, переносы строк,
неразрывные пробелы, разный регистр, длинные официальные названия. Здесь собран
устойчивый поиск колонок (точное совпадение → совпадение по вхождению) и аккуратное
приведение значений к числам (в том числе из текстовых ячеек с запятой-разделителем).
"""
from __future__ import annotations

import re
from typing import Optional

import pandas as pd

_NBSP = "\xa0"


def normalize_name(name) -> str:
    """Приводит название колонки к каноническому виду для сравнения.

    Убирает переносы строк, неразрывные пробелы, схлопывает повторяющиеся пробелы,
    обрезает края и приводит к нижнему регистру.
    """
    text = str(name).replace("\n", " ").replace("\r", " ").replace(_NBSP, " ")
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def find_column(df: pd.DataFrame, candidates) -> Optional[str]:
    """Находит колонку в df по списку допустимых названий.

    Сначала ищется точное совпадение (после нормализации), затем — совпадение по
    вхождению (подпись-кандидат содержится в реальном названии колонки или наоборот).
    Возвращает исходное название колонки из df или None.
    """
    if df is None or len(df.columns) == 0:
        return None

    norm_cols = []
    for col in df.columns:
        norm_cols.append((normalize_name(col), col))

    # 1) Точное совпадение по нормализованному имени.
    for candidate in candidates:
        nc = normalize_name(candidate)
        if not nc:
            continue
        for norm_col, original in norm_cols:
            if norm_col == nc:
                return original

    # 2) Совпадение по вхождению (для длинных официальных названий WB).
    for candidate in candidates:
        nc = normalize_name(candidate)
        if not nc:
            continue
        for norm_col, original in norm_cols:
            if nc in norm_col or norm_col in nc:
                return original

    return None


def to_number(series: pd.Series) -> pd.Series:
    """Аккуратно приводит колонку к числам.

    Числовые колонки возвращаются как есть; текстовые очищаются от пробелов и
    неразрывных пробелов, запятая трактуется как десятичный разделитель.
    Непреобразуемые значения становятся NaN (в суммах игнорируются).
    """
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    cleaned = (
        series.astype(str)
        .str.replace(_NBSP, "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def sum_column(df: pd.DataFrame, column: Optional[str]) -> float:
    """Безопасно суммирует числовые значения указанной колонки."""
    if not column or column not in df.columns:
        return 0.0
    return float(to_number(df[column]).sum())
