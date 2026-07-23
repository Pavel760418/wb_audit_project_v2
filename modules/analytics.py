"""
Аналитические расчёты по продажам, финансам и юнит-экономике.

Каждая функция проверяет наличие нужных данных и возвращает СЛУЖЕБНЫЙ словарь с
результатом и диагностикой (какая колонка сопоставлена, какие есть в файле). Эти
словари используются в Streamlit-диагностике и как источник значений для листов
итогового Excel — но НЕ пишутся в ячейки как есть (раскладку делает report_builder).

Распознавание колонок устойчиво к вариантам названий WB (см. modules/columns.py),
а суммы считаются с аккуратным приведением текстовых чисел (запятая-разделитель).
"""
from __future__ import annotations

import pandas as pd

from wb_config import COLUMN_MAPPING
from modules.columns import find_all_columns, find_column, sum_column, sum_columns


def _finance_candidates(field: str) -> list:
    return COLUMN_MAPPING["finance_weekly"].get(field, [])


def calc_sales_summary(load_results: dict) -> dict:
    finance = load_results.get("finance_weekly")
    if not finance or finance.dataframe is None or finance.status == "missing":
        return {
            "status": "insufficient_data",
            "message": "Финансовый отчёт не загружен — продажи рассчитать нельзя.",
        }

    df = finance.dataframe
    amount_col = find_column(df, _finance_candidates("amount"))
    available = [str(c) for c in df.columns]

    if amount_col is None:
        return {
            "status": "insufficient_data",
            "message": "Не удалось сопоставить колонку суммы к перечислению.",
            "available_columns": available,
        }

    total_amount = sum_column(df, amount_col)
    return {
        "status": "ok",
        "total_amount": round(float(total_amount), 2),
        "rows": int(len(df)),
        "matched_column": amount_col,
        "available_columns": available,
    }


def calc_unit_economics_cabinet(load_results: dict) -> dict:
    finance = load_results.get("finance_weekly")
    if not finance or finance.dataframe is None or finance.status == "missing":
        return {
            "status": "insufficient_data",
            "message": "Финансовый отчёт не загружен — юнит-экономику кабинета рассчитать нельзя.",
        }

    df = finance.dataframe
    amount_col = find_column(df, _finance_candidates("amount"))
    logistics_col = find_column(df, _finance_candidates("logistics"))
    storage_col = find_column(df, _finance_candidates("storage"))
    # Удержания в отчёте WB разбиты на несколько колонок (штрафы, удержания и т.п.) —
    # суммируем все найденные, чтобы чистый результат был полным.
    penalty_cols = find_all_columns(df, _finance_candidates("penalty"))

    revenue = sum_column(df, amount_col)
    logistics = sum_column(df, logistics_col)
    storage = sum_column(df, storage_col)
    penalties = sum_columns(df, penalty_cols)

    net_result = revenue - logistics - storage - penalties

    matched_columns = {
        "amount": amount_col,
        "logistics": logistics_col,
        "storage": storage_col,
        "penalty": ", ".join(penalty_cols) if penalty_cols else None,
    }
    missing_fields = [
        label
        for label, col in [
            ("выручка/сумма", amount_col),
            ("логистика", logistics_col),
            ("хранение", storage_col),
            ("штрафы и удержания", penalty_cols or None),
        ]
        if not col
    ]

    return {
        "status": "partial" if missing_fields else "ok",
        "revenue": round(float(revenue), 2),
        "logistics": round(float(logistics), 2),
        "storage": round(float(storage), 2),
        "penalties": round(float(penalties), 2),
        "net_result": round(float(net_result), 2),
        "missing_fields": missing_fields,
        "matched_columns": matched_columns,
        "available_columns": [str(c) for c in df.columns],
    }


def calc_unit_economics_sku(load_results: dict) -> dict:
    finance = load_results.get("finance_weekly")
    cost_price = load_results.get("cost_price_manual")

    if not finance or finance.dataframe is None or finance.status == "missing":
        return {
            "status": "insufficient_data",
            "message": "Финансовый отчёт не загружен — юнит-экономику по SKU рассчитать нельзя.",
        }

    df = finance.dataframe
    sku_col = find_column(df, _finance_candidates("sku"))
    amount_col = find_column(df, _finance_candidates("amount"))

    has_cost = bool(cost_price and cost_price.dataframe is not None and cost_price.status != "missing")

    # Валовая выручка по SKU доступна, если есть артикул и сумма.
    sku_revenue = []
    if sku_col and amount_col:
        grouped = (
            df.assign(_amount=df[amount_col])
            .groupby(sku_col, dropna=False)["_amount"]
            .apply(lambda s: float(pd.to_numeric(s, errors="coerce").sum()))
        )
        sku_revenue = [
            {"sku": str(idx), "revenue": round(float(val), 2)}
            for idx, val in grouped.sort_values(ascending=False).items()
        ]

    if not has_cost:
        status = "partial" if sku_revenue else "insufficient_data"
        return {
            "status": status,
            "message": (
                "Себестоимость не загружена: доступна только валовая выручка по SKU, без чистой прибыли."
                if sku_revenue
                else "Недостаточно данных: нужны артикул и сумма в финансовом отчёте."
            ),
            "matched_columns": {"sku": sku_col, "amount": amount_col},
            "sku_count": len(sku_revenue),
            "top_sku": sku_revenue[:15],
        }

    return {
        "status": "ok" if sku_revenue else "partial",
        "message": "Себестоимость загружена — доступен расчёт валовой выручки по SKU с учётом себестоимости.",
        "matched_columns": {"sku": sku_col, "amount": amount_col},
        "sku_count": len(sku_revenue),
        "top_sku": sku_revenue[:15],
    }


def calc_stock_risks(load_results: dict) -> dict:
    stocks = load_results.get("stocks_report")
    if not stocks or stocks.dataframe is None or stocks.status == "missing":
        return {
            "status": "insufficient_data",
            "message": "Отчёт по остаткам не загружен — риски дефицита рассчитать нельзя.",
        }

    df = stocks.dataframe
    available_col = find_column(df, COLUMN_MAPPING["stocks_report"]["available"])
    if available_col is None:
        return {
            "status": "insufficient_data",
            "message": "В отчёте по остаткам не удалось сопоставить колонку доступности.",
            "available_columns": [str(c) for c in df.columns],
        }

    from modules.columns import to_number

    available_values = to_number(df[available_col])
    out_of_stock = int((available_values <= 0).sum())
    total_positions = int(len(df))
    return {
        "status": "ok",
        "out_of_stock_count": out_of_stock,
        "total_positions": total_positions,
        "matched_column": available_col,
    }
