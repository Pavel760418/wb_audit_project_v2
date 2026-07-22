"""
Аналитические расчёты по продажам, финансам и юнит-экономике.
Каждая функция проверяет наличие нужных данных и возвращает частичный
результат с пометкой "insufficient_data" вместо падения с ошибкой.
"""
from __future__ import annotations

import pandas as pd


def calc_sales_summary(load_results: dict) -> dict:
    finance = load_results.get("finance_weekly")
    if not finance or finance.dataframe is None or finance.status == "missing":
        return {"status": "insufficient_data", "message": "Нет данных финансового отчёта для расчёта продаж"}

    df = finance.dataframe
    amount_col = _find_column(df, ["К перечислению", "Сумма к перечислению", "Итого к оплате"])
    if amount_col is None:
        return {"status": "insufficient_data", "message": "В файле не найдена колонка суммы"}

    total_amount = pd.to_numeric(df[amount_col], errors="coerce").sum()
    return {"status": "ok", "total_amount": float(total_amount), "rows": len(df)}


def calc_unit_economics_cabinet(load_results: dict) -> dict:
    finance = load_results.get("finance_weekly")
    if not finance or finance.dataframe is None or finance.status == "missing":
        return {"status": "insufficient_data", "message": "Нет финансового отчёта для расчёта юнит-экономики кабинета"}

    df = finance.dataframe
    amount_col = _find_column(df, ["К перечислению", "Сумма к перечислению", "Итого к оплате"])
    logistics_col = _find_column(df, ["Логистика", "Стоимость логистики"])
    storage_col = _find_column(df, ["Хранение", "Стоимость хранения"])
    penalty_col = _find_column(df, ["Штрафы", "Удержания"])

    revenue = pd.to_numeric(df[amount_col], errors="coerce").sum() if amount_col else 0.0
    logistics = pd.to_numeric(df[logistics_col], errors="coerce").sum() if logistics_col else 0.0
    storage = pd.to_numeric(df[storage_col], errors="coerce").sum() if storage_col else 0.0
    penalties = pd.to_numeric(df[penalty_col], errors="coerce").sum() if penalty_col else 0.0

    net_result = revenue - logistics - storage - penalties
    missing_fields = [
        name
        for name, col in [("логистика", logistics_col), ("хранение", storage_col), ("штрафы", penalty_col)]
        if col is None
    ]

    return {
        "status": "partial" if missing_fields else "ok",
        "revenue": float(revenue),
        "logistics": float(logistics),
        "storage": float(storage),
        "penalties": float(penalties),
        "net_result": float(net_result),
        "missing_fields": missing_fields,
    }


def calc_unit_economics_sku(load_results: dict) -> dict:
    finance = load_results.get("finance_weekly")
    cost_price = load_results.get("cost_price_manual")

    if not finance or finance.dataframe is None or finance.status == "missing":
        return {"status": "insufficient_data", "message": "Нет финансового отчёта для юнит-экономики SKU"}

    if not cost_price or cost_price.dataframe is None or cost_price.status == "missing":
        return {
            "status": "partial",
            "message": "Себестоимость не загружена — расчёт возможен только на уровне валовой выручки, без чистой прибыли",
        }

    return {"status": "ok", "message": "Полный расчёт юнит-экономики SKU доступен"}


def calc_stock_risks(load_results: dict) -> dict:
    stocks = load_results.get("stocks_report")
    if not stocks or stocks.dataframe is None or stocks.status == "missing":
        return {"status": "insufficient_data", "message": "Нет отчёта по остаткам для расчёта рисков дефицита"}

    df = stocks.dataframe
    available_col = _find_column(df, ["Доступно", "Остаток"])
    if available_col is None:
        return {"status": "insufficient_data", "message": "В отчёте по остаткам не найдена колонка доступности"}

    low_stock = df[pd.to_numeric(df[available_col], errors="coerce") <= 0]
    return {"status": "ok", "out_of_stock_count": int(len(low_stock))}


def _find_column(df: pd.DataFrame, candidates: list) -> str | None:
    for candidate in candidates:
        if candidate in df.columns:
            return candidate
    return None
