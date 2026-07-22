"""
Загрузка Excel-отчётов WB с graceful degradation.
Отсутствие файла или ошибка чтения не останавливает работу модуля —
источник помечается статусом, а обработка продолжается на доступных данных.
"""
from __future__ import annotations

import pandas as pd
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class LoadResult:
    report_key: str
    status: str  # "ok" | "warning" | "missing" | "error"
    dataframe: Optional[pd.DataFrame] = None
    message: str = ""
    detected_columns: list = field(default_factory=list)


def load_excel_report(report_key: str, file_obj_or_path) -> LoadResult:
    """
    file_obj_or_path: путь к файлу, файловый объект (в т.ч. из st.file_uploader) или None.
    Никогда не бросает исключение наружу — все ошибки конвертируются в LoadResult(status="error"/"missing").
    """
    if file_obj_or_path is None:
        return LoadResult(report_key=report_key, status="missing", message="Файл не предоставлен пользователем")

    try:
        df = pd.read_excel(file_obj_or_path, engine="openpyxl")
    except Exception as exc:  # noqa: BLE001 - обязательно перехватываем всё, чтобы не рушить пайплайн
        return LoadResult(
            report_key=report_key,
            status="error",
            message=f"Не удалось прочитать файл: {exc}",
        )

    if df.empty:
        return LoadResult(
            report_key=report_key,
            status="warning",
            dataframe=df,
            message="Файл прочитан, но не содержит строк данных",
            detected_columns=list(df.columns),
        )

    return LoadResult(
        report_key=report_key,
        status="ok",
        dataframe=df,
        message="Файл успешно загружен",
        detected_columns=list(df.columns),
    )


def load_all_reports(file_map: dict) -> dict:
    """
    file_map: {report_key: file_obj_or_path_or_None}
    Возвращает {report_key: LoadResult} для всех ключей реестра, даже если файл не был передан.
    """
    results = {}
    for report_key, file_obj in file_map.items():
        results[report_key] = load_excel_report(report_key, file_obj)
    return results
