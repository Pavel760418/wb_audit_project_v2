"""
Загрузка Excel-отчётов Wildberries с «мягкой деградацией» (graceful degradation).

Отсутствие файла или ошибка чтения не останавливают работу модуля:
источник помечается статусом, а обработка продолжается на доступных данных.

Дополнительно реализовано устойчивое чтение файлов:
- понятное сообщение при загрузке устаревшего формата .xls;
- автоматический поиск строки заголовков, если сверху есть лишние строки;
- удаление полностью пустых строк и столбцов;
- перехват любых ошибок чтения без падения приложения.
"""
from __future__ import annotations

import io
import os
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd


@dataclass
class LoadResult:
    """Результат загрузки одного логического отчёта.

    Для отчётов с мультизагрузкой (несколько недельных файлов) заполняются
    поля file_details и счётчики files_* — по ним строится статус загрузки.
    """

    report_key: str
    status: str  # "ok" | "warning" | "missing" | "error"
    dataframe: Optional[pd.DataFrame] = None
    message: str = ""
    detected_columns: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    # Диагностика по каждому файлу (актуально для мультизагрузки).
    file_details: list = field(default_factory=list)
    files_total: int = 0
    files_ok: int = 0
    files_warning: int = 0
    files_failed: int = 0


def get_source_name(file_obj_or_path) -> str:
    """Возвращает человекочитаемое имя источника (для диагностики)."""
    if file_obj_or_path is None:
        return "—"
    if isinstance(file_obj_or_path, str):
        return os.path.basename(file_obj_or_path)
    return getattr(file_obj_or_path, "name", "загруженный файл")


def _get_extension(file_obj_or_path) -> str:
    name = None
    if isinstance(file_obj_or_path, str):
        name = file_obj_or_path
    else:
        name = getattr(file_obj_or_path, "name", None)
    if name:
        return os.path.splitext(name)[1].lower()
    return ""


def _to_buffer(file_obj_or_path) -> io.BytesIO:
    """Считывает содержимое источника в буфер, чтобы файл можно было прочитать повторно."""
    if isinstance(file_obj_or_path, str):
        with open(file_obj_or_path, "rb") as fh:
            data = fh.read()
    else:
        try:
            file_obj_or_path.seek(0)
        except Exception:  # noqa: BLE001
            pass
        data = file_obj_or_path.read()
    return io.BytesIO(data)


def _looks_unnamed(columns) -> bool:
    """Признак того, что заголовки не распознаны (много «Unnamed»/пустых колонок)."""
    if len(columns) == 0:
        return True
    unnamed = sum(1 for c in columns if str(c).strip() == "" or str(c).startswith("Unnamed"))
    return unnamed > len(columns) / 2


def _find_header_row(buffer: io.BytesIO, max_scan: int = 15) -> Optional[int]:
    """Ищет наиболее вероятную строку заголовков среди первых строк файла.

    Заголовком считается строка с наибольшим числом непустых текстовых ячеек.
    """
    buffer.seek(0)
    raw = pd.read_excel(buffer, engine="openpyxl", header=None, nrows=max_scan)
    best_row, best_score = None, 0
    for idx in range(len(raw)):
        row = raw.iloc[idx]
        non_empty = row.notna().sum()
        text_cells = sum(1 for v in row if isinstance(v, str) and v.strip())
        score = int(non_empty) + text_cells
        if score > best_score and non_empty >= 2:
            best_score, best_row = score, idx
    return best_row


def read_excel_robust(file_obj_or_path) -> pd.DataFrame:
    """Устойчиво читает Excel-файл в DataFrame.

    Бросает ValueError с понятным русским сообщением при неустранимой ошибке.
    """
    ext = _get_extension(file_obj_or_path)
    if ext == ".xls":
        raise ValueError(
            "Загружен файл в устаревшем формате .xls. Откройте его в Excel и сохраните как "
            "«Книга Excel (*.xlsx)», затем загрузите заново."
        )
    if ext and ext not in (".xlsx", ".xlsm"):
        raise ValueError(
            f"Формат файла «{ext}» не поддерживается. Загрузите отчёт в формате .xlsx."
        )

    buffer = _to_buffer(file_obj_or_path)
    try:
        df = pd.read_excel(buffer, engine="openpyxl")
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"Не удалось прочитать файл как Excel (.xlsx): {exc}") from exc

    # Если заголовки не распознаны (лишние строки сверху) — ищем строку заголовков.
    if _looks_unnamed(df.columns):
        header_row = _find_header_row(buffer)
        if header_row is not None:
            buffer.seek(0)
            df = pd.read_excel(buffer, engine="openpyxl", header=header_row)

    # Убираем полностью пустые строки и столбцы, чистим имена колонок.
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.reset_index(drop=True)
    return df


def load_excel_report(report_key: str, file_obj_or_path) -> LoadResult:
    """Загружает один Excel-файл.

    Никогда не бросает исключение наружу — любые ошибки превращаются в
    LoadResult со статусом "error"/"missing"/"warning".
    """
    if file_obj_or_path is None:
        return LoadResult(
            report_key=report_key,
            status="missing",
            message="Файл не загружен",
        )

    name = get_source_name(file_obj_or_path)

    try:
        df = read_excel_robust(file_obj_or_path)
    except ValueError as exc:
        return LoadResult(report_key=report_key, status="error", message=f"{name}: {exc}")
    except Exception as exc:  # noqa: BLE001 - перехватываем всё, чтобы не уронить пайплайн
        return LoadResult(
            report_key=report_key,
            status="error",
            message=f"{name}: не удалось прочитать файл ({exc}).",
        )

    if df.empty or len(df.columns) == 0:
        return LoadResult(
            report_key=report_key,
            status="warning",
            dataframe=df,
            message=f"{name}: файл прочитан, но не содержит строк с данными.",
            detected_columns=list(df.columns),
        )

    return LoadResult(
        report_key=report_key,
        status="ok",
        dataframe=df,
        message=f"{name}: файл успешно загружен.",
        detected_columns=list(df.columns),
    )


def load_all_reports(file_map: dict) -> dict:
    """Загружает все отчёты реестра.

    Значение в file_map может быть: None, один файл/путь, либо список файлов.
    Для еженедельного финансового отчёта список файлов объединяется в один период.
    """
    # Локальный импорт, чтобы избежать циклической зависимости с finance_merge.
    from modules.finance_merge import load_finance_reports

    results = {}
    for report_key, value in file_map.items():
        if report_key == "finance_weekly":
            results[report_key] = load_finance_reports(value)
            continue

        # Прочие отчёты пока не поддерживают мультизагрузку: берём первый файл.
        if isinstance(value, (list, tuple)):
            value = value[0] if value else None
        results[report_key] = load_excel_report(report_key, value)
    return results
