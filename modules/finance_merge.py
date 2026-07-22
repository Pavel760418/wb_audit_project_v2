"""
Объединение нескольких недельных финансовых отчётов Wildberries в единый период.

В кабинете Wildberries финансовые отчёты формируются понедельно, поэтому анализ
месяца или квартала требует загрузки нескольких недельных файлов. Этот модуль:

- принимает список файлов (или один файл, или None);
- нормализует названия колонок к каноническим (варианты подписей → единая подпись);
- проверяет, что файл действительно похож на финансовый отчёт WB;
- объединяет данные в один DataFrame и удаляет полностью дублирующиеся строки;
- продолжает расчёт по доступным файлам, если часть файлов не читается;
- ведёт подробную диагностику по каждому файлу.
"""
from __future__ import annotations

import pandas as pd

from config import COLUMN_MAPPING
from modules.loader import LoadResult, get_source_name, load_excel_report

_FINANCE_MAPPING = COLUMN_MAPPING["finance_weekly"]


def _normalize_to_list(value) -> list:
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return [f for f in value if f is not None]
    return [value]


def _alias_to_canonical() -> dict:
    """Строит соответствие «вариант подписи (в нижнем регистре) → каноническая подпись»."""
    mapping = {}
    for aliases in _FINANCE_MAPPING.values():
        canonical = aliases[0]
        for alias in aliases:
            mapping[alias.strip().lower()] = canonical
    return mapping


def _known_finance_columns() -> set:
    known = set()
    for aliases in _FINANCE_MAPPING.values():
        known.update(alias.strip().lower() for alias in aliases)
    return known


def count_finance_columns(df: pd.DataFrame) -> int:
    """Сколько распознанных финансовых колонок присутствует в файле."""
    known = _known_finance_columns()
    return sum(1 for col in df.columns if str(col).strip().lower() in known)


def normalize_finance_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Приводит названия колонок к каноническим, чтобы файлы недель совпадали по структуре."""
    alias_map = _alias_to_canonical()
    rename = {}
    for col in df.columns:
        key = str(col).strip().lower()
        canonical = alias_map.get(key)
        if canonical and canonical != col:
            rename[col] = canonical
    if rename:
        df = df.rename(columns=rename)
    # На случай, если после переименования появились одинаковые колонки — оставляем первую.
    df = df.loc[:, ~df.columns.duplicated()]
    return df


def load_finance_reports(value) -> LoadResult:
    """Загружает и объединяет один или несколько недельных финансовых отчётов."""
    files = _normalize_to_list(value)

    if not files:
        return LoadResult(
            report_key="finance_weekly",
            status="missing",
            message="Финансовые отчёты не загружены. Это обязательный источник данных.",
        )

    file_details = []
    frames = []
    warnings = []
    column_sets = []

    for file_obj in files:
        name = get_source_name(file_obj)
        single = load_excel_report("finance_weekly", file_obj)
        detail = {
            "name": name,
            "status": single.status,
            "message": single.message,
            "rows": 0,
        }

        if single.status in ("ok", "warning") and single.dataframe is not None:
            matched = count_finance_columns(single.dataframe)
            if matched == 0:
                detail["status"] = "error"
                detail["message"] = (
                    f"{name}: не найдены ожидаемые колонки финансового отчёта "
                    "(например «К перечислению»). Возможно, загружен файл другого типа."
                )
            else:
                df = normalize_finance_columns(single.dataframe)
                detail["rows"] = len(df)
                if single.status == "warning":
                    warnings.append(single.message)
                frames.append(df)
                column_sets.append(tuple(df.columns))
        # status "error"/"missing" — файл в расчёт не входит, диагностика уже заполнена.

        file_details.append(detail)

    files_total = len(files)
    files_ok = sum(1 for d in file_details if d["status"] == "ok")
    files_warning = sum(1 for d in file_details if d["status"] == "warning")
    files_failed = sum(1 for d in file_details if d["status"] in ("error", "missing"))

    if not frames:
        return LoadResult(
            report_key="finance_weekly",
            status="error",
            message=(
                f"Загружено файлов: {files_total}, но ни один не удалось использовать. "
                "Проверьте, что это недельные финансовые отчёты в формате .xlsx."
            ),
            file_details=file_details,
            files_total=files_total,
            files_ok=files_ok,
            files_warning=files_warning,
            files_failed=files_failed,
        )

    # Предупреждаем, если структура файлов различается (разный набор колонок).
    if len(set(column_sets)) > 1:
        warnings.append(
            "Загруженные недельные отчёты отличаются по набору колонок — "
            "данные объединены по совпадающим полям, недостающие значения оставлены пустыми."
        )

    combined = pd.concat(frames, ignore_index=True, sort=False)
    rows_before = len(combined)
    combined = combined.drop_duplicates(ignore_index=True)
    removed_duplicates = rows_before - len(combined)

    status = "ok" if (files_failed == 0 and files_warning == 0 and not warnings) else "warning"

    message = (
        f"Недельных отчётов загружено: {files_total}. "
        f"В расчёт вошло: {len(frames)}. "
        f"Строк после объединения: {len(combined)} "
        f"(удалено дубликатов: {removed_duplicates})."
    )

    return LoadResult(
        report_key="finance_weekly",
        status=status,
        dataframe=combined,
        message=message,
        detected_columns=list(combined.columns),
        warnings=warnings,
        file_details=file_details,
        files_total=files_total,
        files_ok=files_ok,
        files_warning=files_warning,
        files_failed=files_failed,
    )
