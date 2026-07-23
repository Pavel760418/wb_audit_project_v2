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

from wb_config import COLUMN_MAPPING
from modules.columns import find_column, normalize_name
from modules.loader import LoadResult, get_source_name, load_excel_report

_FINANCE_MAPPING = COLUMN_MAPPING["finance_weekly"]

# Поля, у которых в отчёте WB ровно одна колонка — их безопасно приводить к
# каноническому имени, чтобы недельные файлы совпадали по структуре при объединении.
# Поле "penalty" НЕ нормализуется: в отчёте WB удержания разбиты на несколько
# колонок (штрафы, удержания, прочие удержания), которые нельзя схлопывать в одну.
_SINGLE_VALUE_FIELDS = ("sku", "date", "amount", "logistics", "storage")


def _normalize_to_list(value) -> list:
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return [f for f in value if f is not None]
    return [value]


def _alias_to_canonical() -> dict:
    """Соответствие «вариант подписи (нормализованный) → каноническая подпись».

    Строится только для одно-колоночных полей — многоколоночные удержания
    остаются под своими исходными названиями, чтобы не потерять данные.
    """
    mapping = {}
    for field in _SINGLE_VALUE_FIELDS:
        aliases = _FINANCE_MAPPING.get(field, [])
        if not aliases:
            continue
        canonical = aliases[0]
        for alias in aliases:
            mapping[normalize_name(alias)] = canonical
    return mapping


def count_finance_columns(df: pd.DataFrame) -> int:
    """Сколько полей финансового отчёта удалось сопоставить в файле.

    Использует устойчивый поиск (точное совпадение → по вхождению), чтобы
    распознавать длинные официальные названия колонок Wildberries.
    """
    matched = 0
    for aliases in _FINANCE_MAPPING.values():
        if find_column(df, aliases) is not None:
            matched += 1
    return matched


def normalize_finance_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Приводит названия одно-колоночных полей к каноническим для совпадения недель.

    Переименование НЕ создаёт дубликатов и НЕ удаляет колонки: если каноническое
    имя уже занято другой колонкой, исходная колонка сохраняется под своим именем.
    Так данные не теряются (в отличие от прежнего схлопывания дубликатов).
    """
    alias_map = _alias_to_canonical()
    existing = {normalize_name(c) for c in df.columns}
    taken = set()
    rename = {}
    for col in df.columns:
        canonical = alias_map.get(normalize_name(col))
        if not canonical or canonical == col:
            taken.add(normalize_name(col))
            continue
        canonical_norm = normalize_name(canonical)
        # Переименовываем только если каноническое имя ещё не занято и не назначено.
        if canonical_norm not in existing and canonical_norm not in taken:
            rename[col] = canonical
            taken.add(canonical_norm)
        else:
            taken.add(normalize_name(col))
    if rename:
        df = df.rename(columns=rename)
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
