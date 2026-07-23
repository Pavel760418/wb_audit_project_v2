"""
Веб-интерфейс аудита кабинета Wildberries.

Пользователь загружает доступные отчёты, при необходимости несколько недельных
финансовых отчётов, скачивает шаблоны ручных таблиц, запускает расчёт и получает
итоговый Excel-файл. Отсутствие части файлов не останавливает расчёт — недоступные
блоки отмечаются явно (принцип «мягкой деградации»).

Тяжёлая бизнес-логика вынесена в модули (loader, finance_merge, report_builder,
templates); здесь — только интерфейс и вызов единой pipeline-функции.
"""
from __future__ import annotations

import io
from datetime import date

import streamlit as st

from wb_config import GROUP_MANUAL, GROUP_OPTIONAL, GROUP_REQUIRED, REPORTS_REGISTRY
from main import run_audit_pipeline
from modules import templates
from modules.diagnostics import finance_column_diagnostics

st.set_page_config(page_title="Аудит кабинета Wildberries", layout="wide")

st.title("Аудит кабинета Wildberries")
st.caption(
    "Загрузите доступные отчёты и запустите расчёт. Обязателен только еженедельный "
    "финансовый отчёт — его можно загрузить сразу за несколько недель. Отсутствующие "
    "файлы не остановят расчёт: соответствующие блоки будут отмечены как частичные или недоступные."
)


def _reports_by_group(group: str) -> list:
    return [(key, meta) for key, meta in REPORTS_REGISTRY.items() if meta.get("group") == group]


def _render_report_hint(meta: dict) -> None:
    st.caption(f"Где взять: {meta['where']}")
    st.caption(f"Если не загрузить: {meta['unavailable_if_missing']}")


uploaded_files: dict = {}

# ----------------------------------------------------------------------------
# 1. Обязательные данные (еженедельный финансовый отчёт, поддержка мультизагрузки)
# ----------------------------------------------------------------------------
st.header("1. Обязательные данные")
for report_key, meta in _reports_by_group(GROUP_REQUIRED):
    st.subheader(f"🔴 {meta['title']}")
    _render_report_hint(meta)
    if meta.get("supports_multiple"):
        st.info(
            "Финансовые отчёты в кабинете формируются по неделям. Чтобы проанализировать "
            "месяц или квартал, загрузите сразу несколько недельных файлов — они объединятся "
            "в один период, а дубликаты строк будут удалены автоматически."
        )
        uploaded_files[report_key] = st.file_uploader(
            "Выберите один или несколько недельных финансовых отчётов (.xlsx)",
            type=["xlsx"],
            key=report_key,
            accept_multiple_files=True,
        )
    else:
        uploaded_files[report_key] = st.file_uploader(
            "Выберите файл (.xlsx)", type=["xlsx"], key=report_key
        )

st.divider()

# ----------------------------------------------------------------------------
# 2. Дополнительные данные (выгрузки из кабинета WB)
# ----------------------------------------------------------------------------
st.header("2. Дополнительные данные")
st.caption("Отчёты из кабинета Wildberries. Каждый расширяет аналитику, но не является обязательным.")
optional_reports = _reports_by_group(GROUP_OPTIONAL)
cols = st.columns(2)
for idx, (report_key, meta) in enumerate(optional_reports):
    with cols[idx % 2]:
        with st.expander(f"🟡 {meta['title']}"):
            _render_report_hint(meta)
            uploaded_files[report_key] = st.file_uploader(
                "Выберите файл (.xlsx)", type=["xlsx"], key=report_key
            )

st.divider()

# ----------------------------------------------------------------------------
# 3. Ручные справочники (заполняются пользователем по шаблонам)
# ----------------------------------------------------------------------------
st.header("3. Ручные справочники")
st.caption("Таблицы, которые заполняются вручную. Готовые шаблоны можно скачать в разделе 8.")
manual_reports = _reports_by_group(GROUP_MANUAL)
cols = st.columns(2)
for idx, (report_key, meta) in enumerate(manual_reports):
    with cols[idx % 2]:
        with st.expander(f"🟢 {meta['title']}"):
            _render_report_hint(meta)
            uploaded_files[report_key] = st.file_uploader(
                "Выберите файл (.xlsx)", type=["xlsx"], key=report_key
            )

st.divider()

# ----------------------------------------------------------------------------
# 4. Период анализа
# ----------------------------------------------------------------------------
st.header("4. Период анализа")
st.caption("Необязательно. Период указывается в итоговом отчёте для справки.")
period_cols = st.columns(2)
with period_cols[0]:
    period_start = st.date_input("Начало периода", value=None, format="YYYY-MM-DD")
with period_cols[1]:
    period_end = st.date_input("Конец периода", value=None, format="YYYY-MM-DD")

st.divider()


# ----------------------------------------------------------------------------
# 5. Статус загруженных файлов (до запуска расчёта)
# ----------------------------------------------------------------------------
def _count_selected(value) -> int:
    if value is None:
        return 0
    if isinstance(value, (list, tuple)):
        return len([f for f in value if f is not None])
    return 1


st.header("5. Статус загруженных файлов")
status_rows = []
for report_key, meta in REPORTS_REGISTRY.items():
    count = _count_selected(uploaded_files.get(report_key))
    if meta.get("supports_multiple"):
        state = f"загружено файлов: {count}" if count else "не загружено"
    else:
        state = "загружен" if count else "не загружен"
    status_rows.append(
        {
            "Отчёт": meta["title"],
            "Обязателен": "Да" if meta["required"] else "Нет",
            "Состояние": state,
        }
    )
st.dataframe(status_rows, width="stretch", hide_index=True)

finance_selected = _count_selected(uploaded_files.get("finance_weekly"))
if finance_selected == 0:
    st.warning(
        "Загрузите хотя бы один еженедельный финансовый отчёт — без него базовый расчёт невозможен."
    )
else:
    st.success(f"Готово к расчёту: недельных финансовых отчётов выбрано — {finance_selected}.")

st.divider()

# ----------------------------------------------------------------------------
# 6. Запуск расчёта
# ----------------------------------------------------------------------------
st.header("6. Запуск расчёта")
run_clicked = st.button(
    "Запустить расчёт",
    type="primary",
    disabled=finance_selected == 0,
)

if run_clicked:
    with st.spinner("Обрабатываю загруженные отчёты…"):
        workbook, load_results = run_audit_pipeline(
            uploaded_files,
            period_start.isoformat() if isinstance(period_start, date) else None,
            period_end.isoformat() if isinstance(period_end, date) else None,
        )
        buffer = io.BytesIO()
        workbook.save(buffer)
        buffer.seek(0)
        st.session_state["report_bytes"] = buffer.getvalue()
        st.session_state["load_results"] = load_results
    st.success("Расчёт завершён на основе реально загруженных данных.")

st.divider()

# ----------------------------------------------------------------------------
# 7. Скачать итоговый Excel (сразу после расчёта — до диагностики,
# чтобы кнопка не пропадала, если ниже что-то упадёт)
# ----------------------------------------------------------------------------
st.header("7. Скачать итоговый Excel")
if "report_bytes" in st.session_state:
    st.download_button(
        label="Скачать итоговый Excel-файл",
        data=st.session_state["report_bytes"],
        file_name="WB_Audit_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Итоговый файл появится здесь после запуска расчёта.")

# Диагностика результатов — после кнопки скачивания.
if "load_results" in st.session_state:
    load_results = st.session_state["load_results"]

    finance = load_results.get("finance_weekly")
    if finance and getattr(finance, "file_details", None):
        st.subheader("Недельные финансовые отчёты")
        st.caption(
            f"Всего: {finance.files_total} · успешно: {finance.files_ok} · "
            f"с предупреждениями: {finance.files_warning} · не вошло в расчёт: {finance.files_failed}"
        )
        st.dataframe(
            [
                {
                    "Файл": d.get("name", "—"),
                    "Статус": d.get("status", ""),
                    "Строк": d.get("rows", 0),
                    "Комментарий": d.get("message", ""),
                }
                for d in finance.file_details
            ],
            width="stretch",
            hide_index=True,
        )
        for warn in getattr(finance, "warnings", []) or []:
            st.warning(warn)

    st.subheader("Диагностика загруженных файлов")
    diag_rows = []
    for report_key, result in load_results.items():
        diag_rows.append(
            {
                "Отчёт": REPORTS_REGISTRY[report_key]["title"],
                "Статус": result.status,
                "Комментарий": result.message,
                "Найденные колонки": ", ".join(result.detected_columns) if result.detected_columns else "—",
            }
        )
    st.dataframe(diag_rows, width="stretch", hide_index=True)

    # Диагностика сопоставления колонок финансового отчёта — помогает понять,
    # почему блок посчитан частично или выручка равна нулю.
    try:
        finance_diag = finance_column_diagnostics(load_results)
    except Exception as exc:  # noqa: BLE001 — UI не должен падать из‑за диагностики
        st.warning(f"Не удалось построить диагностику колонок: {exc}")
        finance_diag = None

    if finance_diag:
        with st.expander("Диагностика распознавания колонок (финансовый отчёт)"):
            unmatched = [f["label"] for f in finance_diag["fields"] if not f["matched"]]
            if unmatched:
                st.warning("Не удалось сопоставить поля: " + ", ".join(unmatched))
            else:
                st.success("Все ключевые поля финансового отчёта распознаны.")
            st.dataframe(
                [
                    {
                        "Поле": f["label"],
                        "Сопоставленная колонка": f["matched"] or "не сопоставлено",
                        "Допустимые варианты": ", ".join(f["candidates"]),
                    }
                    for f in finance_diag["fields"]
                ],
                width="stretch",
                hide_index=True,
            )
            st.caption("Все колонки, найденные в файле:")
            st.write(", ".join(finance_diag["available_columns"]) or "—")

st.divider()

# ----------------------------------------------------------------------------
# 8. Скачать шаблоны Excel
# ----------------------------------------------------------------------------
st.header("8. Скачать шаблоны Excel")
st.caption(
    "Готовые шаблоны для ручного заполнения. Не переименовывайте колонки — модуль ориентируется на них."
)
template_cols = st.columns(2)
for idx, (template_key, meta) in enumerate(templates.TEMPLATES.items()):
    with template_cols[idx % 2]:
        st.markdown(f"**{meta['title']}**")
        st.caption(meta["description"])
        st.download_button(
            label=f"Скачать шаблон «{meta['title']}»",
            data=templates.build_template_bytes(template_key),
            file_name=meta["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"template_{template_key}",
        )
