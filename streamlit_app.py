"""
Streamlit-оболочка для тестовой загрузки Excel-отчётов WB и отладки ingestion-слоя.
Использует ту же pipeline-функцию, что и main.py, чтобы CLI и веб-версия совпадали.
"""
from __future__ import annotations

import io
from datetime import date

import streamlit as st

from config import REPORTS_REGISTRY
from main import run_audit_pipeline

st.set_page_config(page_title="WB Audit — тестовая загрузка", layout="wide")

st.title("Аудит кабинета Wildberries — тестовый модуль загрузки")
st.caption("Загрузите доступные отчёты. Отсутствующие файлы не остановят расчёт — блоки будут помечены как частичные или недоступные.")

uploaded_files = {}
with st.sidebar:
    st.header("Период анализа")
    period_start = st.date_input("Начало периода", value=None)
    period_end = st.date_input("Конец периода", value=None)

st.subheader("Загрузка отчётов")
cols = st.columns(2)
for idx, (report_key, meta) in enumerate(REPORTS_REGISTRY.items()):
    col = cols[idx % 2]
    with col:
        label = f"{'🔴 Обязателен' if meta['required'] else '🟡 Опционально'} — {meta['title']}"
        st.markdown(f"**{label}**")
        st.caption(f"Где искать в кабинете: {meta['where']}")
        uploaded = st.file_uploader(
            label="Выберите Excel-файл",
            type=["xlsx", "xls"],
            key=report_key,
            label_visibility="collapsed",
        )
        uploaded_files[report_key] = uploaded

st.divider()

if st.button("Запустить расчёт", type="primary"):
    period_start_str = period_start.isoformat() if isinstance(period_start, date) else None
    period_end_str = period_end.isoformat() if isinstance(period_end, date) else None

    with st.spinner("Обрабатываю загруженные отчёты..."):
        workbook, load_results = run_audit_pipeline(uploaded_files, period_start_str, period_end_str)

    st.success("Расчёт завершён на основе реально загруженных данных.")

    st.subheader("Диагностика загруженных файлов")
    diag_rows = []
    for report_key, result in load_results.items():
        title = REPORTS_REGISTRY[report_key]["title"]
        diag_rows.append(
            {
                "Отчёт": title,
                "Статус": result.status,
                "Комментарий": result.message,
                "Найденные колонки": ", ".join(result.detected_columns) if result.detected_columns else "—",
            }
        )
    st.dataframe(diag_rows, use_container_width=True)

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="Скачать итоговый Excel-файл",
        data=buffer,
        file_name="WB_Audit_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Загрузите файлы (хотя бы еженедельный финансовый отчёт) и нажмите «Запустить расчёт».")
