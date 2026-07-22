# AGENTS.md

## Cursor Cloud specific instructions

This is a small, self-contained Python project (no database, no external services, no secrets). It is a Wildberries seller-cabinet audit tool with a **Streamlit web UI** (`streamlit_app.py`) and a **CLI** (`main.py`); both share the same pipeline (`run_audit_pipeline` in `main.py`). The UI and report content are in Russian. Dependencies are in `requirements.txt` (pandas, openpyxl, numpy, streamlit, xlsxwriter).

### Environment / dependencies
- The startup update script installs everything into a project virtualenv at `.venv/`. Use that interpreter, e.g. `.venv/bin/python` and `.venv/bin/streamlit`.
- Non-obvious gotcha: `python3 -m venv` does **not** work in this environment (`ensurepip` is missing and the `apt` mirrors are blocked by egress, so `python3.12-venv` cannot be installed). The env is therefore created with `virtualenv` (installed via `pip install --break-system-packages virtualenv`). Don't switch back to `python -m venv`.

### Run the web app (dev mode)
- `.venv/bin/streamlit run streamlit_app.py` (add `--server.headless true --server.port 8501` in the cloud VM). App serves on port 8501.

### Run the CLI
- `.venv/bin/python main.py` — interactive; it prompts for report file paths (Enter to skip) and a period, then writes `output/WB_Audit_Report_<timestamp>.xlsx`.

### Testing notes
- There is no test suite, linter config, or build step in this repo. "Build/run" == running Streamlit or the CLI.
- Core behavior is **graceful degradation**: missing/invalid input files never crash the pipeline; the corresponding report blocks are marked partial or "no data". The only report expected for a meaningful run is the weekly finance report (`finance_weekly`), which drives sales and cabinet unit-economics blocks.
- To smoke-test end to end, feed an `.xlsx` with columns like `Артикул поставщика`, `Дата продажи`, `К перечислению`, `Логистика`, `Хранение`, `Штрафы` (see `config.py` `COLUMN_MAPPING`) into the finance uploader, then click "Запустить расчёт".
- `output/` and `input/` are gitignored working dirs.
