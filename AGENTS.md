# AGENTS.md

## Cursor Cloud specific instructions

### Repository layout note
- The `main` branch is essentially empty (only `.gitkeep`). The actual application code lives on the `codex/develop-patent-data-processing-software` branch (`app.py`, `launcher.py`, `requirements.txt`, `README.md`). Do development from a branch that contains that code.

### What this is
- A single-service Streamlit app: a Chinese-language patent table classification tool ("专利表格分类工具"). It ingests an uploaded Excel/CSV of patent data, normalizes applicants, classifies each patent, and generates downloadable Word (`.docx`) / Excel (`.xlsx`) reports. Details in `README.md`.

### Run (dev)
- `python3 -m streamlit run app.py` (default port 8501).
- Non-obvious: the `streamlit` console script installs to `~/.local/bin`, which is not on `PATH`, so invoke via `python3 -m streamlit` rather than bare `streamlit`.
- `launcher.py` + `build_windows_exe.bat` exist only to package a Windows `.exe` via PyInstaller; they are not usable on Linux — do not try to run them here.

### Test / lint
- There is no automated test suite and no configured linter in this repo. Use `python3 -m py_compile app.py launcher.py` for a quick syntax check. The core pipeline can be exercised headlessly by importing `app` and calling `app.build_rows(df, {})` then `app.generate_docs(rows)`.
