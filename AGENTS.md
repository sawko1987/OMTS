# AGENTS.md

Desktop app for generating "Извещение на замену материалов" (Material
Replacement Notice) Excel documents from a template, with a SQLite-backed
catalog, a material-replacement history store, and a PySide6 GUI. Internal-use
tool, single-user, Windows.

## Setup commands

- Create venv:        `python -m venv venv` (or run `install_dependencies.bat`)
- Install deps:       `pip install -r requirements.txt`
- Start dev (Windows):`run.bat`  /  `python main.py`
- Start dev (Linux):  `./run.sh`  (also requires Qt xcb system packages — see README)
- Tests:              `python -m unittest discover -s . -p "test_*.py"`
- Lint / Typecheck:   not configured — use IDE defaults; no `ruff` / `mypy` / `pyright` setup

`requirements.txt` pins three runtime deps: `PySide6>=6.5.0`, `openpyxl>=3.1.0`,
`xlrd>=2.0.0`.

## Project layout

- `main.py` — entry point; wires logging, `QApplication`, `MainWindow`
- `app/` — application code (no business logic in `main.py`)
  - `app/config.py` — paths, `MONTHS`, `WORKSHOPS`, dir bootstrap
  - `app/models.py` — dataclasses: `CatalogEntry`, `MaterialChange`, `PartChanges`, `DocumentData`, `MaterialReplacementSet`
  - `app/excel_generator.py` — **`ExcelGenerator`**; template fill, pagination, multi-sheet split
  - `app/numbering.py` — **`NumberingManager`**; per-month document numbering (DB + JSON fallback)
  - `app/catalog_loader.py` — load + cache `catalog/catalog.xlsx`; replacement-history suggestions
  - `app/database.py` — SQLite schema + connection pool (`data/app.db`)
  - `app/document_store.py`, `app/product_store.py`, `app/history_store.py`, `app/banned_replacements_store.py` — domain stores
  - `app/settings_manager.py` — `data/settings.json`
  - `app/gui/` — PySide6 widgets. Largest: `changes_table_widget.py` (~87 KB), `replacement_sets_editor_widget.py` (~64 KB), `main_window.py` (~44 KB)
  - `app/migrate_to_sqlite.py`, `app/database_restore.py` — one-shot migration / restore scripts
- `catalog/catalog.xlsx` — source-of-truth material catalog (Деталь, Цех, Роль, До, Ед.изм., Норма)
- `templates/Izveshchenie_template.xlsx` — Excel template the generator fills
- `data/` — `app.db`, `history.json`, `numbering.json`, `settings.json`, `template_config.json` + auto-generated `app.db.bak-*` backups
- `output/` — generated `.xlsx` documents (kept under VCS by `.gitkeep`)
- `tools/` — standalone one-off repair scripts (e.g. `restore_bottom_block_via_excel.py`)
- `*.py` at repo root — utility scripts and `test_*.py` unittest suites

## Code style

- Python 3.8+ target; uses `from __future__` style imports sparingly.
- Dataclasses for domain models; `typing` annotations on public functions.
- Module-level `logger = logging.getLogger(__name__)`; `setup_logging()` in `main.py`.
- All user-facing strings are in **Russian** (project domain: industrial
  document automation in Russian). Comments and docstrings mix Russian (most
  legacy) and English (newer). Match the surrounding module's tone.
- Paths always go through `app/config.py` constants — never hard-code
  `Path("data/...")` in business code.
- `app/config.py` auto-creates required directories at import time — don't
  repeat that elsewhere.

## Testing instructions

- Tests live at repo root as `test_*.py` (no `tests/` package, no pytest
  config). Run with: `python -m unittest discover -s . -p "test_*.py"`
- Framework: `unittest`. Some tests use an in-memory / temp `sqlite3` to
  isolate `DatabaseManager` (see `test_numbering.py` for the pattern).
- For Excel-output tests: open the generated `.xlsx` with `openpyxl.load_workbook`,
  assert cell values and `merged_cells.ranges` — do NOT use a real Excel
  process; this codebase stays headless-testable.
- Add a test next to any new `app/*.py` logic. GUI widget changes are usually
  verified manually (no screenshot tests configured).
- All tests must pass before opening a PR.

## PR & commit conventions

- Default branch: **`main`**. Branch from `main`; never push to it directly.
- Commit messages: short Russian-or-English one-liner describing the *what*;
  recent history is mixed (e.g. `Fix approver name encoding in generated
  document`, `070bd91 Исправление нумерации извещений: …`). No strict
  conventional-commits prefix is enforced.
- Open PR via `gh pr create` once tests pass. CI is not configured — local
  test pass is the gate.

## Security

- No secrets in repo. `data/*.db` and `data/*.json` are committed (they hold
  user work product, not credentials).
- `data/app.db.bak-*` files are auto-created on certain operations; prune
  old ones periodically but keep at least a few recent.
- Excel temp files (`~$*.xlsx`) are in `.gitignore` — keep them out.
- Project is single-user / internal; no auth, no network I/O. Don't add any
  without explicit user approval.
