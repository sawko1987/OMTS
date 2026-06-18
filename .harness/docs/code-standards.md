# OMTS — Code Standards

Project conventions any agent should follow when touching code in this
repo. Linked from each rein's `agent.md` so the rules live in one place.

## Python version & deps

- Python 3.8+ target.
- Runtime deps pinned in `requirements.txt`: `PySide6>=6.5.0`,
  `openpyxl>=3.1.0`, `xlrd>=2.0.0`. Don't add new deps without an
  explicit user OK; if you do, pin a version.

## Paths

- All file paths come from `app/config.py`. Never hard-code
  `Path("data/...")` or `Path("templates/...")` in business code.
- `app/config.py` auto-creates `data/`, `catalog/`, `templates/`,
  `output/` at import. Don't repeat that elsewhere.

## Logging

- `logger = logging.getLogger(__name__)` at module top.
- `logger.info` for lifecycle events, `logger.warning` for recoverable
  issues, `logger.exception` (or `logger.error(..., exc_info=True)`) inside
  `except` blocks.
- The root logger config is in `main.py::setup_logging`. Don't override
  handlers in other modules.

## Domain models

- Dataclasses only (`@dataclass`); don't introduce `pydantic` or an ORM.
- New domain concept → new dataclass in `app/models.py`, with `Optional`
  foreign-key IDs as `Optional[int] = None`.

## Excel generation rules (`app/excel_generator.py`)

- Never write into a possibly-merged cell with `ws.cell(r, c)` — use
  `get_merged_cell_value(ws, r, c)` instead.
- A4 is `paperSize=9`; pagination math assumes it. When changing
  `_paper_height_inches` or `_page_capacity_points`, regenerate a real
  `.xlsx` and re-open it with `openpyxl` to confirm row counts.
- `additional_page_number` is the seam where data flows onto sheet N+1.
  Tests must cover both sides of that boundary.
- Hidden rows: don't treat `row_dimensions[row].hidden` as 0 height —
  `_row_height_points` documents why.

## GUI rules (`app/gui/`)

- Match each widget's existing signal/slot naming and layout.
- All UI strings in Russian. Don't introduce a translation framework —
  the project is single-language by design.
- Don't allocate a `QApplication` in tests; the suite stays headless.
  Widget changes are verified manually.
- Window geometry / last-used values go in `data/settings.json` via
  `SettingsManager`.

## Database rules (`app/database.py`)

- Schema changes need a migration in `app/migrate_to_sqlite.py`.
- Additive changes only — never destructive renames or column drops in
  `data/app.db`; users have real data there.
- `NumberingManager`'s JSON fallback must still work on a fresh install
  with no DB. Test that path explicitly.

## Tests

- `unittest` at the repo root. New behavior → new `test_*.py` next to
  the existing ones.
- DB-touching tests use a real `sqlite3` file in a tempdir; see
  `test_numbering.py` for the `TempDatabaseManager` pattern. Reuse it.
- Excel-output tests open the generated file in a *separate*
  `openpyxl.load_workbook` call (not the generator's writer) to confirm
  what's actually on disk.
- No `pytest`, no `pytest-qt`, no screenshot tests.

## Commit messages

- Short one-liner (Russian or English) describing the *what*. No strict
  conventional-commits prefix is enforced — match recent history.
- Reference the touched area: `excel_generator:`, `numbering:`,
  `gui/main_window:`, etc., when the diff is large.
