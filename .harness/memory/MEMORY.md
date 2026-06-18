# OMTS — Project Memory

Shared notes that any rein should be able to read. Add durable findings
here (e.g. "pagination regression in v1.4 — root cause was X"); avoid
session-specific scratch.

## Hot files (read these before touching)

- `app/excel_generator.py` — 1286 lines; pagination math is the most
  regression-prone code in the repo. Read
  `_find_table_header_row`, `_page_capacity_points`, `_row_height_points`,
  and `get_merged_cell_value` before changing anything in the fill
  pipeline.
- `app/gui/changes_table_widget.py` — 87 KB; the central editing surface.
  Match its signal/slot conventions.
- `app/gui/replacement_sets_editor_widget.py` — 64 KB; the "Наборы" tab.
- `app/gui/main_window.py` — 44 KB; the window assembly point.
- `app/database.py` — schema is the source of truth for `data/app.db`;
  any column change needs a migration.

## Landmines (things that bit us already)

- `data/app.db.bak-*` is auto-generated on certain operations; if you see
  a fresh one in `git status` you almost certainly want to keep it
  out of the commit (`*.bak-*` is in `.gitignore` — make sure new
  patterns get added there too if needed).
- `app/numbering.py` has a JSON-fallback path. Don't "clean it up" — it
  runs on a fresh install before the DB exists, and the `migrate_to_sqlite`
  flow relies on it.
- `output/` is committed (with `.gitkeep`) on purpose — generated
  documents are part of the deliverable record. Don't add it to
  `.gitignore`.

## Open follow-ups (suggested by current code)

- No CI is configured. Tests are run manually via
  `python -m unittest discover -s . -p "test_*.py"`.
- No `pyproject.toml` / `setup.py` — install is `pip install -r
  requirements.txt` only. Don't introduce a packaging step without
  asking.
- GUI widget changes have no automated tests. If a widget grows complex
  enough to need tests, prefer extracting its logic into a plain Python
  class in `app/` (not in `app/gui/`) so it can be unit-tested headless.
