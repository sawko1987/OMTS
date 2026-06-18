# OMTS — Architecture & Data Flow

Quick map for any agent joining the project. Read this once before touching
`app/excel_generator.py` or `app/gui/changes_table_widget.py`.

## Module map

```
main.py
  └─ app.gui.main_window.MainWindow
        ├─ app.gui.document_info_widget.DocumentInfoWidget      (Tab 1: реквизиты)
        ├─ app.gui.changes_table_widget.ChangesTableWidget     (Tab 2: изменения материалов)
        └─ app.gui.replacement_sets_editor_widget.ReplacementSetsEditorWidget  (Tab 3: наборы)
              │
              ├─ app.excel_generator.ExcelGenerator             (writes the final .xlsx)
              │     └─ app.numbering.NumberingManager          (per-month sequence)
              │           └─ app.database.DatabaseManager
              ├─ app.catalog_loader.CatalogLoader               (catalog/catalog.xlsx)
              ├─ app.document_store.DocumentStore
              ├─ app.product_store.ProductStore
              ├─ app.history_store.HistoryStore
              ├─ app.banned_replacements_store.BannedReplacementsStore
              └─ app.settings_manager.SettingsManager
```

## Data flow on "Сгенерировать документ"

1. **Input** — `DocumentInfoWidget` collects `DocumentData` header fields
   (date, validity, products, reason, TKO conclusion).
2. **Material changes** — `ChangesTableWidget` holds a list of
   `PartChanges`, each with `MaterialChange` rows (catalog + is_changed +
   after_*). The user types "after" with help from
   `app/catalog_loader.py` (history suggestions).
3. **Numbering** — `ExcelGenerator.__init__` constructs a
   `NumberingManager`, which checks `data/app.db`; falls back to
   `data/numbering.json` if the DB is absent.
4. **Template load** — `ExcelGenerator` opens
   `templates/Izveshchenie_template.xlsx` via `openpyxl.load_workbook`.
5. **Fill** — header cells, the per-part material table, and the "Вручено"
   block (workshops aggregated from changed materials). All writes go
   through `get_merged_cell_value` to avoid stomping merged ranges.
6. **Pagination** — `_find_table_header_row` + `_page_capacity_points` +
   `_row_height_points` decide whether the data fits on one sheet. If not,
   parts marked with `additional_page_number=N` land on sheet N+1.
7. **Save** — output goes to a user-chosen path (default `output/`) and is
   also recorded in `DocumentStore` (DB) + the `numbering` table.

## Persistence layout (`data/`)

- `app.db` — SQLite: documents, parts, materials, numbering, sets, history,
  banned replacements, products, settings (schema in `app/database.py`).
- `app.db.bak-YYYYMMDD-HHMMSS` — auto-created by restore / migration code;
  prune the old ones occasionally but keep a few recent.
- `history.json` — fallback history store (used when DB is absent).
- `numbering.json` — fallback numbering store (used when DB is absent).
- `settings.json` — UI / app settings (window geometry, last-used values).
- `template_config.json` — cached analysis of
  `templates/Izveshchenie_template.xlsx` (header row, capacities).

## Code-style rules of thumb

- New domain concept → add a dataclass in `app/models.py`.
- New path → add a constant in `app/config.py`; never inline `Path(...)`.
- New user-facing string → keep in Russian; match the surrounding tone.
- New logging → `logger = logging.getLogger(__name__)` at module top.
- New dependency → add a version pin to `requirements.txt` and call it
  out in the commit message.
