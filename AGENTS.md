# Excel File Merger

## Project overview
This application assists users that need to merge the worksheets in multiple excel files.

## Project Stack

- python 3.14
- openpyxl
- pyinstaller
- tkinter

## Build + Testing
We use Just to run commands as an alternative to make:
    - just build, for building an executable

## Architecture
- Use object oriented programming to implement the aplication. 
- Favor composition rather than inheritance.
- Use Observer Design Pattern to handles events separately from the gui

## Boundaries & Guidelines
- Do not `AGENTS.md` or `justfile` without asking
- The names of classes, functions and variables must be explicit about what they do.
- Use docstring and comments for class and function definitions
- Use python typehints

## Features
- users can select excel files or a folder from the file explorer to feed the application with data.
- the application uses the data to pre generate a schema, which can be changed by the user.
    - if there are multiple worksheets, there are multiple schemas
- the application generates an excel file with all rows from the ingested files.

## Implementation Plan
Flow: select files/folder -> merge worksheet rows **by sheet name** (order-independent, all warnings soft)
-> pre-generate one content-based schema per merged sheet -> user edits schemas in tkinter GUI
-> apply schemas + save output.

Example: file1: a,b ; file2: b,a -> merged result: a1+a2, b1+b2 (paired by name, not position).

### Step 1 - Domain model + observer infrastructure [x]
- `src/domain/worksheet_schema.py` - `WorksheetSchema`: sheet title + ordered column names (content only, no type inference).
- `src/domain/merged_workbook.py` - `MergedWorkbook`: merged rows per sheet, each paired with its schema; sheet titles/order from the first file, extra names appended.
- `src/events/event_bus.py` - minimal `EventBus`/`Observer` (observer pattern); events: `FilesLoaded`, `RowsMerged`, `SchemasGenerated`, `SchemaEdited`, `MergeCompleted`, `MergeFailed`, `ValidationWarning`.

### Step 2 - Ingestion + row merging by sheet name [x]
- `src/services/file_ingestion_service.py` - accepts file list or folder; groups worksheets across files **by name**; header from first file containing the sheet, data rows appended from all.
- Soft validation via `ValidationWarning` (never raises):
    - sheet name in only some files -> still merged into output, warn;
    - same name, different headers -> proceed with first file's header, warn.

### Step 3 - Schema pre-generation [x]
- `src/services/schema_generator.py` - scans each merged sheet's header/content -> `WorksheetSchema` per unique sheet name; emits `SchemasGenerated`.

### Step 4 - Tkinter GUI [x]
- `src/gui/main_window.py` - file/folder pickers (`filedialog`), file list, Merge button; merge runs on a background thread; GUI only subscribes to `EventBus` and delegates to services (composition, no logic in widgets).
- `src/gui/schema_editor_view.py` - one tab per merged sheet: rename/reorder/add/remove columns; edits emit `SchemaEdited`.
- Warnings panel surfacing `ValidationWarning` events.
- `src/gui_main.py` - new GUI entry point.

### Step 5 - Output, CLI deprecation, packaging [x]
- `src/services/workbook_writer.py` - applies edited schemas (column selection/order/renaming, optional `Formatter` styling) to merged rows; saves result workbook.
- Deprecate CLI: `src/app.py`, `src/excel_merger.py`, `src/formatter.py` kept untouched as unused code.
- Update `excelmerger.spec` for GUI entry + tkinter bundling; verify with `just build` and a manual run against `test_input/`.

### Verification
- Unit-check name-based grouping with `test_input/` files (different sheet orders, unmatched names, header mismatches) -> confirm warnings emitted, merged output contains all rows, schemas editable, saved file correct.

## Phase 2 - Blank removal + background color preservation
Goal: merged result must NOT contain blank rows/columns, and cell background
colors (RGB) must be preserved end-to-end.

Agreed rules:
- Blank = no value. Background color never counts as content -> colored-but-empty rows are removed.
- A column is dropped only when its header cell is empty AND all its data cells are empty; any column with a header stays.
- RGB solid fills only; theme-palette colors are ignored (treated as colorless).
- Blank columns are dropped before schema generation (never shown in the editor).

### Step 6 - Domain: cell content with background color [x]
- New `src/domain/cell_content.py` - `CellContent` value object: `value: Any` + `background_color: str | None` (RGB hex).
- `MergedWorksheet.header_row` / `data_rows` switch from raw values to `CellContent`.
- `SchemaGenerator` reads through `.value` (schemas stay content-only; column-count detection unchanged).

### Step 7 - Ingestion: read fills, drop blanks [x]
- `FileIngestionService` stops using `values_only=True`; reads each cell's value plus solid-fill RGB (`fill.fill_type == 'solid'` and `start_color.rgb` is a string; theme colors -> None).
- Blank rows: dropped per source worksheet during ingestion (all cell values None/empty-string).
- Blank columns: computed per merged worksheet after all files merge -> drop only when header empty + all data empty; rewrite rows so indices stay consistent before schema generation.
- Header-mismatch validation compares values only (colors may differ across files without triggering warnings).
- `ValidationWarning` events report dropped blank rows/columns counts per sheet.

### Step 8 - Writer: apply colors + verification [x]
- `WorkbookWriter` writes each cell's value and applies `PatternFill(solid, background_color)` when present; fills remap through `source_column_indices` alongside values (reorder/rename/remove keep colors aligned); user-added columns written colorless.
- Verify with a synthetic colored workbook (blank row, blank column, colored-empty row, named-but-dataless column) + `test_input/`: blanks gone, header columns kept, colors correct after schema edits; `just build` green; exe smoke-run.

### Phase 2 Verification
- Confirm blank rows/columns removed, header-only columns kept, RGB background colors preserved at correct positions after edits, and packaged GUI still builds and launches.

## Phase 3 Pre-Plan (generic — to be refined after your cache study)

### Step 9 — Cache strategy definition (design decisions)
Pin down the answers that shape everything else; your study should inform these:
- What is cached (granularity): per-source-file ingestion results vs. merged workbook vs. both.
- Cache key & invalidation: how entries are identified (content hash, path+mtime, user-chosen?) and when they become stale.
- Lifetime & eviction: session-only vs. persistent; manual removal vs. LRU/size caps.
- Storage: location (app-data vs. project-local) and format (JSON, pickle, sqlite, ...).
- Output: a concrete Phase 3 plan (replacing this pre-plan in the AGENTS.md tracker, with your approval).

### Step 10 — Cache core behind an interface
- Domain types for the cached payload + entry metadata (serializable, color/datetime-safe).
- A CacheService-style component (composition + observer events per AGENTS.md): store, lookup, list, remove, clear — isolated from merging logic so the strategy chosen in Step 9 only affects this layer.
- New cache events on the EventBus (stored/hit/refreshed/removed + soft-failure warnings for corrupt/stale entries).

### Step 11 — Integration + verification
- Wire the cache into the ingestion/merge pipeline (cached and fresh sources merge together; pipeline stays order- and name-based as today).
- GUI surface: whatever selection/management UX Step 9 dictates (e.g., cached-entries panel, status feedback on hits).
- Verification: round-trip fidelity (values, colors, types), correctness of re-merge with mixed cached+new sources, invalidation behavior, just build + exe smoke run.