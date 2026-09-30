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