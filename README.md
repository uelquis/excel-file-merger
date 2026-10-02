# Excel File Merger

**English** | [Português](README.pt-BR.md)

## Overview

Excel File Merger is a desktop application that merges the worksheets of
multiple `.xlsx` files into a single workbook.

## Features

- **File or folder selection** — add individual `.xlsx` files or an entire
  folder through the file explorer.
- **Name-based sheet merging** — same-named worksheets across files are
  merged into one sheet; sheet order follows the first file, extra names are
  appended.
- **Background color preservation** — RGB cell background colors are carried
  through the merge and stay aligned with their data even after schema edits.
- **Editable schemas** — the application pre-generates one schema per merged
  sheet; rename, reorder, add or remove columns in the built-in editor
  before saving.
- **Soft validation** — data inconsistencies (sheet present in only some
  files, header mismatches, dropped blanks) appear as warnings in a dedicated
  panel; the merge is never blocked by them.

## How it works

1. **Select** — add `.xlsx` files or a folder.
2. **Merge** — worksheets are paired by name and their rows combined;
   blank rows/columns are removed and cell colors captured.
3. **Review schemas** — one tab per merged sheet lets you rename, reorder,
   add or remove columns.
4. **Save** — the edited schemas are applied and the merged workbook is
   written to the output `.xlsx` of your choice.

## Prerequisites

- Python 3.14+
- [uv](https://docs.astral.sh/uv/)
- [just](https://github.com/casey/just)

## Build & run

```bash
# install dependencies
uv sync

# build the executable (output: dist/)
just build

# run from source (development)
uv run python src/gui_main.py
```
