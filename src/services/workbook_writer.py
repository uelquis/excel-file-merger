"""Service that applies edited schemas to merged rows and saves the output."""

from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from domain.cell_content import CellContent
from domain.merged_workbook import MergedWorkbook, MergedWorksheet
from domain.worksheet_schema import WorksheetSchema
from events.event_bus import EventBus, MergeCompleted


class WorkbookWriter:
    """Writes the merged workbook to disk, honoring the user-edited schemas.

    For each merged worksheet the output sheet receives the schema's column
    names as header and every merged data row remapped through the schema's
    source column indices. This applies the user's column selection, ordering
    and renaming. Each cell's preserved RGB background color is re-applied as
    a solid PatternFill, remapped alongside its value so colors stay aligned
    through edits; columns added by the user are written empty and colorless.
    An optional formatter object (anything exposing ``apply(worksheet)``) can
    style each output worksheet.
    """

    def __init__(self, event_bus: EventBus, formatter: Any | None = None) -> None:
        """Create the writer.

        Args:
            event_bus: Bus used to publish the MergeCompleted event.
            formatter: Optional styling object with an ``apply(worksheet)``
                method (for example the legacy Formatter); None disables styling.
        """
        self._event_bus = event_bus
        self._formatter = formatter

    def write_workbook(self, merged_workbook: MergedWorkbook, output_path: Path) -> Path:
        """Apply schemas, save the result workbook and publish MergeCompleted.

        Existing output files are overwritten. Parent folders are created
        when missing.

        Args:
            merged_workbook: Merged workbook with (possibly edited) schemas.
            output_path: Destination .xlsx path.

        Returns:
            The output_path the workbook was saved to.
        """
        output_workbook = Workbook()
        try:
            output_workbook.remove(output_workbook.active)  # type: ignore[arg-type]
        except Exception:
            pass

        for merged_worksheet in merged_workbook.worksheets:
            self._write_worksheet(merged_worksheet, output_workbook)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_workbook.save(output_path)

        self._event_bus.publish(MergeCompleted(output_path=output_path))
        return output_path

    def _write_worksheet(
        self, merged_worksheet: MergedWorksheet, output_workbook: Workbook
    ) -> None:
        """Write one merged worksheet into the output workbook.

        Args:
            merged_worksheet: Merged worksheet (schema optional; when absent
                the raw header and rows are written unchanged, colors included).
            output_workbook: Workbook to create the output sheet in.
        """
        schema = merged_worksheet.schema
        sheet_title = schema.sheet_title if schema is not None else merged_worksheet.sheet_title
        output_worksheet = output_workbook.create_sheet(title=sheet_title)

        if schema is None:
            self._write_cell_row(output_worksheet, merged_worksheet.header_row)
            for data_row in merged_worksheet.data_rows:
                self._write_cell_row(output_worksheet, data_row)
        else:
            self._write_cell_row(
                output_worksheet, self._build_schema_header_cells(merged_worksheet, schema)
            )
            for data_row in merged_worksheet.data_rows:
                self._write_cell_row(
                    output_worksheet,
                    self._map_data_row_to_schema(data_row, schema.source_column_indices),
                )

        if self._formatter is not None:
            self._formatter.apply(output_worksheet)

    def _write_cell_row(
        self, output_worksheet: Worksheet, cells: list[CellContent]
    ) -> None:
        """Append one row of cells, writing values and re-applying fills.

        Args:
            output_worksheet: Sheet to append the row to.
            cells: Row of CellContent to write; empty rows are skipped.
        """
        if not cells:
            return
        output_worksheet.append([cell.value for cell in cells])
        written_row = output_worksheet[output_worksheet.max_row]
        for column_offset, cell in enumerate(cells):
            if cell.background_color is not None:
                written_row[column_offset].fill = PatternFill(
                    fill_type="solid",
                    start_color=cell.background_color,
                    end_color=cell.background_color,
                )

    def _build_schema_header_cells(
        self, merged_worksheet: MergedWorksheet, schema: WorksheetSchema
    ) -> list[CellContent]:
        """Build the output header: edited names over remapped header colors.

        The header cells are remapped through the schema's source indices (so
        background colors follow any reorder), then each value is replaced by
        the user's edited column name.

        Args:
            merged_worksheet: Worksheet providing the original header cells.
            schema: Edited schema supplying the output column names/order.

        Returns:
            Header cells with schema names as values and preserved colors.
        """
        remapped_header_cells = self._map_data_row_to_schema(
            merged_worksheet.header_row, schema.source_column_indices
        )
        return [
            CellContent(value=column_name, background_color=header_cell.background_color)
            for column_name, header_cell in zip(schema.column_names, remapped_header_cells)
        ]

    def _map_data_row_to_schema(
        self, data_row: list[CellContent], source_column_indices: list[int | None]
    ) -> list[CellContent]:
        """Remap one merged row into the edited schema's column layout.

        Both the value and the background color travel together, so reorders,
        renames and removals keep colors aligned with their data.

        Args:
            data_row: Raw merged row cells.
            source_column_indices: Per output column, the source column index
                (None for user-added columns, written empty and colorless).

        Returns:
            Row cells in the schema's column order; missing or added columns
            become empty colorless CellContent.
        """
        mapped_row: list[CellContent] = []
        for source_column_index in source_column_indices:
            if (
                source_column_index is None
                or source_column_index >= len(data_row)
                or source_column_index < -len(data_row)
            ):
                mapped_row.append(CellContent())
            else:
                mapped_row.append(data_row[source_column_index])
        return mapped_row
