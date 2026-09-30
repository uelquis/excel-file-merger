"""Service that applies edited schemas to merged rows and saves the output."""

from pathlib import Path
from typing import Any

from openpyxl import Workbook

from domain.merged_workbook import MergedWorkbook, MergedWorksheet
from events.event_bus import EventBus, MergeCompleted


class WorkbookWriter:
    """Writes the merged workbook to disk, honoring the user-edited schemas.

    For each merged worksheet the output sheet receives the schema's column
    names as header and every merged data row remapped through the schema's
    source column indices. This applies the user's column selection,
    ordering and renaming; columns added by the user are written empty.
    An optional formatter object (anything exposing ``apply(worksheet)``)
    can style each output worksheet.
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
                the raw header and rows are written unchanged).
            output_workbook: Workbook to create the output sheet in.
        """
        schema = merged_worksheet.schema
        sheet_title = schema.sheet_title if schema is not None else merged_worksheet.sheet_title
        output_worksheet = output_workbook.create_sheet(title=sheet_title)

        if schema is None:
            output_worksheet.append(merged_worksheet.header_row)
            for data_row in merged_worksheet.data_rows:
                output_worksheet.append(data_row)
        else:
            output_worksheet.append(list(schema.column_names))
            for data_row in merged_worksheet.data_rows:
                output_worksheet.append(
                    self._map_data_row_to_schema(data_row, schema.source_column_indices)
                )

        if self._formatter is not None:
            self._formatter.apply(output_worksheet)

    def _map_data_row_to_schema(
        self, data_row: list[Any], source_column_indices: list[int | None]
    ) -> list[Any]:
        """Remap one merged data row into the edited schema's column layout.

        Args:
            data_row: Raw merged row values.
            source_column_indices: Per output column, the source column index
                (None for user-added columns, written empty).

        Returns:
            Row values in the schema's column order; missing or added
            columns become None.
        """
        mapped_row: list[Any] = []
        for source_column_index in source_column_indices:
            if (
                source_column_index is None
                or source_column_index >= len(data_row)
                or source_column_index < -len(data_row)
            ):
                mapped_row.append(None)
            else:
                mapped_row.append(data_row[source_column_index])
        return mapped_row
