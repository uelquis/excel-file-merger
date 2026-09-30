"""Service that pre-generates one content-based schema per merged worksheet."""

from typing import Any

from domain.merged_workbook import MergedWorkbook, MergedWorksheet
from domain.worksheet_schema import WorksheetSchema
from events.event_bus import EventBus, SchemasGenerated


class SchemaGenerator:
    """Scans merged worksheets and pre-generates editable schemas.

    Schemas are content-based: column names come from the merged header
    row (no type inference). When the header is missing values or data
    rows are wider than the header (possible after soft header-mismatch
    merges), placeholder names are generated so every column stays
    editable by the user in the GUI.
    """

    def __init__(self, event_bus: EventBus) -> None:
        """Create the schema generator.

        Args:
            event_bus: Bus used to publish the SchemasGenerated event.
        """
        self._event_bus = event_bus

    def generate_schemas(self, merged_workbook: MergedWorkbook) -> MergedWorkbook:
        """Pre-generate one WorksheetSchema per merged worksheet.

        Args:
            merged_workbook: Workbook with merged rows (output of the
                ingestion service).

        Returns:
            The same workbook, with each worksheet's schema populated.
        """
        for merged_worksheet in merged_workbook.worksheets:
            merged_worksheet.schema = self._generate_schema_for_worksheet(merged_worksheet)

        self._event_bus.publish(SchemasGenerated(merged_workbook=merged_workbook))
        return merged_workbook

    def _generate_schema_for_worksheet(
        self, merged_worksheet: MergedWorksheet
    ) -> WorksheetSchema:
        """Build the schema of a single merged worksheet from its content.

        Args:
            merged_worksheet: Worksheet whose header and data rows are scanned.

        Returns:
            WorksheetSchema with one column name per detected column.
        """
        column_count = self._detect_column_count(merged_worksheet)
        column_names = [
            self._column_name_from_header(merged_worksheet.header_row, column_index)
            for column_index in range(column_count)
        ]
        return WorksheetSchema(
            sheet_title=merged_worksheet.sheet_title,
            column_names=self._deduplicate_column_names(column_names),
            source_column_indices=list(range(column_count)),
        )

    def _detect_column_count(self, merged_worksheet: MergedWorksheet) -> int:
        """Detect how many columns the worksheet content spans.

        Uses the widest of the header row and all data rows so ragged
        content (from header mismatches) does not lose columns.

        Args:
            merged_worksheet: Worksheet to inspect.

        Returns:
            Number of columns detected (0 for an empty worksheet).
        """
        widest_data_row_length = max(
            (len(data_row) for data_row in merged_worksheet.data_rows), default=0
        )
        return max(len(merged_worksheet.header_row), widest_data_row_length)

    def _column_name_from_header(self, header_row: list[Any], column_index: int) -> str:
        """Derive one column name from the header row.

        Args:
            header_row: Header values of the merged worksheet.
            column_index: Zero-based index of the column to name.

        Returns:
            The stripped header value as text, or a 'ColumnN' placeholder
            when the header is missing or the value is empty.
        """
        if column_index < len(header_row):
            header_value = header_row[column_index]
            if header_value is not None and str(header_value).strip() != "":
                return str(header_value).strip()
        return f"Column{column_index + 1}"

    def _deduplicate_column_names(self, column_names: list[str]) -> list[str]:
        """Ensure column names are unique by appending numeric suffixes.

        Args:
            column_names: Raw column names, possibly containing duplicates.

        Returns:
            Names where each repeated occurrence gets a ' (2)', ' (3)', ...
            suffix; first occurrences stay unchanged.
        """
        occurrence_count_by_name: dict[str, int] = {}
        deduplicated_column_names: list[str] = []
        for column_name in column_names:
            occurrence_count = occurrence_count_by_name.get(column_name, 0) + 1
            occurrence_count_by_name[column_name] = occurrence_count
            if occurrence_count == 1:
                deduplicated_column_names.append(column_name)
            else:
                deduplicated_column_names.append(f"{column_name} ({occurrence_count})")
        return deduplicated_column_names
