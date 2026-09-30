"""Service that ingests excel files and merges worksheet rows by sheet name."""

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from domain.merged_workbook import MergedWorkbook, MergedWorksheet
from events.event_bus import EventBus, FilesLoaded, RowsMerged, ValidationWarning


class FileIngestionService:
    """Loads user-selected files/folders and merges their worksheets by name.

    Merging is order-independent: every worksheet with the same name across
    all ingested files is merged into a single result worksheet (for example
    file1: a,b + file2: b,a -> a1+a2, b1+b2). Validation is soft: any
    inconsistency publishes a ValidationWarning event and the merge proceeds.

    Composed with an EventBus instead of owning observers directly.
    """

    EXCEL_FILE_SUFFIX = ".xlsx"

    def __init__(self, event_bus: EventBus) -> None:
        """Create the ingestion service.

        Args:
            event_bus: Bus used to publish FilesLoaded, RowsMerged and
                ValidationWarning events.
        """
        self._event_bus = event_bus

    def collect_excel_file_paths(self, selected_paths: list[Path]) -> list[Path]:
        """Expand the user selection into an ordered list of .xlsx files.

        Folders are expanded into the .xlsx files they contain; individual
        files are kept as-is. Selection order is preserved and duplicates
        are removed.

        Args:
            selected_paths: Files and/or folders chosen by the user.

        Returns:
            Ordered, deduplicated list of .xlsx file paths.
        """
        collected_file_paths: list[Path] = []
        for selected_path in selected_paths:
            if selected_path.is_dir():
                collected_file_paths.extend(
                    sorted(selected_path.glob(f"*{self.EXCEL_FILE_SUFFIX}"))
                )
            elif selected_path.suffix.lower() == self.EXCEL_FILE_SUFFIX:
                collected_file_paths.append(selected_path)
            else:
                self._event_bus.publish(
                    ValidationWarning(
                        f"Skipped '{selected_path.name}': not a {self.EXCEL_FILE_SUFFIX} file"
                    )
                )

        unique_file_paths: list[Path] = []
        for file_path in collected_file_paths:
            if file_path not in unique_file_paths:
                unique_file_paths.append(file_path)

        self._event_bus.publish(FilesLoaded(file_paths=unique_file_paths))
        return unique_file_paths

    def merge_worksheet_rows(self, file_paths: list[Path]) -> MergedWorkbook:
        """Merge the rows of all worksheets, grouped by sheet name.

        For each unique sheet name across all files, a MergedWorksheet is
        created holding the header from the first file containing that sheet
        and the data rows appended from every file. Result sheet order
        follows the first file; names appearing only in later files are
        appended in order of appearance.

        Soft validations published as ValidationWarning events:
        - a sheet name present in only some files (still merged);
        - same sheet name with different headers (first file's header wins);
        - a file that could not be opened (skipped).

        Args:
            file_paths: Ordered list of .xlsx files to merge.

        Returns:
            MergedWorkbook with one merged worksheet per unique sheet name
            (schemas are not generated yet).

        Raises:
            ValueError: If no excel file was provided.
        """
        if not file_paths:
            raise ValueError("No excel files provided to merge")

        merged_workbook = MergedWorkbook()
        sheet_presence_by_title: dict[str, list[str]] = {}

        for file_path in file_paths:
            sheets_in_file = self._read_file_sheets(file_path)
            if sheets_in_file is None:
                continue

            for sheet_title, header_row, data_rows in sheets_in_file:
                sheet_presence_by_title.setdefault(sheet_title, []).append(file_path.name)
                existing_merged_worksheet = merged_workbook.get_worksheet_by_title(sheet_title)

                if existing_merged_worksheet is None:
                    merged_workbook.add_worksheet(
                        MergedWorksheet(
                            sheet_title=sheet_title,
                            header_row=header_row,
                            data_rows=data_rows,
                        )
                    )
                    continue

                if header_row != existing_merged_worksheet.header_row:
                    self._event_bus.publish(
                        ValidationWarning(
                            f"Sheet '{sheet_title}' in '{file_path.name}' has a different header "
                            "than the first file containing it; keeping the first header",
                            sheet_title=sheet_title,
                        )
                    )
                existing_merged_worksheet.append_data_rows(data_rows)

        self._warn_about_partially_present_sheets(sheet_presence_by_title, file_paths)

        self._event_bus.publish(RowsMerged(merged_workbook=merged_workbook))
        return merged_workbook

    def _read_file_sheets(
        self, file_path: Path
    ) -> list[tuple[str, list, list[list]]] | None:
        """Open one workbook, read all sheets' rows, then close it.

        Args:
            file_path: Path of the .xlsx file to read.

        Returns:
            List of (sheet_title, header_row, data_rows) tuples per worksheet,
            or None if the file could not be opened.
        """
        try:
            workbook = load_workbook(file_path, read_only=True)
        except Exception as load_error:
            self._event_bus.publish(
                ValidationWarning(
                    f"Skipped '{file_path.name}': could not be opened ({load_error})"
                )
            )
            return None

        try:
            sheets_in_file: list[tuple[str, list, list[list]]] = []
            for source_worksheet in workbook.worksheets:
                sheets_in_file.append(
                    self._read_header_and_data_rows(source_worksheet)
                )
            return sheets_in_file
        finally:
            workbook.close()

    def _read_header_and_data_rows(
        self, source_worksheet: Worksheet
    ) -> tuple[str, list, list[list]]:
        """Split a source worksheet into its title, header row and data rows.

        Args:
            source_worksheet: Worksheet to read from top to bottom.

        Returns:
            Tuple of (sheet_title, header_row, data_rows); header_row is empty
            when the worksheet has no rows at all.
        """
        all_rows = [list(row) for row in source_worksheet.iter_rows(values_only=True)]
        if not all_rows:
            return source_worksheet.title, [], []
        return source_worksheet.title, all_rows[0], all_rows[1:]

    def _warn_about_partially_present_sheets(
        self, sheet_presence_by_title: dict[str, list[str]], file_paths: list[Path]
    ) -> None:
        """Warn about sheet names that are not present in every ingested file.

        Args:
            sheet_presence_by_title: Sheet title -> names of files containing it.
            file_paths: All files that were merged.
        """
        total_file_count = len(file_paths)
        for sheet_title, present_file_names in sheet_presence_by_title.items():
            if len(present_file_names) < total_file_count:
                self._event_bus.publish(
                    ValidationWarning(
                        f"Sheet '{sheet_title}' only found in {', '.join(present_file_names)}; "
                        "merged with the rows available",
                        sheet_title=sheet_title,
                    )
                )
