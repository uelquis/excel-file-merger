"""Defines the merged workbook domain object holding merged worksheet data."""

from dataclasses import dataclass, field
from typing import Any

from domain.worksheet_schema import WorksheetSchema


@dataclass
class MergedWorksheet:
    """One merged worksheet: all rows from same-named source worksheets.

    Attributes:
        sheet_title: Title of the merged worksheet (taken from the first
            ingested file containing a worksheet with this name).
        header_row: Column values of the header, from the first file that
            contains this worksheet. May be ragged when source files have
            mismatching headers.
        data_rows: All non-header rows appended from every source file.
        schema: Content schema pre-generated from this worksheet's data.
            It is None until the schema generator runs.
    """

    sheet_title: str
    header_row: list[Any] = field(default_factory=list)
    data_rows: list[list[Any]] = field(default_factory=list)
    schema: WorksheetSchema | None = None

    def append_data_rows(self, rows: list[list[Any]]) -> None:
        """Append data rows coming from one source worksheet.

        Args:
            rows: Non-header rows to append to this merged worksheet.
        """
        self.data_rows.extend(rows)


class MergedWorkbook:
    """Collection of merged worksheets, kept in first-seen sheet order.

    Sheet order follows the first ingested file; sheet names that only
    appear in later files are appended at the end, in order of appearance.
    """

    def __init__(self) -> None:
        """Create an empty merged workbook."""
        self._merged_worksheets: list[MergedWorksheet] = []
        self._merged_worksheets_by_title: dict[str, MergedWorksheet] = {}

    def add_worksheet(self, merged_worksheet: MergedWorksheet) -> None:
        """Register a merged worksheet in the workbook.

        Args:
            merged_worksheet: Worksheet to add; its title must be unique
                within this workbook.

        Raises:
            ValueError: If a worksheet with the same title already exists.
        """
        if merged_worksheet.sheet_title in self._merged_worksheets_by_title:
            raise ValueError(
                f"Worksheet '{merged_worksheet.sheet_title}' already exists in the merged workbook"
            )
        self._merged_worksheets.append(merged_worksheet)
        self._merged_worksheets_by_title[merged_worksheet.sheet_title] = merged_worksheet

    def get_worksheet_by_title(self, sheet_title: str) -> MergedWorksheet | None:
        """Return the merged worksheet with the given title, or None.

        Args:
            sheet_title: Title of the worksheet to look up.
        """
        return self._merged_worksheets_by_title.get(sheet_title)

    @property
    def worksheets(self) -> list[MergedWorksheet]:
        """All merged worksheets, in first-seen sheet order."""
        return list(self._merged_worksheets)

    @property
    def sheet_titles(self) -> list[str]:
        """Titles of all merged worksheets, in first-seen sheet order."""
        return [worksheet.sheet_title for worksheet in self._merged_worksheets]
