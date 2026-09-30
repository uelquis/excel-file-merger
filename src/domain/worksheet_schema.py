"""Defines the schema describing the content structure of a merged worksheet."""

from dataclasses import dataclass, field


@dataclass
class WorksheetSchema:
    """Describes the expected content structure of one merged worksheet.

    The schema holds only content information: the sheet title and the
    ordered list of column names. No type inference is performed; columns
    are treated as plain names that the user may edit freely.

    Attributes:
        sheet_title: Title of the worksheet this schema describes.
        column_names: Ordered column names as they appear in the sheet.
        source_column_indices: Per output column, the index of the source
            column in the merged data rows (None for columns added by the
            user, which are written empty). Kept in sync with column_names
            by every edit method so renames/reorders map back correctly.
    """

    sheet_title: str
    column_names: list[str] = field(default_factory=list)
    source_column_indices: list[int | None] = field(default_factory=list)

    def rename_column(self, column_index: int, new_column_name: str) -> None:
        """Rename the column at the given index.

        Args:
            column_index: Zero-based position of the column to rename.
            new_column_name: New name to assign to the column.

        Raises:
            IndexError: If column_index is out of range.
        """
        self.column_names[column_index] = new_column_name

    def add_column(self, column_name: str, column_index: int | None = None) -> None:
        """Add a new user-defined column to the schema (written empty).

        Args:
            column_name: Name of the column to add.
            column_index: Position to insert the column at. When None the
                column is appended at the end of the schema.
        """
        if column_index is None:
            self.column_names.append(column_name)
            self.source_column_indices.append(None)
        else:
            self.column_names.insert(column_index, column_name)
            self.source_column_indices.insert(column_index, None)

    def remove_column(self, column_index: int) -> str:
        """Remove the column at the given index and return its name.

        Args:
            column_index: Zero-based position of the column to remove.

        Returns:
            The name of the removed column.

        Raises:
            IndexError: If column_index is out of range.
        """
        if column_index < len(self.source_column_indices):
            self.source_column_indices.pop(column_index)
        return self.column_names.pop(column_index)

    def move_column(self, from_index: int, to_index: int) -> None:
        """Move a column from one position to another (reordering).

        Args:
            from_index: Zero-based current position of the column.
            to_index: Zero-based target position of the column.

        Raises:
            IndexError: If either index is out of range.
        """
        moved_column_name = self.column_names.pop(from_index)
        self.column_names.insert(to_index, moved_column_name)
        if from_index < len(self.source_column_indices):
            moved_source_index = self.source_column_indices.pop(from_index)
            self.source_column_indices.insert(to_index, moved_source_index)
