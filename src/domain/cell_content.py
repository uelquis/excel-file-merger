"""Value object pairing a cell's value with its preserved background color."""

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class CellContent:
    """Immutable content of one worksheet cell.

    Carries the cell's value plus the RGB solid background color captured at
    ingestion time, so colors survive the merge and are re-applied on output.
    Background color is treated purely as presentation: it never counts as
    content when deciding whether a row/column is blank.

    Attributes:
        value: The raw cell value (None when the cell is empty).
        background_color: RGB hex string of the solid fill (for example
            'FFFF0000' or 'FF0000'), or None when the cell has no captured
            RGB background. Theme-palette colors are not captured (None).
    """

    value: Any = None
    background_color: str | None = None

    def has_value(self) -> bool:
        """Return True when the cell carries a non-empty value.

        A value is considered present when it is not None and, for strings,
        not empty after stripping surrounding whitespace.
        """
        if self.value is None:
            return False
        if isinstance(self.value, str):
            return self.value.strip() != ""
        return True

    @staticmethod
    def from_value(value: Any) -> "CellContent":
        """Build a colorless CellContent from a plain value.

        Convenience factory used when constructing rows without color
        information (for example empty placeholder columns).

        Args:
            value: Raw cell value to wrap.
        """
        return CellContent(value=value, background_color=None)
