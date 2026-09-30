"""Minimal observer-pattern event bus used to notify application events.

Services publish domain events; the GUI (and any other interested party)
subscribes as an Observer. This keeps event handling separate from the GUI
as required by the project architecture guidelines.
"""

from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

from domain.merged_workbook import MergedWorkbook
from domain.worksheet_schema import WorksheetSchema


@dataclass(frozen=True)
class FilesLoaded:
    """Published after the user-selected files/folder have been ingested.

    Attributes:
        file_paths: Excel file paths that were loaded, in ingestion order.
    """

    file_paths: list[Path]


@dataclass(frozen=True)
class RowsMerged:
    """Published after all worksheet rows have been merged by sheet name.

    Attributes:
        merged_workbook: Workbook holding the merged worksheets (schemas
            are not generated yet at this point).
    """

    merged_workbook: MergedWorkbook


@dataclass(frozen=True)
class SchemasGenerated:
    """Published after one schema per merged worksheet was pre-generated.

    Attributes:
        merged_workbook: Workbook whose merged worksheets now carry schemas.
    """

    merged_workbook: MergedWorkbook


@dataclass(frozen=True)
class SchemaEdited:
    """Published after the user edited one worksheet schema in the GUI.

    Attributes:
        sheet_title: Title of the worksheet whose schema was edited.
        schema: The edited schema.
    """

    sheet_title: str
    schema: WorksheetSchema


@dataclass(frozen=True)
class MergeCompleted:
    """Published after the merged output workbook was saved successfully.

    Attributes:
        output_path: Path of the saved output file.
    """

    output_path: Path


@dataclass(frozen=True)
class MergeFailed:
    """Published when the merge pipeline fails with an error.

    Attributes:
        error_message: Human-readable description of the failure.
    """

    error_message: str


@dataclass(frozen=True)
class ValidationWarning:
    """Published for soft validation issues that never stop the merge.

    Attributes:
        message: Human-readable description of the warning.
        sheet_title: Worksheet the warning refers to, when applicable.
    """

    message: str
    sheet_title: str | None = None


# Union of every event type flowing through the EventBus.
MergeEvent = (
    FilesLoaded
    | RowsMerged
    | SchemasGenerated
    | SchemaEdited
    | MergeCompleted
    | MergeFailed
    | ValidationWarning
)


class Observer(Protocol):
    """Interface implemented by any object interested in merge events."""

    def on_event(self, event: MergeEvent) -> None:
        """Handle one published event.

        Args:
            event: The event instance to handle.
        """
        ...


class EventBus:
    """Publish/subscribe hub connecting services (publishers) and observers.

    Observers are notified synchronously in subscription order. GUI
    observers are responsible for marshalling work onto the tkinter main
    thread themselves.
    """

    def __init__(self) -> None:
        """Create an event bus with no subscribers."""
        self._observers: list[Observer] = []

    def subscribe(self, observer: Observer) -> None:
        """Register an observer to receive future events.

        Args:
            observer: Object implementing the Observer protocol.
        """
        if observer not in self._observers:
            self._observers.append(observer)

    def unsubscribe(self, observer: Observer) -> None:
        """Remove a previously registered observer.

        Args:
            observer: Object to stop notifying. No-op if not subscribed.
        """
        if observer in self._observers:
            self._observers.remove(observer)

    def publish(self, event: MergeEvent) -> None:
        """Deliver an event to all subscribed observers, in order.

        Args:
            event: The event instance to deliver.
        """
        for observer in list(self._observers):
            observer.on_event(event)
