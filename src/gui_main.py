"""GUI entry point: wires services, event bus and main window together."""

import tkinter as tk

from events.event_bus import EventBus
from gui.main_window import MainWindow
from services.file_ingestion_service import FileIngestionService
from services.schema_generator import SchemaGenerator
from services.workbook_writer import WorkbookWriter


def build_main_window(root: tk.Tk) -> MainWindow:
    """Compose the application objects and return the main window.

    Args:
        root: Tkinter root window.

    Returns:
        Fully wired MainWindow ready to be displayed.
    """
    event_bus = EventBus()
    file_ingestion_service = FileIngestionService(event_bus)
    schema_generator = SchemaGenerator(event_bus)
    workbook_writer = WorkbookWriter(event_bus)
    return MainWindow(
        root=root,
        event_bus=event_bus,
        file_ingestion_service=file_ingestion_service,
        schema_generator=schema_generator,
        workbook_writer=workbook_writer,
    )


def main() -> None:
    """Create the root window and start the tkinter main loop."""
    root = tk.Tk()
    build_main_window(root)
    root.mainloop()


if __name__ == "__main__":
    main()
