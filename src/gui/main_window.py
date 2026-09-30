"""Main application window: file selection, merge trigger and warnings panel."""

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from domain.merged_workbook import MergedWorkbook
from events.event_bus import (
    EventBus,
    FilesLoaded,
    MergeCompleted,
    MergeEvent,
    MergeFailed,
    SchemasGenerated,
    ValidationWarning,
)
from gui.schema_editor_view import SchemaEditorView
from services.file_ingestion_service import FileIngestionService
from services.schema_generator import SchemaGenerator
from services.workbook_writer import WorkbookWriter


class MainWindow:
    """Primary window of the application.

    Composes the ingestion, schema generation and workbook writer services
    and observes the EventBus. Contains no merging logic itself: button
    actions delegate to services, and service results arrive back as events.
    Merge and save run on background threads; events are marshalled onto the
    tkinter main loop with ``root.after`` because tkinter is not thread-safe.
    """

    def __init__(
        self,
        root: tk.Tk,
        event_bus: EventBus,
        file_ingestion_service: FileIngestionService,
        schema_generator: SchemaGenerator,
        workbook_writer: WorkbookWriter,
    ) -> None:
        """Build the main window and subscribe it to the event bus.

        Args:
            root: Tkinter root window to populate.
            event_bus: Bus this window observes and services publish to.
            file_ingestion_service: Service used to collect files and merge rows.
            schema_generator: Service used to pre-generate merged sheet schemas.
            workbook_writer: Service used to save the merged output workbook.
        """
        self._root = root
        self._event_bus = event_bus
        self._file_ingestion_service = file_ingestion_service
        self._schema_generator = schema_generator
        self._workbook_writer = workbook_writer
        self._selected_excel_file_paths: list[Path] = []
        self._last_merged_workbook: MergedWorkbook | None = None
        self._schema_editor_view: SchemaEditorView | None = None

        self._event_bus.subscribe(self)

        self._root.title("Excel File Merger")
        self._root.geometry("720x520")
        self._build_widgets()

    def _build_widgets(self) -> None:
        """Create all widgets of the main window."""
        root_frame = ttk.Frame(self._root, padding=10)
        root_frame.pack(fill=tk.BOTH, expand=True)

        root_frame.columnconfigure(0, weight=1)
        root_frame.rowconfigure(1, weight=1)
        root_frame.rowconfigure(2, weight=1)

        self._build_file_selection_buttons(root_frame)
        self._build_file_list(root_frame)
        self._build_warnings_panel(root_frame)
        self._build_merge_button(root_frame)

    def _build_file_selection_buttons(self, parent: ttk.Frame) -> None:
        """Create the buttons used to feed the application with files.

        Args:
            parent: Frame the buttons row is packed into.
        """
        buttons_frame = ttk.Frame(parent)
        buttons_frame.grid(row=0, column=0, sticky=tk.EW)

        self._add_files_button = ttk.Button(
            buttons_frame, text="Add Files...", command=self._on_add_files_clicked
        )
        self._add_files_button.pack(side=tk.LEFT)

        self._add_folder_button = ttk.Button(
            buttons_frame, text="Add Folder...", command=self._on_add_folder_clicked
        )
        self._add_folder_button.pack(side=tk.LEFT, padx=(8, 0))

        self._remove_selected_button = ttk.Button(
            buttons_frame, text="Remove Selected", command=self._on_remove_selected_clicked
        )
        self._remove_selected_button.pack(side=tk.LEFT, padx=(8, 0))

        self._clear_files_button = ttk.Button(
            buttons_frame, text="Clear", command=self._on_clear_files_clicked
        )
        self._clear_files_button.pack(side=tk.LEFT, padx=(8, 0))

    def _build_file_list(self, parent: ttk.Frame) -> None:
        """Create the scrollable list showing the selected excel files.

        Args:
            parent: Frame the list is gridded into.
        """
        file_list_frame = ttk.LabelFrame(parent, text="Selected files", padding=5)
        file_list_frame.grid(row=1, column=0, sticky=tk.NSEW, pady=(10, 0))
        file_list_frame.columnconfigure(0, weight=1)
        file_list_frame.rowconfigure(0, weight=1)

        self._file_listbox = tk.Listbox(file_list_frame, selectmode=tk.EXTENDED)
        self._file_listbox.grid(row=0, column=0, sticky=tk.NSEW)

        file_list_scrollbar = ttk.Scrollbar(
            file_list_frame, orient=tk.VERTICAL, command=self._file_listbox.yview
        )
        file_list_scrollbar.grid(row=0, column=1, sticky=tk.NS)
        self._file_listbox.configure(yscrollcommand=file_list_scrollbar.set)

    def _build_warnings_panel(self, parent: ttk.Frame) -> None:
        """Create the read-only panel surfacing ValidationWarning events.

        Args:
            parent: Frame the panel is gridded into.
        """
        warnings_frame = ttk.LabelFrame(parent, text="Warnings", padding=5)
        warnings_frame.grid(row=2, column=0, sticky=tk.NSEW, pady=(10, 0))
        warnings_frame.columnconfigure(0, weight=1)
        warnings_frame.rowconfigure(0, weight=1)

        self._warnings_text = tk.Text(warnings_frame, height=8, state=tk.DISABLED, wrap=tk.WORD)
        self._warnings_text.grid(row=0, column=0, sticky=tk.NSEW)

        warnings_scrollbar = ttk.Scrollbar(
            warnings_frame, orient=tk.VERTICAL, command=self._warnings_text.yview
        )
        warnings_scrollbar.grid(row=0, column=1, sticky=tk.NS)
        self._warnings_text.configure(yscrollcommand=warnings_scrollbar.set)

    def _build_merge_button(self, parent: ttk.Frame) -> None:
        """Create the Merge button and the status label.

        Args:
            parent: Frame the button row is gridded into.
        """
        merge_frame = ttk.Frame(parent)
        merge_frame.grid(row=3, column=0, sticky=tk.EW, pady=(10, 0))

        self._status_label = ttk.Label(merge_frame, text="Select files or a folder to start.")
        self._status_label.pack(side=tk.LEFT)

        self._merge_button = ttk.Button(
            merge_frame, text="Merge", command=self._on_merge_clicked
        )
        self._merge_button.pack(side=tk.RIGHT)

    # --- button handlers -------------------------------------------------

    def _on_add_files_clicked(self) -> None:
        """Open a file dialog and add the chosen .xlsx files to the selection."""
        chosen_file_paths = filedialog.askopenfilenames(
            title="Select excel files",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        self._add_selected_paths([Path(chosen_path) for chosen_path in chosen_file_paths])

    def _on_add_folder_clicked(self) -> None:
        """Open a folder dialog and add every .xlsx file it contains."""
        chosen_folder = filedialog.askdirectory(title="Select a folder with excel files")
        if chosen_folder:
            self._add_selected_paths([Path(chosen_folder)])

    def _add_selected_paths(self, selected_paths: list[Path]) -> None:
        """Delegate path expansion to the ingestion service.

        The service publishes FilesLoaded with the resolved .xlsx paths,
        which comes back through on_event to refresh the file list.

        Args:
            selected_paths: Files and/or folders chosen by the user.
        """
        if not selected_paths:
            return
        self._file_ingestion_service.collect_excel_file_paths(selected_paths)

    def _on_remove_selected_clicked(self) -> None:
        """Remove the highlighted entries from the selected file list."""
        selected_indices = reversed(self._file_listbox.curselection())
        for list_index in selected_indices:
            del self._selected_excel_file_paths[list_index]
            self._file_listbox.delete(list_index)

    def _on_clear_files_clicked(self) -> None:
        """Remove every file from the selection."""
        self._selected_excel_file_paths.clear()
        self._file_listbox.delete(0, tk.END)

    def _on_merge_clicked(self) -> None:
        """Start the merge + schema generation pipeline on a background thread."""
        if not self._selected_excel_file_paths:
            messagebox.showwarning(
                "No files selected", "Please add excel files or a folder before merging."
            )
            return

        self._set_merge_in_progress(True)
        threading.Thread(target=self._run_merge_pipeline, daemon=True).start()

    def _run_merge_pipeline(self) -> None:
        """Run merging and schema generation; failures become MergeFailed events.

        Executed on a worker thread: it must only touch services and the
        event bus, never tkinter widgets directly.
        """
        try:
            merged_workbook = self._file_ingestion_service.merge_worksheet_rows(
                self._selected_excel_file_paths
            )
            self._schema_generator.generate_schemas(merged_workbook)
        except Exception as pipeline_error:
            self._event_bus.publish(MergeFailed(error_message=str(pipeline_error)))

    # --- observer interface ----------------------------------------------

    def on_event(self, event: MergeEvent) -> None:
        """Receive an event and marshal its handling onto the tkinter thread.

        Args:
            event: Event published by any service on the bus.
        """
        self._root.after(0, lambda: self._handle_event_on_ui_thread(event))

    def _handle_event_on_ui_thread(self, event: MergeEvent) -> None:
        """Update widgets in reaction to one event (tkinter thread only).

        Args:
            event: Event to react to.
        """
        if isinstance(event, FilesLoaded):
            self._refresh_file_list(event.file_paths)
        elif isinstance(event, SchemasGenerated):
            self._set_merge_in_progress(False)
            self._last_merged_workbook = event.merged_workbook
            self._status_label.configure(text="Merge finished. Review the schemas below.")
            self._open_schema_editor(event)
        elif isinstance(event, MergeCompleted):
            self._status_label.configure(text=f"Saved to {event.output_path}")
            messagebox.showinfo("Merge completed", f"Merged file saved as:\n{event.output_path}")
        elif isinstance(event, MergeFailed):
            self._set_merge_in_progress(False)
            self._status_label.configure(text="Merge failed.")
            messagebox.showerror("Merge failed", event.error_message)
        elif isinstance(event, ValidationWarning):
            self._append_warning_message(event.message)

    def _refresh_file_list(self, loaded_file_paths: list[Path]) -> None:
        """Add newly loaded paths to the selection and redraw the listbox.

        Args:
            loaded_file_paths: Resolved .xlsx paths from a FilesLoaded event.
        """
        for loaded_file_path in loaded_file_paths:
            if loaded_file_path not in self._selected_excel_file_paths:
                self._selected_excel_file_paths.append(loaded_file_path)
        self._file_listbox.delete(0, tk.END)
        for excel_file_path in self._selected_excel_file_paths:
            self._file_listbox.insert(tk.END, str(excel_file_path))

    def _append_warning_message(self, warning_message: str) -> None:
        """Append one line to the warnings panel.

        Args:
            warning_message: Human-readable warning to display.
        """
        self._warnings_text.configure(state=tk.NORMAL)
        self._warnings_text.insert(tk.END, f"{warning_message}\n")
        self._warnings_text.see(tk.END)
        self._warnings_text.configure(state=tk.DISABLED)

    def _set_merge_in_progress(self, merge_in_progress: bool) -> None:
        """Enable/disable controls while the background merge is running.

        Args:
            merge_in_progress: True while merging, False when finished.
        """
        desired_button_state = tk.DISABLED if merge_in_progress else tk.NORMAL
        for controlled_button in (
            self._merge_button,
            self._add_files_button,
            self._add_folder_button,
            self._remove_selected_button,
            self._clear_files_button,
        ):
            controlled_button.configure(state=desired_button_state)
        if merge_in_progress:
            self._status_label.configure(text="Merging worksheets...")

    def _open_schema_editor(self, schemas_generated_event: SchemasGenerated) -> None:
        """Open (or raise) the schema editor for the merged workbook.

        Args:
            schemas_generated_event: Event carrying the merged workbook with
                pre-generated schemas.
        """
        if self._schema_editor_view is not None:
            self._schema_editor_view.close()
        self._schema_editor_view = SchemaEditorView(
            master=self._root,
            merged_workbook=schemas_generated_event.merged_workbook,
            event_bus=self._event_bus,
            save_requested_handler=self._on_save_output_requested,
        )

    def _on_save_output_requested(self, output_path: Path) -> None:
        """Save the merged workbook (with edited schemas) on a background thread.

        Args:
            output_path: Destination .xlsx path chosen by the user.
        """
        if self._last_merged_workbook is None:
            return
        merged_workbook_to_save = self._last_merged_workbook
        self._status_label.configure(text=f"Saving to {output_path.name}...")
        threading.Thread(
            target=self._run_save_output,
            args=(merged_workbook_to_save, output_path),
            daemon=True,
        ).start()

    def _run_save_output(self, merged_workbook: MergedWorkbook, output_path: Path) -> None:
        """Run the workbook writer; failures become MergeFailed events.

        Executed on a worker thread: it must only touch services and the
        event bus, never tkinter widgets directly.

        Args:
            merged_workbook: Merged workbook whose edited schemas are applied.
            output_path: Destination .xlsx path.
        """
        try:
            self._workbook_writer.write_workbook(merged_workbook, output_path)
        except Exception as save_error:
            self._event_bus.publish(MergeFailed(error_message=str(save_error)))
