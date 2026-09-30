"""Schema editor window: one tab per merged worksheet with editable columns."""

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Callable

from domain.merged_workbook import MergedWorkbook
from domain.worksheet_schema import WorksheetSchema
from events.event_bus import EventBus, SchemaEdited


class SchemaEditorView:
    """Toplevel window where the user reviews and edits pre-generated schemas.

    Shows one tab per merged worksheet. Each tab is a ColumnEditorTab that
    operates on the worksheet's WorksheetSchema (rename, reorder, add and
    remove columns). Every edit publishes a SchemaEdited event; saving is
    delegated to an injected handler so the view owns no writing logic
    (composition over inheritance).
    """

    def __init__(
        self,
        master: tk.Misc,
        merged_workbook: MergedWorkbook,
        event_bus: EventBus,
        save_requested_handler: Callable[[Path], None] | None = None,
    ) -> None:
        """Build the editor window with one tab per merged worksheet.

        Args:
            master: Parent widget (the main window's root).
            merged_workbook: Merged workbook whose schemas will be edited.
            event_bus: Bus used to publish SchemaEdited events.
            save_requested_handler: Callback invoked with the chosen output
                path when the user requests saving; None hides the save button.
        """
        self._merged_workbook = merged_workbook
        self._event_bus = event_bus
        self._save_requested_handler = save_requested_handler

        self._window = tk.Toplevel(master)
        self._window.title("Review merged schemas")
        self._window.geometry("560x480")
        self._window.transient(master)

        self._notebook = ttk.Notebook(self._window)
        self._notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))

        self._column_editor_tabs: list[ColumnEditorTab] = []
        for merged_worksheet in self._merged_workbook.worksheets:
            if merged_worksheet.schema is None:
                continue
            column_editor_tab = ColumnEditorTab(
                master=self._notebook,
                worksheet_schema=merged_worksheet.schema,
                schema_edited_handler=self._publish_schema_edited,
            )
            self._notebook.add(column_editor_tab.frame, text=merged_worksheet.sheet_title)
            self._column_editor_tabs.append(column_editor_tab)

        self._build_bottom_buttons()

    def _build_bottom_buttons(self) -> None:
        """Create the Save Output and Close buttons at the window bottom."""
        bottom_buttons_frame = ttk.Frame(self._window)
        bottom_buttons_frame.pack(fill=tk.X, padx=10, pady=10)

        if self._save_requested_handler is not None:
            self._save_button = ttk.Button(
                bottom_buttons_frame,
                text="Save Output...",
                command=self._on_save_output_clicked,
            )
            self._save_button.pack(side=tk.RIGHT)

        self._close_button = ttk.Button(
            bottom_buttons_frame, text="Close", command=self.close
        )
        self._close_button.pack(side=tk.RIGHT, padx=(0, 8))

    def _on_save_output_clicked(self) -> None:
        """Ask the user for an output path and delegate the save request."""
        chosen_output_path = filedialog.asksaveasfilename(
            title="Save merged workbook",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            parent=self._window,
        )
        if not chosen_output_path:
            return
        assert self._save_requested_handler is not None
        self._save_requested_handler(Path(chosen_output_path))

    def _publish_schema_edited(self, edited_schema: WorksheetSchema) -> None:
        """Notify observers that one schema was edited by the user.

        Args:
            edited_schema: Schema that changed in the editor.
        """
        self._event_bus.publish(
            SchemaEdited(sheet_title=edited_schema.sheet_title, schema=edited_schema)
        )

    def close(self) -> None:
        """Destroy the editor window."""
        self._window.destroy()


class ColumnEditorTab:
    """Editor for one worksheet schema: list of columns plus edit buttons.

    Composed into a SchemaEditorView notebook (one instance per merged
    worksheet). Mutates the given WorksheetSchema in place and reports each
    change through the injected schema_edited_handler callback.
    """

    def __init__(
        self,
        master: tk.Misc,
        worksheet_schema: WorksheetSchema,
        schema_edited_handler: Callable[[WorksheetSchema], None],
    ) -> None:
        """Build the tab widgets for one schema.

        Args:
            master: Notebook widget this tab lives in.
            worksheet_schema: Schema edited by this tab (mutated in place).
            schema_edited_handler: Callback invoked with the schema after
                every successful edit.
        """
        self._worksheet_schema = worksheet_schema
        self._schema_edited_handler = schema_edited_handler

        self._frame = ttk.Frame(master, padding=10)
        self._frame.columnconfigure(0, weight=1)
        self._frame.rowconfigure(0, weight=1)

        self._column_listbox = tk.Listbox(self._frame, selectmode=tk.SINGLE)
        self._column_listbox.grid(row=0, column=0, sticky=tk.NSEW)

        listbox_scrollbar = ttk.Scrollbar(
            self._frame, orient=tk.VERTICAL, command=self._column_listbox.yview
        )
        listbox_scrollbar.grid(row=0, column=1, sticky=tk.NS)
        self._column_listbox.configure(yscrollcommand=listbox_scrollbar.set)

        self._build_edit_buttons()
        self._refresh_column_listbox()

    @property
    def frame(self) -> ttk.Frame:
        """Tkinter frame containing this tab's widgets."""
        return self._frame

    def _build_edit_buttons(self) -> None:
        """Create the column edit buttons on the right side of the tab."""
        buttons_frame = ttk.Frame(self._frame)
        buttons_frame.grid(row=0, column=2, sticky=tk.N, padx=(10, 0))

        button_specifications = [
            ("Rename...", self._on_rename_column_clicked),
            ("Move Up", self._on_move_column_up_clicked),
            ("Move Down", self._on_move_column_down_clicked),
            ("Add...", self._on_add_column_clicked),
            ("Remove", self._on_remove_column_clicked),
        ]
        for button_text, button_command in button_specifications:
            ttk.Button(buttons_frame, text=button_text, command=button_command).pack(
                fill=tk.X, pady=(0, 5)
            )

    def _refresh_column_listbox(self) -> None:
        """Redraw the listbox to mirror the current schema column order."""
        self._column_listbox.delete(0, tk.END)
        for column_name in self._worksheet_schema.column_names:
            self._column_listbox.insert(tk.END, column_name)

    def _get_selected_column_index(self) -> int | None:
        """Return the index of the selected column, or None when nothing is selected."""
        selected_indices = self._column_listbox.curselection()
        if not selected_indices:
            return None
        return int(selected_indices[0])

    def _notify_schema_edited(self) -> None:
        """Refresh the list and report the edit through the handler callback."""
        self._refresh_column_listbox()
        self._schema_edited_handler(self._worksheet_schema)

    def _on_rename_column_clicked(self) -> None:
        """Ask for a new name and rename the selected column."""
        selected_column_index = self._get_selected_column_index()
        if selected_column_index is None:
            messagebox.showinfo("Rename column", "Select a column to rename.", parent=self._frame)
            return
        current_column_name = self._worksheet_schema.column_names[selected_column_index]
        new_column_name = simpledialog.askstring(
            "Rename column",
            "New column name:",
            initialvalue=current_column_name,
            parent=self._frame,
        )
        if new_column_name is None or new_column_name.strip() == "":
            return
        self._worksheet_schema.rename_column(selected_column_index, new_column_name.strip())
        self._notify_schema_edited()
        self._column_listbox.selection_set(selected_column_index)

    def _on_move_column_up_clicked(self) -> None:
        """Move the selected column one position up (reorder)."""
        self._move_selected_column_by(-1)

    def _on_move_column_down_clicked(self) -> None:
        """Move the selected column one position down (reorder)."""
        self._move_selected_column_by(1)

    def _move_selected_column_by(self, index_offset: int) -> None:
        """Move the selected column by the given offset, keeping it selected.

        Args:
            index_offset: -1 to move up, +1 to move down.
        """
        selected_column_index = self._get_selected_column_index()
        if selected_column_index is None:
            messagebox.showinfo("Move column", "Select a column to move.", parent=self._frame)
            return
        target_column_index = selected_column_index + index_offset
        if not 0 <= target_column_index < len(self._worksheet_schema.column_names):
            return
        self._worksheet_schema.move_column(selected_column_index, target_column_index)
        self._notify_schema_edited()
        self._column_listbox.selection_set(target_column_index)

    def _on_add_column_clicked(self) -> None:
        """Ask for a name and add a new column after the selected position."""
        new_column_name = simpledialog.askstring(
            "Add column", "New column name:", parent=self._frame
        )
        if new_column_name is None or new_column_name.strip() == "":
            return
        selected_column_index = self._get_selected_column_index()
        insert_column_index = (
            None if selected_column_index is None else selected_column_index + 1
        )
        self._worksheet_schema.add_column(new_column_name.strip(), insert_column_index)
        self._notify_schema_edited()
        if insert_column_index is not None:
            self._column_listbox.selection_set(insert_column_index)

    def _on_remove_column_clicked(self) -> None:
        """Remove the selected column from the schema after confirmation."""
        selected_column_index = self._get_selected_column_index()
        if selected_column_index is None:
            messagebox.showinfo("Remove column", "Select a column to remove.", parent=self._frame)
            return
        removed_column_name = self._worksheet_schema.column_names[selected_column_index]
        if not messagebox.askyesno(
            "Remove column",
            f"Remove column '{removed_column_name}' from the output?",
            parent=self._frame,
        ):
            return
        self._worksheet_schema.remove_column(selected_column_index)
        self._notify_schema_edited()
