import os.path
from typing import Any, List

import tkinter as tk

from tkinter import Tk, Label, Button, Grid, Listbox, Frame, LabelFrame, filedialog
from tkinter.font import Font
from tkinter.ttk import OptionMenu, Combobox

import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet

from diff import TableDiff, TableReference


# TODO: Note: When showing a queue of diffs to process, if there is a first, second, or destination file chosen, present
#             a confirmation dialogue to make sure the user hasn't clicked the "start" button without finishing adding
#             the current diff they're in the process of adding.

# TODO: Add a progress bar when processing the diffs.

class MainWindow:
    _ui_window: Tk

    _ui_first_file_button: Button

    _ui_second_file_button: Button

    _ui_first_file_label: Label

    _ui_second_file_label: Label

    _ui_destination_file_label: Label

    _ui_first_file_table_menu: Combobox

    _ui_second_file_table_menu: Combobox

    _ui_button_row_without_queue: Frame | None = None

    _ui_create_diff_button: Button | None

    _ui_diff_queue: Frame | None = None

    first_file_path: str | None = None

    second_file_path: str | None = None

    destination_file_path: str | None = None

    tables_in_first_file: list[TableReference] | None = None

    tables_in_second_file: list[TableReference] | None = None

    table_selected_from_first_file: TableReference | None = None

    table_selected_from_second_file: TableReference | None = None

    diff_queue: list[TableDiff] = []

    def display(self):
        window: Tk = Tk()
        self._ui_window = window
        window.geometry("600x400")
        window.wm_title("ExcelDiff")

        main_label_row = Frame(window)
        main_label_row.pack_configure(side="top", fill="x", pady=10)
        main_label = Label(main_label_row, text="Choose excel files to diff", font=Font(size=24), justify="center")
        main_label_row.bind("<Configure>", lambda e: main_label.config(wraplength=main_label_row.winfo_width()))
        main_label.pack_configure(fill="x", expand=True)

        file_chooser_button_row_container: Frame = Frame(window)
        file_chooser_button_row_container.pack_configure(side="top", fill="x")
        file_chooser_button_row: Frame = Frame(file_chooser_button_row_container)
        file_chooser_button_row.pack_configure(fill="x", expand=True)
        file_chooser_button_row.grid_columnconfigure(index=0, weight=1)
        file_chooser_button_row.grid_columnconfigure(index=1, weight=1)
        file_chooser_button_row.grid_columnconfigure(index=2, weight=1)
        file_chooser_button_row_spacing = 5

        first_file_button: Button = Button(file_chooser_button_row, text="First", command=self.choose_first_file)
        first_file_button.grid_configure(column=0, row=0, sticky="EW", padx=file_chooser_button_row_spacing)
        self._ui_first_file_button = first_file_button
        first_file_path_label: Label = Label(file_chooser_button_row, text="(No file selected)")
        first_file_path_label.grid_configure(column=0, row=1, padx=file_chooser_button_row_spacing)
        self._ui_first_file_label = first_file_path_label
        first_file_table_menu: Combobox = Combobox(file_chooser_button_row, state="readonly")
        first_file_table_menu.grid_configure(column=0, row=2, sticky="EW", padx=file_chooser_button_row_spacing)
        first_file_table_menu.set("Choose a table...")
        self._ui_first_file_table_menu = first_file_table_menu
        first_file_table_menu.bind("<<ComboboxSelected>>", lambda e: self.update_selected_first_file_table())
        first_file_table_menu.bind("<<ComboboxSelected>>", lambda e: self.update_create_diff_button(), add=True)

        second_file_button: Button = Button(file_chooser_button_row, text="Second", command=self.choose_second_file)
        second_file_button.grid_configure(column=1, row=0, sticky="EW", padx=file_chooser_button_row_spacing)
        self._ui_second_file_button = second_file_button
        second_file_path_label: Label = Label(file_chooser_button_row, text="(No file selected)")
        second_file_path_label.grid_configure(column=1, row=1, padx=file_chooser_button_row_spacing)
        self._ui_second_file_label = second_file_path_label
        second_file_table_menu: Combobox = Combobox(file_chooser_button_row, state="readonly")
        second_file_table_menu.grid_configure(column=1, row=2, sticky="EW", padx=file_chooser_button_row_spacing)
        second_file_table_menu.set("Choose a table...")
        self._ui_second_file_table_menu = second_file_table_menu
        second_file_table_menu.bind("<<ComboboxSelected>>", lambda e: self.update_selected_second_file_table())
        second_file_table_menu.bind("<<ComboboxSelected>>", lambda e: self.update_create_diff_button(), add=True)

        output_file_button: Button = Button(file_chooser_button_row, text="Choose diff location", command=self.choose_destination_file)
        output_file_button.grid_configure(column=2, row=0, sticky="EW", padx=file_chooser_button_row_spacing)
        output_file_path_label: Label = Label(file_chooser_button_row, text="(No destination chosen)")
        output_file_path_label.grid_configure(column=2, row=1, padx=file_chooser_button_row_spacing)
        self._ui_destination_file_label = output_file_path_label

        self._show_button_row()

        self.update_first_file_table_menu()
        self.update_second_file_table_menu()

        window.mainloop()

    def choose_first_file(self):
        path: str = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")], title="First file")

        if(path == ""):
            return

        self.first_file_path = path
        self.tables_in_first_file = self._get_tables_in_file(path)
        self._ui_first_file_label.config(text=os.path.basename(path))
        self.update_first_file_table_menu()
        self.update_create_diff_button()

        # self._ui_first_file_table_menu.config(comm=lambda: self._ui_first_file_table_menu.insert("end", "Doot"))
        # self._ui_first_file_table_menu.bind("<<ComboboxSelected>>", lambda e: self._ui_first_file_table_menu.insert("end", "Doot"))
        # self._ui_first_file_table_menu.bind("<Button-1>", lambda e: print("Doot!"))

    def choose_second_file(self):
        path: str = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")], title="Second file")

        if(path == ""):
            return

        self.second_file_path = path
        self.tables_in_second_file = self._get_tables_in_file(path)
        self._ui_second_file_label.config(text=os.path.basename(path))
        self.update_second_file_table_menu()
        self.update_create_diff_button()

    def choose_destination_file(self):
        path: str = filedialog.asksaveasfilename(confirmoverwrite=True, defaultextension=".xlsx", filetypes=[("Excel file", "*.xlsx")])

        if(path == ""):
            return

        self.destination_file_path = path
        self._ui_destination_file_label.config(text=os.path.basename(path))
        self.update_create_diff_button()

    def create_single_diff(self):
        raise NotImplementedError("Not yet implemented.")

    def update_first_file_table_menu(self):
        if(self.first_file_path is None):
            self.table_selected_from_first_file = None
            self._ui_first_file_table_menu.set("Choose a table...")
            self._ui_first_file_table_menu["state"] = "disabled"
            return

        self._ui_first_file_table_menu.config(values=[x.table_name for x in self.tables_in_first_file])
        self._ui_first_file_table_menu["state"] = "readonly"

    def update_second_file_table_menu(self):
        if(self.second_file_path is None):
            self.table_selected_from_second_file = None
            self._ui_second_file_table_menu.set("Choose a table...")
            self._ui_second_file_table_menu["state"] = "disabled"
            return

        self._ui_second_file_table_menu.config(values=[x.table_name for x in self.tables_in_second_file])
        self._ui_second_file_table_menu["state"] = "readonly"

    def update_selected_first_file_table(self):
        table_name: str = self._ui_first_file_table_menu.get()

        matching_tables: list[TableReference] = [x for x in self.tables_in_first_file if x.table_name == table_name]
        self.table_selected_from_first_file = matching_tables[0] if len(matching_tables) > 0 else None

    def update_selected_second_file_table(self):
        table_name: str = self._ui_second_file_table_menu.get()

        matching_tables: list[TableReference] = [x for x in self.tables_in_second_file if x.table_name == table_name]
        self.table_selected_from_second_file = matching_tables[0] if len(matching_tables) > 0 else None

    def update_create_diff_button(self):
        if(self._ui_create_diff_button is None):
            return

        if(None in [self.first_file_path, self.second_file_path, self.destination_file_path,
                    self.table_selected_from_first_file, self.table_selected_from_second_file]):

            self._ui_create_diff_button["state"] = "disabled"
            return

        self._ui_create_diff_button["state"] = "normal"

    def add_diff_to_queue(self, first_filepath: str, second_filepath: str):
        # self.diff_queue.append(TableDiff())
        pass

    def _hide_button_row(self):
        if(self._ui_button_row_without_queue is None):
            return

        self._ui_button_row_without_queue.pack_forget()
        self._ui_button_row_without_queue = None
        self._ui_create_diff_button = None

    def _hide_diff_queue(self):
        if(self._ui_diff_queue is None):
            return

        self._ui_diff_queue.pack_forget()
        self._ui_diff_queue = None

    def _show_button_row(self):
        action_button_row: Frame = Frame(self._ui_window)
        action_button_row.pack_configure(side="bottom", fill="x")
        enqueue_button: Button = Button(action_button_row, text="Enqueue")
        enqueue_button.pack_configure(side="left", padx=5, pady=5)
        enqueue_button["state"] = "disabled"
        create_diff_button: Button = Button(action_button_row, text="Create diff", command=self.create_single_diff)
        create_diff_button.pack_configure(side="right", fill="x", expand=True, padx=5, pady=5)
        self._ui_create_diff_button       = create_diff_button
        self._ui_button_row_without_queue = action_button_row

        self.update_create_diff_button()

    def _show_diff_queue(self):
        diff_queue: Frame = Frame(self._ui_window)
        diff_queue.pack_configure(side="bottom", fill="both", expand=True)

        add_to_queue_button_container: Frame = Frame(diff_queue)
        add_to_queue_button_container.pack_configure(side="top", fill="x")
        add_to_queue_button: Button = Button(add_to_queue_button_container, text="+ Add to queue")
        add_to_queue_button.pack_configure(fill="both", padx=5, pady=5)

        queue_list: Listbox = Listbox(diff_queue, height=2)
        queue_list.pack_configure(fill="both", expand=True, padx=5)

        diff_queue_bottom_button_row: Frame = Frame(diff_queue)
        diff_queue_bottom_button_row.pack_configure(side="bottom", fill="x")
        remove_diff_from_queue_button: Button = Button(diff_queue_bottom_button_row, text="Remove",
                                                       background="#ffafaf", activebackground="#bc6262")
        remove_diff_from_queue_button.pack_configure(fill="none", padx=5, pady=5, side="left")
        remove_diff_from_queue_button["state"] = "disabled"
        create_diffs_from_queue_button: Button = Button(diff_queue_bottom_button_row, text="Create diffs")
        create_diffs_from_queue_button.pack_configure(fill="x", expand=True, padx=5, pady=5, side="right")

        queue_list.insert("end", "One")
        queue_list.insert("end", "Two")
        queue_list.insert("end", "Three")
        queue_list.insert("end", "Four")
        queue_list.insert("end", "Five")

        self._ui_diff_queue = diff_queue

    def _switch_to_button_row(self):
        self._hide_button_row()
        self._hide_diff_queue()
        self._show_button_row()

    def _switch_to_diff_queue(self):
        self._hide_button_row()
        self._hide_diff_queue()
        self._show_diff_queue()

    def _get_tables_in_file(self, filepath: str) -> list[TableReference]:
        result: list[TableReference] = []

        wb: Workbook = openpyxl.load_workbook(filepath, data_only=True)
        sheet_names = wb.sheetnames

        for sheet_name in sheet_names:
            sheet: Worksheet = wb.get_sheet_by_name(sheet_name)

            table_name: str

            for table_name in sheet.tables.keys():
                result.append(TableReference(filepath, sheet_name, table_name))

        return result
