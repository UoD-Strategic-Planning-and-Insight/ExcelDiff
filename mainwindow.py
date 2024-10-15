import tkinter as tk

from tkinter import Tk, Label, Button, Grid, Listbox, Frame, LabelFrame
from tkinter.font import Font

class MainWindow:
    _ui_window: Tk

    _ui_button_row_without_queue: Frame | None = None

    _ui_diff_queue: Frame | None = None

    def display(self):
        window: Tk = Tk()
        self._ui_window = window
        window.geometry("600x400")

        main_label_row = Frame(window)
        main_label_row.pack_configure(side="top", fill="x", pady=10)
        main_label = Label(main_label_row, text="Choose excel files to diff", font=Font(size=24), justify="center")
        main_label_row.bind("<Configure>", lambda e: main_label.config(wraplength=main_label_row.winfo_width()))
        main_label.pack_configure(fill="x", expand=True)

        file_chooser_button_row_container: Frame = Frame(window)
        # file_chooser_button_row_container.pack_configure(side="top", fill="both", expand=True)
        file_chooser_button_row_container.pack_configure(side="top", fill="x")
        file_chooser_button_row: Frame = Frame(file_chooser_button_row_container)
        file_chooser_button_row.pack_configure(fill="x", expand=True)

        first_file_button_container: Frame = Frame(file_chooser_button_row)
        first_file_button_container.pack_configure(side="left", fill="both", expand=True)
        first_file_button: Button = Button(first_file_button_container, text="First")
        first_file_button.pack_configure(side="top", fill="x", expand=True, padx=5, pady=5)
        first_file_path_label: Label = Label(first_file_button_container, text="first.xlsx")
        first_file_path_label.pack_configure(side="bottom", fill="x", expand=True)

        second_file_button_container: Frame = Frame(file_chooser_button_row)
        second_file_button_container.pack_configure(side="right", fill="both", expand=True)
        second_file_button: Button = Button(second_file_button_container, text="Second")
        second_file_button.pack_configure(side="top", fill="x", expand=True, padx=5, pady=5)
        second_file_path_label: Label = Label(second_file_button_container, text="second.xlsx")
        second_file_path_label.pack_configure(side="bottom", fill="x", expand=True)

        self._show_button_row()

        window.mainloop()

    def _hide_button_row(self):
        if(self._ui_button_row_without_queue is None):
            return

        self._ui_button_row_without_queue.pack_forget()
        self._ui_button_row_without_queue = None

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
        create_diff_button: Button = Button(action_button_row, text="Create diff")
        create_diff_button.pack_configure(side="right", fill="x", expand=True, padx=5, pady=5)

        self._ui_button_row_without_queue = action_button_row

    def _show_diff_queue(self):
        diff_queue: Frame = Frame(self._ui_window)
        diff_queue.pack_configure(side="bottom", fill="both", expand=True)

        add_to_queue_button_container: Frame = Frame(diff_queue)
        add_to_queue_button_container.pack_configure(side="top", fill="x")
        add_to_queue_button: Button = Button(add_to_queue_button_container, text="+ Add to queue")
        add_to_queue_button.pack_configure(fill="both", padx=5, pady=5)

        queue_list: Listbox = Listbox(diff_queue, height=2)
        queue_list.pack_configure(fill="both", expand=True, padx=5)

        create_diffs_from_queue_button_container: Frame = Frame(diff_queue)
        create_diffs_from_queue_button_container.pack_configure(side="bottom", fill="x")
        create_diffs_from_queue_button: Button = Button(create_diffs_from_queue_button_container, text="Create diffs")
        create_diffs_from_queue_button.pack_configure(fill="both", padx=5, pady=5)

        self._ui_diff_queue = diff_queue

    def _switch_to_button_row(self):
        self._hide_button_row()
        self._hide_diff_queue()
        self._show_button_row()

    def _switch_to_diff_queue(self):
        self._hide_button_row()
        self._hide_diff_queue()
        self._show_diff_queue()




