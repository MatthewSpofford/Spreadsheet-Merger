import os
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import *

import merger.merger as merger


_initial_dir = "./"


class MergerGUI(Frame):

    def __init__(self, root):
        super().__init__(root, padding=("20", "20", "20", "20"))
        self._root = root

        self.columnconfigure(2, weight=1)
        self.rowconfigure(2, weight=1)

        self._main_select = SpreadsheetSelect(root, 0, "Original Spreadsheet")
        if __debug__:
            self._main_select.file_path = "C:\\workspace\\Spreadsheet-Merger\\samples\\" \
                                          "NortheastOpportunities_081122_ColumnHeaders.xlsx"

        self._new_select = SpreadsheetSelect(root, 1, "New Spreadsheet")
        if __debug__:
            self._new_select.file_path = "C:\\workspace\\Spreadsheet-Merger\\samples\\" \
                                         "NortheastOpportunities_081122 - Copy.xlsx"

        self._merge_btn = Button(root, text="Merge", command=self.merge_spreadsheets)
        self._merge_btn.grid(column=0, row=2, columnspan=3, padx=20, pady=10)

        for child in self.winfo_children():
            child.grid_configure(padx=5, pady=5)

        self._main_select._file_path_entry.focus()

    def merge_spreadsheets(self):
        # Validate file selector input for all selections
        selectors = [self._main_select, self._new_select]
        for select in selectors:
            select.file_path = select.file_path.strip()

            # Output error if file path is invalid and cancel merge
            if not os.path.isfile(select.file_path):
                messagebox.showerror("Merge Error",
                                     "The file path given for the " + select.label_text.lower()
                                     + " is not a valid path.")
                return

        # Begin spreadsheet merging process
        try:
            merger.merge_spreadsheets(self._main_select.file_path,
                                      self._new_select.file_path)
        except Exception as e:
            messagebox.showerror("Merge Error", str(e))

        # Close app
        self._root.quit()


class SpreadsheetSelect:

    def __init__(self, root, row, label_text: str):
        # super().__init__(root)

        # self.columnconfigure(2, weight=1)
        # self.rowconfigure(0, weight=1)

        self.label_text = label_text

        self._label = Label(root, text=self.label_text + ":")
        self._label.grid(column=0, row=row, padx=20, pady=15)

        self._file_path = StringVar()
        self._file_path_entry = Entry(root, textvariable=self._file_path)
        self._file_path_entry.grid(column=1, row=row, padx=0, pady=15)

        self._select_btn = Button(root, text="Select...", command=self._select_file)
        self._select_btn.grid(column=2, row=row, padx=20, pady=15)

    def _select_file(self):
        global _initial_dir
        file_path = filedialog.askopenfilename(filetypes=[("Excel Spreadsheet", ".xlsx")],
                                               title="Select Spreadsheet",
                                               initialdir=_initial_dir)

        # Save initial dir in the parent directory of this newly selected file
        _initial_dir = os.path.dirname(file_path)

        self._file_path.set(file_path)

    @property
    def file_path(self):
        return self._file_path.get()

    @file_path.setter
    def file_path(self, file_path):
        self._file_path.set(file_path)
