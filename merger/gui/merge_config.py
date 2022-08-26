import logging
import os
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import *

from merger.merger import merge_spreadsheets
from merger._config import Config, ConfigProperty

_initial_dir = Config.get(ConfigProperty.INITIAL_DIR)
_column_width = 3


class MergerGUI(Frame):
    def __init__(self, root):
        super().__init__(root, padding=("20", "20", "20", "20"))
        self._root = root
        self._root.bind("<Return>", self.merge_spreadsheets)

        self._main_select = SpreadsheetSelect(root, 0, "Original Spreadsheet")
        self._main_select.file_path = Config.get(ConfigProperty.ORIGINAL_PATH)
        if __debug__ and self._main_select.file_path == "":
            self._main_select.file_path = "/samples/main.xlsx"

        self._new_select = SpreadsheetSelect(root, 1, "Appending Spreadsheet")
        # if __debug__:
        #     self._new_select.file_path = "/samples/additions.xlsx"

        self._col_key = EntryWithLabel(root, 2, "Column Key")
        self._col_key.entry_text = Config.get(ConfigProperty.COLUMN_KEY)

        # self._sheet_name = EntryWithLabel(root, 3, "Worksheet Name to Merge")
        # if __debug__:
        #     self._sheet_name.entry_text = "Manager_Opportunity_Dashboard"
        # self._sheet_name.entry_text = Config.get(Property.sheet_name)

        self._replace_orig_check = CheckToHideEntry(root,
                                                    "Replace original spreadsheet after merge",
                                                    "Merged Spreadsheet Name")
        self._replace_orig_check.grid(column=0, row=4, columnspan=_column_width)
        self._replace_orig_check.checked = Config.get(ConfigProperty.REPLACE_ORIGINAL)
        self._replace_orig_check.entry_text = Config.get(ConfigProperty.MERGED_FILENAME)

        self._merge_btn = Button(root, text="Merge", command=self.merge_spreadsheets)
        self._merge_btn.grid(column=0, row=5, columnspan=_column_width, padx=20, pady=10)

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
            if not self._replace_orig_check.checked and \
               (self._replace_orig_check.entry_text is None or len(self._replace_orig_check.entry_text) == 0):
                raise Exception("A new name for the file needs to be provided. Or, check the box to have the original "
                                "spreadsheet be replaced")

            # Save to the configuration file
            Config.set({
                ConfigProperty.ORIGINAL_PATH: self._main_select.file_path,
                ConfigProperty.COLUMN_KEY: self._col_key.entry_text,
                ConfigProperty.REPLACE_ORIGINAL: self._replace_orig_check.checked,
                ConfigProperty.MERGED_FILENAME: self._replace_orig_check.entry_text,
                ConfigProperty.INITIAL_DIR: _initial_dir,
            })
            merge_spreadsheets(self._main_select.file_path,
                               self._new_select.file_path,
                               self._col_key.entry_text,
                               # self._sheet_name.entry_text,
                               self._replace_orig_check.entry_text)


            messagebox.showinfo("Merge Complete", "Spreadsheets merged successfully!")
        except BaseException as e:
            logging.exception("Merge exception!")
            messagebox.showerror("Merge Exception", str(e))
            return


class SpreadsheetSelect:
    def __init__(self, root, row, label_text: str, overwrite_init_dir=True):
        # super().__init__(root)

        # self.columnconfigure(2, weight=1)
        # self.rowconfigure(0, weight=1)

        self.label_text = label_text

        self._label = Label(root, text=self.label_text + ":")
        self._label.grid(column=0, row=row, padx=20, pady=15)

        self._file_path = StringVar()
        self._file_path_entry = Entry(root, textvariable=self._file_path)
        self._file_path_entry.grid(column=1, columnspan=_column_width - 2,
                                   row=row, padx=0, pady=15)

        self._select_btn = Button(root, text="Select...", command=self._select_file)
        self._select_btn.grid(column=_column_width - 1, row=row, padx=20, pady=15)

        self._overwrite_init_dir = overwrite_init_dir

    def _select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Spreadsheet", ".xlsx")],
                                                    title="Select Spreadsheet",
                                                    initialdir=_initial_dir)

    @property
    def file_path(self):
        return self._file_path.get()

    @file_path.setter
    def file_path(self, file_path):
        self._file_path.set(file_path)

        self._file_path.set(file_path)

        # Save initial dir in the parent directory of this newly selected file
        if self._overwrite_init_dir:
            global _initial_dir
            _initial_dir = os.path.dirname(file_path)


class EntryWithLabel:
    def __init__(self, root, row: int, label_text: str):
        self._col_key_label = Label(root, text=f"{label_text}:")
        self._col_key_label.grid(column=0, row=row, padx=20, pady=15)

        self._entry_text = StringVar()
        self._entry = Entry(root, textvariable=self._entry_text)
        self._entry.grid(column=1, columnspan=_column_width - 1,
                         row=row, padx=0, pady=15)

    @property
    def entry_text(self) -> str:
        return self._entry_text.get()

    @entry_text.setter
    def entry_text(self, entry_text: str):
        self._entry_text.set(entry_text)

    def disable_entry(self, disable: bool):
        new_state = DISABLED if disable else NORMAL
        self._entry.config(state=new_state)


class CheckToHideEntry(Frame):
    def __init__(self, root, check_label_text: str, entry_label: str, additional_check_handle=None,
                 check_state=True):
        """
        :param additional_check_handle: is a function that is called after handling the initial check of the checkbox.
        This function takes a bool as a parameter.
        """
        super().__init__(root, padding=("20", "20", "20", "20"), relief="solid")

        self._check_state = IntVar(value=int(check_state))
        self._check = Checkbutton(self, variable=self._check_state, text=check_label_text, command=self._check_handler)
        self._check.grid(column=0, row=0, columnspan=_column_width)

        self._entry = EntryWithLabel(self, 1, entry_label)

        # Invoke check command on initialization
        self._additional_check_handle = additional_check_handle
        self._check_handler()

    @property
    def checked(self) -> bool:
        return bool(self._check_state.get())

    @checked.setter
    def checked(self, value: bool):
        self._check_state.set(int(value))
        self._check_handler()

    def _check_handler(self):
        checked = self.checked
        self._entry.disable_entry(disable=checked)

        if checked:
            self._entry.entry_text = ""

        if self._additional_check_handle is not None:
            self._additional_check_handle(checked)

    @property
    def entry_text(self) -> str:
        return self._entry.entry_text

    @entry_text.setter
    def entry_text(self, text: str):
        self._entry.entry_text = text

    @property
    def additional_check_handle(self):
        return self._additional_check_handle

    @additional_check_handle.setter
    def additional_check_handle(self, handle):
        self._additional_check_handle = handle
        self._check_handler()
