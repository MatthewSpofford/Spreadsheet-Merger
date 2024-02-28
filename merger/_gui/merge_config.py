import os
import traceback
from pathlib import Path
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *

from merger import app
from merger.exceptions import MergeException
from merger._gui import loading_screen, error_message
from merger._config import Config, ConfigProperty
from merger.nonblocking_merger import NonblockingMerger

_COLUMN_WIDTH = 3
_ROW_HEIGHT = 6


ORIGINAL_SPREADSHEET_LABEL = "Original Spreadsheet"
APPENDING_SPREADSHEET_LABEL = "Appending Spreadsheet"
COLUMN_KEY_LABEL = "ID Column"
REPLACE_ORIG_LABEL = "Replace original spreadsheet after merge"
NEW_SPREADSHEET_LABEL = "Merged Spreadsheet Name"
MERGE_BTN_LABEL = "Merge"


class MergeConfig(Frame):
    """Starting page of the GUI for entering spreadsheet merging configuration info"""

    initial_dir = Path(Config.get(ConfigProperty.INITIAL_DIR))
    new_file_path = Path()

    def __init__(self, root):
        super().__init__(root, padding=("20", "20", "20", "20"))

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.root = root
        self.root.bind("<Return>", lambda key: self.merge_spreadsheets())

        self._main_select = SpreadsheetSelect(self, 0, ORIGINAL_SPREADSHEET_LABEL)
        self._main_select.file_path = Config.get(ConfigProperty.ORIGINAL_PATH)
        if __debug__ and self._main_select.file_path == "":
            self._main_select.file_path = "C:\\workspace\\Spreadsheet-Merger\\sample\\main.xlsx"

        self._new_select = SpreadsheetSelect(self, 1, APPENDING_SPREADSHEET_LABEL)
        self._new_select.file_path = self.new_file_path
        if __debug__ and self._new_select.file_path == Path():
            self._new_select.file_path = "C:\\workspace\\Spreadsheet-Merger\\samples\\additions.xlsx"

        self._col_key = EntryWithLabel(self, 2, COLUMN_KEY_LABEL)
        self._col_key.entry_text = Config.get(ConfigProperty.COLUMN_KEY)

        self._replace_orig_check = CheckToHideEntry(self, REPLACE_ORIG_LABEL, NEW_SPREADSHEET_LABEL)
        self._replace_orig_check.grid(column=0, row=4, columnspan=_COLUMN_WIDTH)
        self._replace_orig_check.checked = Config.get(ConfigProperty.REPLACE_ORIGINAL)
        self._replace_orig_check.entry_text = Config.get(ConfigProperty.MERGED_FILENAME)

        self._merge_btn = Button(self, text=MERGE_BTN_LABEL, command=self.merge_spreadsheets)
        self._merge_btn.grid(column=0, row=5, columnspan=_COLUMN_WIDTH, padx=120, pady=10, sticky="nsew")

    def merge_spreadsheets(self):
        try:
            # Validate file selector input for all selections
            selectors = [self._main_select, self._new_select]
            for select in selectors:
                # Output error if file path is invalid and cancel merge
                if not select.file_path.is_file():
                    raise MergeException("The file path given for the " + select.label_text.lower() + " is not a valid "
                                         "path.")

            # Validate that the merge file name is properly filled
            if not self._replace_orig_check.checked and \
               (self._replace_orig_check.entry_text is None or len(self._replace_orig_check.entry_text) == 0):
                raise MergeException("A new name for the file needs to be provided. Or, check the box to have the "
                                     "original spreadsheet be replaced")

            # Save new file path for when the configuration window is viewed again
            self.new_file_path = self._new_select.file_path

            replacement_file = self._replace_orig_check.entry_text if not self._replace_orig_check.checked else None
            merger = NonblockingMerger(Path(self._main_select.file_path),
                                       self.new_file_path,
                                       self._col_key.entry_text,
                                       replacement_file)

            # Save to the configuration file
            Config.set({
                ConfigProperty.ORIGINAL_PATH: self._main_select.file_path,
                ConfigProperty.COLUMN_KEY: self._col_key.entry_text,
                ConfigProperty.REPLACE_ORIGINAL: self._replace_orig_check.checked,
                ConfigProperty.MERGED_FILENAME: self._replace_orig_check.entry_text,
                ConfigProperty.INITIAL_DIR: self.initial_dir,
            })

            app.switch_frames(loading_screen.LoadingScreen, merger)
        except MergeException as e:
            error_message.display_error("Merge Configuration Error", e)
            app.switch_frames(MergeConfig)
            return
        except BaseException:
            error_message.display_fatal_error("Merge Configuration Failure", traceback.format_exc())
            app.switch_frames(MergeConfig)
            return


class SpreadsheetSelect:
    SELECT_SPREADSHEET_TITLE = "Select Spreadsheet"

    def __init__(self, root, row, label_text: str, overwrite_init_dir=True):
        self.label_text = label_text

        self._label = Label(root, text=self.label_text + ":")
        self._label.grid(column=0, row=row, padx=20, pady=15, sticky="nse")

        self._file_path = StringVar()
        self._file_path_entry = Entry(root, textvariable=self._file_path)
        self._file_path_entry.grid(column=1, columnspan=_COLUMN_WIDTH - 2,
                                   row=row, padx=0, pady=15, sticky="nsew")

        self._select_btn = Button(root, text="Select...", command=self._select_file)
        self._select_btn.grid(column=_COLUMN_WIDTH - 1, row=row, padx=20, pady=15, sticky="nsew")

        self._overwrite_init_dir = overwrite_init_dir

    def _select_file(self):
        new_file_path = filedialog.askopenfilename(
            filetypes=[("Excel Spreadsheet", ".xlsx")],
            title=self.SELECT_SPREADSHEET_TITLE,
            initialdir=MergeConfig.initial_dir)
        if not new_file_path:
            return
        self.file_path = new_file_path

    @property
    def file_path(self):
        return Path(self._file_path.get().strip())

    @file_path.setter
    def file_path(self, file_path):
        self._file_path.set(file_path)

        self._file_path.set(file_path)

        # Save initial dir in the parent directory of this newly selected file
        if self._overwrite_init_dir:
            MergeConfig.initial_dir = os.path.dirname(file_path)


class EntryWithLabel:
    def __init__(self, root, row: int, label_text: str):
        self._col_key_label = Label(root, text=f"{label_text}:")
        self._col_key_label.grid(column=0, row=row, padx=20, pady=15, sticky="nse")

        self._entry_text = StringVar()
        self._entry = Entry(root, textvariable=self._entry_text)
        self._entry.grid(column=1, columnspan=_COLUMN_WIDTH - 1,
                         row=row, padx=20, pady=15, sticky="nsew")

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
        Arguments:
            additional_check_handle: is a function that is called after handling the initial check of the checkbox.
            This function takes a bool as a parameter.
        """
        super().__init__(root, padding=("20", "20", "20", "20"), relief="solid")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._check_state = IntVar(value=int(check_state))
        self._check = Checkbutton(self, variable=self._check_state, text=check_label_text, command=self._check_handler)
        self._check.grid(column=0, row=0, columnspan=_COLUMN_WIDTH, sticky="ns")

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
