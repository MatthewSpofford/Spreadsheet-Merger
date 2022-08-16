from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *

from merger.merger import merge_spreadsheets


class MergerGUI(Frame):

  def __init__(self, root):
    super().__init__(root, padding=("20", "20", "20", "20"))

    self.columnconfigure(2, weight=1)
    self.rowconfigure(2, weight=1)

    self._main_select = SpreadsheetSelect(root, 0, "Original Spreadsheet:")

    self._new_select = SpreadsheetSelect(root, 1, "New Spreadsheet:")

    self._merge_btn = Button(root, text="Merge", command=self.merge_spreadsheets)
    self._merge_btn.grid(column=0, row=2, columnspan=3, padx=20, pady=10)

    for child in self.winfo_children():
      child.grid_configure(padx=5, pady=5)

    self._main_select._file_path_entry.focus()

  def merge_spreadsheets(self):
    try:
      merge_spreadsheets(self._main_select.file_path,
                         self._main_select.file_path)
    except Exception as e:
      print(e)

    from merger.__main__ import root
    root.quit()


class SpreadsheetSelect:

  def __init__(self, root, row, label_text):
    # super().__init__(root)

    # self.columnconfigure(2, weight=1)
    # self.rowconfigure(0, weight=1)

    self._label = Label(root, text=label_text)
    self._label.grid(column=0, row=row, padx=20, pady=15)

    self._file_path = StringVar()
    self._file_path_entry = Entry(root, textvariable=self._file_path)
    self._file_path_entry.grid(column=1, row=row, padx=0, pady=15)

    self._select_btn = Button(root, text="Select...", command=self._select_file)
    self._select_btn.grid(column=2, row=row, padx=20, pady=15)

  def _select_file(self):
    file_path = filedialog.askopenfilename(filetypes=[("Excel Spreadsheet", ".xlsx")],
                                           title="Select Spreadsheet",
                                           initialdir="./")
    self._file_path.set(file_path)

  @property
  def file_path(self):
    return self._file_path.get()

