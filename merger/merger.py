import os
import signal
from enum import Enum, auto
from multiprocessing import Process, Pipe
from multiprocessing.dummy.connection import Connection
from typing import Optional, NamedTuple, Union, get_args, Tuple

import openpyxl as pyxl
import openpyxl.writer.excel
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet


class MessageStatus(str, Enum):
    INIT = auto()
    INDEXING = auto()
    MERGING = auto()
    SAVING = auto()
    COMPLETE = auto()

    _INFO = auto()


_indexing_scale_down = 0.3
MessageGroup = Tuple
f"""
MessageStatus[0] - status of the merging process\n
int[1] - current progress value (context dependent on the status string)\n
int[2] - current maximum value that is trying to be reached (context dependent on the status string)\n
int[3] - value to shift the current progress value by for representing the overall merge progress
float[4] - scale down value for current progress (context dependent)
"""
MessageType = Union[MessageGroup, str, Exception]


class Merger:
    def __init__(self, main_file_path: str, new_file_path: str, column_key: str, merged_file_name: str):
        self.merged_file_path = None
        self.main_file_path = main_file_path
        self.new_file_path = new_file_path
        self.column_key = column_key
        self.merged_file_name = merged_file_name

        if os.path.realpath(self.new_file_path) == os.path.realpath(self.main_file_path):
            raise Exception("Spreadsheet file paths cannot be identical.")

        self._main_wb: Workbook = pyxl.load_workbook(self.main_file_path)
        self._main_wb.active = 0
        self._main_sheet: Worksheet = self._main_wb.active
        self._main_max_row = self._main_sheet.max_row

        self._new_wb: Workbook = pyxl.load_workbook(self.new_file_path)  # , read_only=True)
        self._new_wb.active = 0
        self._new_sheet: Worksheet = self._new_wb.active
        self._new_max_row = self._new_sheet.max_row

        # Locate column label positions
        self._label_indices = self._locate_labels()
        if self.column_key not in self._label_indices:
            raise Exception(f"The key '{self.column_key}' is not a column label in either of the spreadsheets.")
        if self._label_indices[self.column_key].main_label_index is None:
            raise Exception(f"The key '{self.column_key}' is not a column label in `{self.main_file_path}`.")
        if self._label_indices[self.column_key].new_label_index is None:
            raise Exception(f"The key '{self.column_key}' is not a column label in `{self.new_file_path}`.")

        # May be used later for non-blocking functionality for passing around data about the progress
        self._merger_conn: Optional[Connection] = None

    def _update_status(self, message: MessageType):
        if self._merger_conn is not None:
            self._merger_conn.send(message)

    def merge(self):
        import cProfile, time
        from pstats import SortKey

        start = time.time()
        # with cProfile.Profile() as pr:
        try:
            self._update_status((MessageStatus.MERGING, 0, self._new_max_row - 1, self._main_max_row - 1))
            key_col_indices = self._label_indices[self.column_key]

            main_key_rows = self._map_main_keys(key_col_indices)

            # Find all rows in the main sheet based on the current key value in the new sheet
            new_rows = self._new_sheet.iter_rows(min_row=2, max_row=self._new_max_row)
            for new_key_row_index, new_key_row in enumerate(new_rows):
                new_key_val = new_key_row[key_col_indices.new_label_index].value

                # If a row wasn't found for this key in the main sheet, it must be new and inserted into the sheet
                if new_key_val not in main_key_rows:
                    self._main_sheet.insert_rows(2)
                    main_key_row = next(
                        self._main_sheet.iter_rows(min_row=2, max_row=2, max_col=len(self._label_indices)))
                else:
                    main_key_row = main_key_rows[new_key_val]

                # Update all previous keys in the main sheet with new sheet data
                self._update_row(main_key_row, new_key_row)

                # Send status update about completed row
                self._update_status((MessageStatus.MERGING,
                                     new_key_row_index + 1,
                                     self._new_max_row - 1,
                                     (self._main_max_row - 1) * _indexing_scale_down,
                                     1.0,))

            self._update_status(MessageStatus.SAVING)

            # If the merged file name is blank, assume that the main file will be used instead
            if self.merged_file_name is None or len(self.merged_file_name) == 0:
                self.merged_file_path = self.main_file_path
            else:
                merged_file_dir = os.path.dirname(self.main_file_path)
                merged_file_ext = os.path.splitext(self.main_file_path)[1]
                self.merged_file_path = os.path.join(f"{merged_file_dir}", f"{self.merged_file_name}{merged_file_ext}")

            # Creates a copy
            openpyxl.writer.excel.save_workbook(self._main_wb, self.merged_file_path)
            self._update_status(MessageStatus.COMPLETE)
        except BaseException as e:
            if self._merger_conn is not None:
                self._merger_conn.send(e)
            raise e
        finally:
            self._clean_stop()

        print(time.time() - start)
        # pr.print_stats(SortKey.TIME)

    def _clean_stop(self):
        # Close workbook files
        self._main_wb.close()
        self._new_wb.close()

    class _LabelIndices(NamedTuple):
        main_label_index: int
        new_label_index: Optional[int]

    def _locate_labels(self):
        """Map column label indexes used in new sheet to the main sheet indexes"""
        label_indices = {}

        # Iterate over all the column labels in the main sheet
        for (main_label_cell,) in self._main_sheet.iter_cols(max_row=1):
            # if the main column label is empty, skip it
            if main_label_cell.value is None:
                continue

            # Find new sheet position of current column label being examined in the main sheet
            for (new_label_cell,) in self._new_sheet.iter_cols(max_row=1):

                # if the new column label is empty, skip it
                if new_label_cell.value is None:
                    continue

                # Keep track of where the label exists in each
                if main_label_cell.value.lower() == new_label_cell.value.lower():
                    label_indices[main_label_cell.value] = self._LabelIndices(main_label_cell.column - 1,
                                                                              new_label_cell.column - 1)
                    break

            # Column label in main does not exist in new, and should be ignored later on
            if main_label_cell.value not in label_indices:
                label_indices[main_label_cell.value] = self._LabelIndices(main_label_cell.column, None)

        return label_indices

    # TODO Make sure that the _locate_row_key or whatever it's called doesn't always start at 2,
    #      because that can shift if a new row is inserted over time, and we can make it a bit more optimal
    @staticmethod
    def _locate_key_row(sheet: pyxl.workbook.workbook.Worksheet,
                        column_key_index: int, key_value: str) -> Optional[int]:
        """Using the given key value, output the row data it is located on within the key
        column of the given worksheet"""
        for possible_key_row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
            if str(possible_key_row[column_key_index].value) == key_value:
                return possible_key_row

        # The given key wasn't found
        return None

    def _update_row(self, main_row, new_row):
        for indexes in self._label_indices.values():
            # If label does not exist in the new sheet, just ignore this label
            if indexes.new_label_index is None:
                continue

            main_row[indexes.main_label_index].value = new_row[indexes.new_label_index].value

    def _map_main_keys(self, key_col_indices):
        main_key_rows = {}
        for key_row_index, key_row in enumerate(self._main_sheet.iter_rows(min_row=2)):
            main_key_val = key_row[key_col_indices.main_label_index].value
            main_key_rows[main_key_val] = key_row

            self._update_status((MessageStatus.INDEXING,
                                 key_row_index + 1,
                                 self._main_max_row - 1,
                                 0,
                                 _indexing_scale_down))

        return main_key_rows


class NonblockingMerger(Merger):
    def __init__(self, *args):
        super().__init__(*args)

        self._nonblock_conn, self._merger_conn = Pipe()
        self._merge_proc = Process(target=self._start_merge, daemon=True)

    def merge(self):
        self._merge_proc.start()

    def get_max_progress(self):
        return (self._new_max_row - 1) + (self._main_max_row - 1) * _indexing_scale_down

    def _start_merge(self):
        # Set up a termination handler for the new process in the case it needs to be cancelled
        def termination_handler(sig_num, stack_frame):
            print("STOPPING MERGE PROCESS")
            self._clean_stop()
            exit(1)

        signal.signal(signal.SIGTERM, termination_handler)

        # Send initialization status before merging
        self._update_status(MessageStatus.INIT)

        # Begin the blocking merge process on the new process
        super().merge()

    def stop(self):
        self._merge_proc.terminate()
        self._merge_proc.join()

    def is_stopped(self):
        return self._merge_proc.is_alive()

    def get_status(self) -> Optional[MessageType]:
        message: Optional[MessageType] = None
        while self._nonblock_conn.poll():
            message = self._nonblock_conn.recv()

        # If the message received is not of the expected type, then something has gone wrong
        if not (isinstance(message, get_args(MessageType)) or message is None):
            raise Exception("Received invalid status update during merge process.")

        # An error occurred in the merge process
        if isinstance(message, Exception):
            raise message

        return message
