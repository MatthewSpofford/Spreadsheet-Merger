import logging
import signal
import sys
import traceback
from dataclasses import dataclass
from enum import Enum, auto
from multiprocessing import Pipe, Process
from pathlib import Path
from typing import Optional, get_args, Union

from merger.merger import Merger

LOG = logging.getLogger(__name__)


class MergeStatus(str, Enum):
    INIT = auto()
    INDEXING = auto()
    MERGING = auto()
    SAVING = auto()
    COMPLETE = auto()


_INDEXING_SCALE_DOWN = 0.3


@dataclass
class MergeMessage:
    status: MergeStatus
    """Merging status"""
    progress: int
    """Current merging progress value"""
    progress_str: str
    """Current merging progress string representation"""


MessageType = Union[MergeMessage, tuple]
"""
Tuple only used if an exception is raised. It follows the structure of sys.exc_info() 
(https://docs.python.org/3/library/sys.html#sys.exc_info), except that the traceback has been converted to a formatted
string
"""


class NonblockingMerger(Merger):
    def __init__(self, main_file_path: Path, new_file_path: Path, column_key: str,
                 merged_file_name: Optional[str] = None):
        super().__init__(main_file_path, new_file_path, column_key, merged_file_name)

        # Used for passing around data about the progress
        self._nonblock_conn, self._merger_conn = Pipe()
        self._merge_proc = Process(target=self._start_merge, daemon=True)

        # Maximum progress is subtracted to account for the header rows on each sheet.
        # The maximum row count for the original sheet is used to determine the indexing progress.
        self._max_indexing_progress = self._original_max_row - self.original_header_row
        # The maximum row count for the new sheet is used to determine the merging progress.
        self._max_merging_progress = self._new_max_row - self.new_header_row

    def merge(self):
        self._merge_proc.start()

    @property
    def max_progress(self):
        return self._max_indexing_progress + self._max_merging_progress

    def _start_merge(self):
        # Set up a termination handler for the new process in the case it needs to be cancelled
        def termination_handler(*_):
            LOG.error("STOPPING MERGE PROCESS")
            self._clean_stop()
            exit(1)

        signal.signal(signal.SIGTERM, termination_handler)

        # Send initialization status before merging
        self._update_status(
            MergeMessage(MergeStatus.INIT, 0, "Initializing merging processing..."))

        # Begin the blocking merge process on the new process
        super().merge()

    def _update_status(self, message: MessageType):
        self._merger_conn.send(message)

    def _hook_initialization(self):
        """
        Hook that runs immediately before the merge process is about to start
        """
        self._update_status(
            MergeMessage(MergeStatus.MERGING, 0, "Indexing the original spreadsheet..."))

    def _hook_row_indexed(self, row_index):
        """
        Hook that runs when a row has been indexed, but not yet merged.

        Args:
            row_index (int): Index of the row that was just indexed pre-merge.
        """
        # The main sheet's max row amount is the "maximum" for indexing progress, because only the main sheet is indexed
        progress = row_index
        message_str = f"Indexed Rows:\n{progress}/{self._max_indexing_progress}"
        self._update_status(
            MergeMessage(MergeStatus.INDEXING, progress, message_str))

    def _hook_row_merged(self, row_index):
        """
        Hook that runs immediately after a row in the spreadsheet has been merged

        Args:
            row_index (int): Index of the row that was just merged
        """
        progress = row_index + 1
        message_str = f"Merged Rows:\n{progress}/{self._max_merging_progress}"
        self._update_status(
            MergeMessage(MergeStatus.MERGING, progress + self._max_indexing_progress, message_str))

    def _hook_pre_saving(self):
        """
        Hook that runs right before saving the merged rows to a new file.
        """
        self._update_status(
            MergeMessage(MergeStatus.SAVING, self.max_progress, "Saving merged file..."))

    def _hook_success(self):
        """
        Hook that runs when the merge process completes successfully.
        """
        self._update_status(
            MergeMessage(MergeStatus.COMPLETE, self.max_progress, "Saving merged file..."))

    def _hook_exception(self):
        exc_info = list(sys.exc_info())
        exc_info[2] = traceback.format_exc()
        self._update_status(tuple(exc_info))

    def stop(self):
        if self._merge_proc.is_alive():
            self._merge_proc.terminate()
        self._merge_proc.join()

    def is_stopped(self):
        return self._merge_proc.is_alive()

    def get_status(self) -> Optional[MergeMessage]:
        message: Optional[MessageType] = None
        while self._nonblock_conn.poll():
            message = self._nonblock_conn.recv()

        # If the message received is not of the expected type, then something has gone wrong
        if not (isinstance(message, get_args(MessageType)) or message is None):
            raise Exception(f"Received invalid status update during merge process. Message was: {message}")

        # An error occurred in the merge process
        # Tuple follows the structure of sys.exc_inf() (https://docs.python.org/3/library/sys.html#sys.exc_info)
        if isinstance(message, tuple):
            exc_type, exc, exc_traceback = message
            raise Exception(f"Failure occurred during merge: {exc}:\n\n{exc_traceback}")

        return message
