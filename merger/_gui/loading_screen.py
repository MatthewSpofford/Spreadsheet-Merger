import traceback
from tkinter import *
from tkinter import messagebox
from tkinter.ttk import *

from merger import app
from merger._gui import error_message, merge_config
from merger.nonblocking_merger import NonblockingMerger, MergeStatus, MergeMessage

PROGRESS_STATUS_DELAY_MS = 50


class LoadingScreen(Frame):
    """
    GUI for displaying the loading progress of the spreadsheet merger, as well as handling starting the merge, and
    polling the merge progress
    """

    def __init__(self, root, nonblocking_merger: NonblockingMerger):
        super().__init__(root, padding=("20", "20", "20", "20"))
        self.root = root
        self.merger = nonblocking_merger

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._progress = IntVar()
        self._progress_bar = Progressbar(self, orient=HORIZONTAL, variable=self._progress,
                                         maximum=self.merger.max_progress)
        self._progress_bar.grid(column=0, row=0, padx=20, pady=30, sticky="nsew")

        self._status_text = StringVar()
        self._status_label = Label(self, textvariable=self._status_text, justify="center")
        self._status_label.grid(column=0, row=1, padx=20, pady=0, sticky="ns")

        self._btn_frame = Frame(self)
        self._btn_frame.grid(column=0, row=2, padx=120, pady=20, sticky="nsew")
        self._btn_frame.grid_rowconfigure(0, weight=1)
        self._btn_frame.grid_columnconfigure(0, weight=1)

        self._cancel = Button(self._btn_frame, text="Cancel", command=self.cancel)
        self._cancel.grid(column=0, row=0, padx=5, pady=0, sticky="nsew")

        # Begin the nonblocking merge process
        self.merger.merge()

        # Create a handler to begin the status polling of the merger
        self._update_progress_handler = app.root.after(PROGRESS_STATUS_DELAY_MS, self.update_progress)

    def cancel(self):
        # Cancel status polling handler
        app.root.after_cancel(self._update_progress_handler)

        # Stop the merge process and switch back to the config menu
        self.merger.stop()
        app.switch_frames(merge_config.MergeConfig)

    def update_progress(self):
        try:
            message = self.merger.get_status()

            if isinstance(message, MergeMessage):
                self._status_text.set(message.progress_str)
                self._progress.set(message.progress)

                # If the merge process successfully completed, notify the user
                if message.status == MergeStatus.COMPLETE:
                    # Set the progress bar to 100% just in case it was not completely filled
                    self._progress.set(self._progress_bar["maximum"])

                    # Display completion dialog box
                    messagebox.showinfo("Merge Complete", "Spreadsheets merged successfully!")
                    self.cancel()
                    return

            # Keep polling to see if the merging process has completed
            self._update_progress_handler = app.root.after(PROGRESS_STATUS_DELAY_MS, self.update_progress)
        except BaseException:
            self._status_text.set("Merging error occurred.")
            error_message.display_fatal_error("Merging Error", traceback.format_exc())
            self.cancel()
