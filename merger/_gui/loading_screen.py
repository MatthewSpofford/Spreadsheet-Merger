from tkinter import *
from tkinter.ttk import *

import merger.merger
from merger import app
from merger import merger
from merger._gui.merge_config import MergeConfig
from merger.merger import NonblockingMerger


class LoadingScreen(Frame):

    def __init__(self, root, merger: NonblockingMerger):
        super().__init__(root, padding=("20", "20", "20", "20"))
        self.root = root
        self.merger = merger

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.curr_rows = IntVar()
        self._progress = Progressbar(self, orient=HORIZONTAL, variable=self.curr_rows, maximum=200,
                                     mode="determinate")
        self._progress.grid(column=0, row=0, padx=20, pady=30, sticky="nsew")

        self._label_text = StringVar()
        self._remaining_label = Label(self, textvariable=self._label_text, justify="center")
        self._remaining_label.grid(column=0, row=1, padx=20, pady=0, sticky="ns")

        self._btn_frame = Frame(self)
        self._btn_frame.grid(column=0, row=2, padx=120, pady=20, sticky="nsew")
        self._btn_frame.grid_rowconfigure(0, weight=1)
        self._btn_frame.grid_columnconfigure(0, weight=1)

        self._cancel = Button(self._btn_frame, text="Cancel", command=self.cancel)
        self._cancel.grid(column=0, row=0, padx=5, pady=0, sticky="nsew")

        # self._continue = Button(self._btn_frame, text="Continue")
        # self._continue.config(state=DISABLED)
        # self._continue.grid(column=1, row=0, padx=5, pady=0, sticky="nsew")

        # Begin the nonblocking merge process
        self.merger.merge()

    def cancel(self):
        # Stop the merge process and switch back to the config menu
        self.merger.stop()
        app.switch_frames(MergeConfig)


    _merging_text_prefix = "Number of Rows Completed:\n"

    def update_progress(self):
        status = self.merger.get_status()

        if isinstance(status, Exception):
            self._label_text.set("Merging error occurred.")
            app.display_error("Merging Error", status)
            self.cancel()
            return

        elif isinstance(status, str):
            status_str = {
                merger.MessageStatus.INIT: "Initializing merging processing...",
                merger.MessageStatus.SAVING: "Saving merged file...",
                merger.MessageStatus.COMPLETE: "Merge complete!"
            }
            self._label_text.set(status_str[status])

        elif isinstance(status, merger.MessageGroup) and status[0] == merger.MessageStatus.MERGING:
            self._label_text.set(self._merging_text_prefix + f"{status[1]}/{status[1]}")


        app.after(20, self.update_progress)
