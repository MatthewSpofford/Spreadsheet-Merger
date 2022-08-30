from tkinter import *
from tkinter import messagebox
from tkinter.ttk import *

import merger.merger
from merger import app
from merger import merger
from merger._gui import merge_config
from merger.merger import NonblockingMerger


class LoadingScreen(Frame):

    def __init__(self, root, nonblocking_merger: NonblockingMerger):
        super().__init__(root, padding=("20", "20", "20", "20"))
        self.root = root
        self.merger = nonblocking_merger

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._progress = IntVar()
        self._progress_bar = Progressbar(self, orient=HORIZONTAL, variable=self._progress,
                                         maximum=self.merger.get_max_progress())
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
        self._update_progress_handler = app.root.after(50, self.update_progress)

    def cancel(self):
        # Cancel status polling handler
        app.root.after_cancel(self._update_progress_handler)

        # Stop the merge process and switch back to the config menu
        self.merger.stop()
        app.switch_frames(merge_config.MergeConfig)

    def update_progress(self):
        try:
            status = self.merger.get_status()

            # If only string based statuses have been sent, with no numerical values attached
            # then update the states label accordingly
            if isinstance(status, str):
                status_str = {
                    merger.MessageStatus.INIT: "Initializing merging processing...",
                    merger.MessageStatus.SAVING: "Saving merged file...",
                    merger.MessageStatus.COMPLETE: "Merge complete!"
                }
                self._status_text.set(status_str[status])

            # Update the progress text according the grouped message data
            if isinstance(status, merger.MessageGroup):
                status_prefix_str = {
                    merger.MessageStatus.INDEXING: "Indexed Rows:\n",
                    merger.MessageStatus.MERGING: "Merged Rows:\n",
                }

                self._status_text.set(status_prefix_str[status[0]] + f"{status[1]}/{status[2]}")
                self._progress.set((status[1] * status[4]) + status[3])

            # If the merge process successfully completed, notify the user
            if status == merger.MessageStatus.COMPLETE:
                # Set the progress bar to 100% just in case it was not completely filled
                self._progress.set(self._progress_bar["maximum"])

                # Display completion dialog box
                messagebox.showinfo("Merge Complete", "Spreadsheets merged successfully!")
                self.cancel()
            # Otherwise, keep polling to see if the merging process has completed
            else:
                self._update_progress_handler = app.root.after(200, self.update_progress)
        except BaseException as e:
            self._status_text.set("Merging error occurred.")
            app.display_error("Merging Error", e)
            self.cancel()
