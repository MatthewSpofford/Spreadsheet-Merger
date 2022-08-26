from tkinter import *
from tkinter.ttk import *


class LoadingScreen(Frame):

    _label_text_prefix = "Number of Rows Remaining:\n"

    def __init__(self, root):
        super().__init__(root, padding=("20", "20", "20", "20"))
        self._root = root

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.curr_rows = IntVar()
        self._progress = Progressbar(self, orient=HORIZONTAL, variable=self.curr_rows, maximum=200,
                                     mode="determinate")
        self._progress.grid(column=0, row=0, padx=20, pady=30, sticky="nsew")

        self._label_text = StringVar()
        self._label_text.set(LoadingScreen._label_text_prefix + "0/200")
        self._remaining_label = Label(self, textvariable=self._label_text, justify="center")
        self._remaining_label.grid(column=0, row=1, padx=20, pady=0, sticky="ns")

        self._btn_frame = Frame(self)
        self._btn_frame.grid(column=0, row=2, padx=120, pady=20, sticky="nsew")
        self._btn_frame.grid_rowconfigure(0, weight=1)
        self._btn_frame.grid_columnconfigure(0, weight=1)

        self._cancel = Button(self._btn_frame, text="Cancel", command=lambda: self.curr_rows.set(self.curr_rows.get() + 10))
        self._cancel.grid(column=0, row=0, padx=5, pady=0, sticky="nsew")

        # self._continue = Button(self._btn_frame, text="Continue")
        # self._continue.config(state=DISABLED)
        # self._continue.grid(column=1, row=0, padx=5, pady=0, sticky="nsew")

        self.create
