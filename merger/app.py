from tkinter import *

import merger
from merger.gui.merge_config import MergerGUI

root = Tk()


def init():
    root.title(f"Spreadsheet Merger (v{merger.__version__})")
    root.resizable(width=False, height=False)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    switch_frame(MergerGUI)

    root.mainloop()


def switch_frame(new_frame, *additional_params):
    root.unbind("<Return>")
    root.children.clear()

    gui: Frame = new_frame(root, *additional_params)
    gui.grid(column=0, row=0, sticky=(N, W, E, S))
