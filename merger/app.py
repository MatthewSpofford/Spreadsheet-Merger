import multiprocessing
from tkinter import *

import merger
from ._gui.merge_config import MergeConfig


root = Tk()


def init():
    multiprocessing.freeze_support()

    title_str = f"Spreadsheet Merger (v{merger.__version__})"
    if __debug__:
        title_str += " === DEBUG BUILD ==="
    root.title(title_str)
    root.minsize(width=400, height=200)
    root.resizable(width=False, height=False)
    root.grid_columnconfigure(0, weight=1)
    root.grid_rowconfigure(0, weight=1)

    switch_frames(MergeConfig)

    try:
        root.mainloop()
    except KeyboardInterrupt:
        for child in multiprocessing.process.active_children():
            child.terminate()
            child.join()
        exit(0)


def switch_frames(new_frame, *additional_params):
    root.unbind("<Return>")

    for name in list(root.children.keys()):
        root.children[name].destroy()

    gui: Frame = new_frame(root, *additional_params)
    gui.grid(sticky="nsew")
