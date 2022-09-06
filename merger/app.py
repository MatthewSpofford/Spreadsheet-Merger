import logging
import multiprocessing
from tkinter import *
from tkinter import messagebox

import merger
from ._gui.merge_config import MergeConfig


root = Tk()


def init():
    multiprocessing.freeze_support()

    root.title(f"Spreadsheet Merger (v{merger.__version__})")
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

    # Copy the children names so that an exception isn't thrown when the list of children changes size
    children_names = dict.fromkeys(root.children.keys()).keys()
    for name in children_names:
        root.children[name].destroy()

    gui: Frame = new_frame(root, *additional_params)
    gui.grid(sticky="nsew")


def display_error(title: str, e: BaseException):
    logging.exception(title)
    messagebox.showerror(title, str(e))
