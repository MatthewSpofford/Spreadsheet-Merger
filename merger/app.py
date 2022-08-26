from tkinter import *

import merger
from ._gui.loading_screen import LoadingScreen
from ._gui.merge_config import MergeConfig

root = Tk()


def init():
    root.title(f"Spreadsheet Merger (v{merger.__version__})")
    root.minsize(width=400, height=200)
    root.resizable(width=False, height=False)
    root.grid_columnconfigure(0, weight=1)
    root.grid_rowconfigure(0, weight=1)

    switch_frame(MergeConfig)

    root.mainloop()


def switch_frame(new_frame, *additional_params):
    root.unbind("<Return>")

    # Copy the children names so that an exception isn't thrown when the list of children changes size
    children_names = dict.fromkeys(root.children.keys()).keys()
    for name in children_names:
        root.children[name].destroy()

    gui: Frame = new_frame(root, *additional_params)
    gui.grid(sticky="nsew")
