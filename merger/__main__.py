from tkinter import *
from tkinter.ttk import *
from .gui import MergerGUI


__version__ = "1.0.0-rc.2"


def main():
    root = Tk()
    root.title(f"Spreadsheet Merger (v{__version__})")
    root.resizable(width=False, height=False)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    gui = MergerGUI(root)
    gui.grid(column=0, row=0, sticky=(N, W, E, S))

    root.bind("<Return>", gui.merge_spreadsheets)

    root.mainloop()


if __name__ == "__main__":
    main()
