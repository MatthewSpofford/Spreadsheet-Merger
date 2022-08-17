from tkinter import *
from tkinter.ttk import *
from .gui import MergerGUI


if __name__ == "__main__":
    root = Tk()
    root.title("Spreadsheet Merger")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    gui = MergerGUI(root)
    gui.grid(column=0, row=0, sticky=(N, W, E, S))

    root.bind("<Return>", gui.merge_spreadsheets)

    root.mainloop()
