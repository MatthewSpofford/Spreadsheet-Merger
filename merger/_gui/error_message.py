import logging
from tkinter import messagebox, Frame, Button, LEFT, ACTIVE, Text
from tkinter.simpledialog import Dialog

from merger import app


def display_error(title: str, message: BaseException):
    logging.exception(title)
    messagebox.showerror(title, str(message))


def display_fatal_error(title: str, traceback: str):
    logging.error(f"Exception '{title}' occurred:\n{traceback}")
    ExceptionDialog(title, traceback)


class ExceptionDialog(Dialog):
    def __init__(self, title, message):
        self.message = message
        super().__init__(app.root, title)

    def body(self, master):
        self.resizable(width=False, height=False)

        text = Text(master)
        text.insert("1.0", self.message)
        text.config(state="disabled")
        text.pack()

    def buttonbox(self):
        box = Frame(self)
        ok_button = Button(box, text="OK", width=10, command=self.ok, default=ACTIVE)
        ok_button.pack(side=LEFT, padx=5, pady=5)

        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.ok)

        box.pack()
