import logging
import traceback
from tkinter import messagebox


def display_error(title: str, message: BaseException):
    logging.exception(title)
    messagebox.showerror(title, str(message))


def display_fatal_error(title: str, message: BaseException):
    display_error(title, message)
