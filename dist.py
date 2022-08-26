import multiprocessing

from merger import app

if __name__ == "__main__":
    multiprocessing.freeze_support()
    app.init()
