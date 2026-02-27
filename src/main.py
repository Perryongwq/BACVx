from tkinter import *
from Pages.MainPage import MainWindow
from version import get_full_version

if __name__ == "__main__":
    print(f"Starting {get_full_version()}")
    root = Tk()
    MainWindow(root)
    root.mainloop()