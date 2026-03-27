import tkinter as tk
from tkinter import ttk

from src.ui import App


def main():
    root = tk.Tk()
    style = ttk.Style(root)

    try:
        style.theme_use("clam")
    except Exception:
        pass

    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()