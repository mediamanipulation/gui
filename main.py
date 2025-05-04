# main.py

import tkinter as tk
from gui.app import ExcelTransformerApp

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelTransformerApp(root)
    root.mainloop()
