import tkinter as tk
from tkinter import filedialog

def select_file(ext:tuple):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(filetypes=[(ext)])

def select_files(ext:tuple):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilenames(filetypes=[(ext)])