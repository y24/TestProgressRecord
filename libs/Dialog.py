import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def select_file(ext:tuple):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(filetypes=[(ext)])

def select_files(ext:tuple):
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilenames(filetypes=[(ext)])

def show_warning(title, message):
    messagebox.showwarning(title=title, message=message)

def show_info(title, message):
    messagebox.showinfo(title=title, message=message)