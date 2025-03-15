import tkinter as tk
from tkinter import filedialog
from tkinter import Tk, messagebox, Toplevel

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

def show_messagebox(root, type:str, title:str, message):
    # 一時的な Toplevel を作成（非表示）
    temp_window = Toplevel(root)
    temp_window.withdraw()  # ウィンドウを非表示にする
    temp_window.geometry(f"1x1+{root.winfo_x()+50}+{root.winfo_y()+50}")  # 位置を指定

    # そのウィンドウを親にして messagebox を表示
    if type == "info":
        messagebox.showinfo(title, message, parent=temp_window)
    elif type == "warn":
        messagebox.showwarning(title=title, message=message, parent=temp_window)
    elif type == "error":
        messagebox.showerror(title=title, message=message, parent=temp_window)

    # 一時ウィンドウを削除
    temp_window.destroy()

def ask(title, message):
    return messagebox.askquestion(title=title, message=message)
