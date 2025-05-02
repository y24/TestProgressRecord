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

def show_messagebox(root, type:str, title:str, message:str):
    # 一時的な Toplevel を作成（非表示）
    temp_window = Toplevel(root if root else None)
    temp_window.withdraw()  # ウィンドウを非表示にする
    
    # rootが指定されている場合のみ位置を指定
    if root:
        temp_window.geometry(f"1x1+{root.winfo_x()+50}+{root.winfo_y()+50}")  # 位置を指定

    # そのウィンドウを親にして messagebox を表示
    if type == "info":
        messagebox.showinfo(title, message, parent=temp_window)
    elif type == "warning":
        messagebox.showwarning(title=title, message=message, parent=temp_window)
    elif type == "error":
        messagebox.showerror(title=title, message=message, parent=temp_window)
    else:
        print("Error: 該当するタイプなし")

    # 一時ウィンドウを削除
    temp_window.destroy()

def ask_question(root, title:str, message:str):
    # 一時的な Toplevel を作成（非表示）
    temp_window = Toplevel(root if root else None)
    temp_window.withdraw()  # ウィンドウを非表示にする
    
    # rootが指定されている場合のみ位置を指定
    if root:
        temp_window.geometry(f"1x1+{root.winfo_x()+50}+{root.winfo_y()+50}")  # 位置を指定

    # メッセージボックスを表示
    response = messagebox.askquestion(title=title, message=message, parent=temp_window)

    # 一時ウィンドウを削除
    temp_window.destroy()

    return response

def ask_yes_no_cancel(root, title:str, message:str):
    # 一時的な Toplevel を作成（非表示）
    temp_window = Toplevel(root if root else None)
    temp_window.withdraw()  # ウィンドウを非表示にする
    
    # rootが指定されている場合のみ位置を指定
    if root:
        temp_window.geometry(f"1x1+{root.winfo_x()+50}+{root.winfo_y()+50}")  # 位置を指定

    # メッセージボックスを表示
    response = messagebox.askyesnocancel(title=title, message=message, parent=temp_window)

    # 一時ウィンドウを削除
    temp_window.destroy()

    return response
