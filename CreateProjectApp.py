from typing import List
import tkinter as tk
from tkinter import messagebox

class CreateProjectApp:
    def __init__(self, initial_files: List[str] = None, initial_json_path: str = None):
        self.root = tk.Tk()
        self.root.title("Project Editor")
        self.root.geometry("1000x400")
        self.file_saved = False  # 保存フラグ
        self.initial_files = initial_files or []
        self.create_widgets()

    def save_project(self):
        # ... 既存のコード ...
        messagebox.showinfo("成功", f"プロジェクト情報を保存しました: {json_path}")
        self.file_saved = True

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()

    def on_closing(self):
        if self.file_saved:
            messagebox.showinfo(
                "保存完了",
                "ファイルが保存されました。\n設定を反映するには、ファイル > ファイルを再読み込み を実行してください。"
            )
        self.root.destroy() 