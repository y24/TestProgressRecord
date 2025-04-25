import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import sys
from pathlib import Path
from typing import List, Dict, Any

class CreateProjectApp:
    def __init__(self, initial_files: List[str] = None):
        self.root = tk.Tk()
        self.root.title("プロジェクト作成")
        self.root.geometry("800x600")
        
        self.project_data = {
            "project_name": "",
            "files": [],
            "excel_path": ""
        }
        
        self.initial_files = initial_files or []
        self.create_widgets()
        
    def create_widgets(self):
        # プロジェクト名称
        ttk.Label(self.root, text="プロジェクト名称:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.project_name_entry = ttk.Entry(self.root, width=50)
        self.project_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        # ファイル情報フレーム
        self.files_frame = ttk.LabelFrame(self.root, text="ファイル情報")
        self.files_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        
        # ファイル情報追加ボタン
        ttk.Button(self.root, text="ファイル情報を追加", command=self.add_file_info).grid(row=2, column=0, columnspan=2, pady=5)
        
        # Excelファイルパス
        ttk.Label(self.root, text="書込先Excelファイル:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.excel_path_entry = ttk.Entry(self.root, width=50)
        self.excel_path_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.root, text="参照", command=self.select_excel_file).grid(row=3, column=2, padx=5, pady=5)
        
        # 保存ボタン
        ttk.Button(self.root, text="保存", command=self.save_project).grid(row=4, column=0, columnspan=2, pady=20)
        
        # 初期ファイルがある場合は追加
        if self.initial_files:
            for file in self.initial_files:
                self.add_file_info(file)
        
        # グリッドの設定
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        
    def add_file_info(self, initial_file: str = None):
        frame = ttk.Frame(self.files_frame)
        frame.pack(fill="x", padx=5, pady=5)
        
        # ファイル名
        ttk.Label(frame, text="ファイル名:").grid(row=0, column=0, padx=5, pady=2)
        filename_entry = ttk.Entry(frame, width=30)
        filename_entry.grid(row=0, column=1, padx=5, pady=2)
        if initial_file:
            filename_entry.insert(0, initial_file)
        
        # 識別子
        ttk.Label(frame, text="識別子:").grid(row=0, column=2, padx=5, pady=2)
        identifier_entry = ttk.Entry(frame, width=20)
        identifier_entry.grid(row=0, column=3, padx=5, pady=2)
        
        # URL
        ttk.Label(frame, text="URL:").grid(row=0, column=4, padx=5, pady=2)
        url_entry = ttk.Entry(frame, width=30)
        url_entry.grid(row=0, column=5, padx=5, pady=2)
        
        # 環境別フラグ
        ttk.Label(frame, text="環境別フラグ:").grid(row=0, column=6, padx=5, pady=2)
        env_flag_var = tk.StringVar(value="False")
        env_flag_check = ttk.Checkbutton(frame, variable=env_flag_var, onvalue="True", offvalue="False")
        env_flag_check.grid(row=0, column=7, padx=5, pady=2)
        
        # 削除ボタン
        ttk.Button(frame, text="削除", command=lambda: frame.destroy()).grid(row=0, column=8, padx=5, pady=2)
        
        # エントリを保持
        frame.entries = {
            "filename": filename_entry,
            "identifier": identifier_entry,
            "url": url_entry,
            "env_flag": env_flag_var
        }
        
    def select_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Excelファイルを選択",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.excel_path_entry.delete(0, tk.END)
            self.excel_path_entry.insert(0, file_path)
            
    def save_project(self):
        # プロジェクト名の取得
        project_name = self.project_name_entry.get().strip()
        if not project_name:
            messagebox.showerror("エラー", "プロジェクト名称を入力してください")
            return
            
        # ファイル情報の取得
        files = []
        for frame in self.files_frame.winfo_children():
            if isinstance(frame, ttk.Frame):
                file_info = {
                    "filename": frame.entries["filename"].get().strip(),
                    "identifier": frame.entries["identifier"].get().strip(),
                    "url": frame.entries["url"].get().strip(),
                    "env_flag": frame.entries["env_flag"].get() == "True"
                }
                files.append(file_info)
                
        if not files:
            messagebox.showerror("エラー", "少なくとも1つのファイル情報を入力してください")
            return
            
        # Excelファイルパスの取得
        excel_path = self.excel_path_entry.get().strip()
        if not excel_path:
            messagebox.showerror("エラー", "Excelファイルを選択してください")
            return
            
        # プロジェクトデータの作成
        self.project_data = {
            "project_name": project_name,
            "files": files,
            "excel_path": excel_path
        }
        
        # projectsディレクトリの作成
        projects_dir = Path("projects")
        projects_dir.mkdir(exist_ok=True)
        
        # JSONファイルの保存
        json_path = projects_dir / f"{project_name}.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(self.project_data, f, ensure_ascii=False, indent=2)
            
        messagebox.showinfo("成功", f"プロジェクト情報を保存しました: {json_path}")
        
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    initial_files = sys.argv[1:] if len(sys.argv) > 1 else None
    app = CreateProjectApp(initial_files)
    app.run()
