import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import sys
from pathlib import Path
from typing import List, Dict, Any
import re

class ProjectEditorApp:
    def __init__(self, initial_files: List[str] = None, initial_json_path: str = None):
        self.root = tk.Tk()
        self.root.title("ProjectEditor")
        self.root.geometry("1000x400")
        self.file_saved = False
        
        self.project_data = {
            "project_name": "",
            "files": [],
            "excel_path": ""
        }
        
        self.initial_files = initial_files or []
        self.create_widgets()
        
        # 初期JSONファイルが指定されている場合は読み込む
        if initial_json_path:
            self.load_project_from_path(initial_json_path)
        
    def create_widgets(self):
        # メニューバーの作成
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # ファイルメニュー
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="プロジェクトを読込", command=self.load_project)
        file_menu.add_command(label="プロジェクトを保存", command=self.save_project)
        
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
        
        # ボタンフレーム
        button_frame = ttk.Frame(self.root)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        # 保存ボタン
        ttk.Button(button_frame, text="保存", command=self.save_project).pack(side="left", padx=5)
        
        # 初期ファイルがある場合は追加
        if self.initial_files:
            for file in self.initial_files:
                self.add_file_info(file)
        else:
            # 初期ファイルがない場合は1つの空の入力欄を追加
            self.add_file_info()
        
        # グリッドの設定
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        
    def add_file_info(self, initial_file: str = None, file_data: Dict[str, Any] = None):
        frame = ttk.Frame(self.files_frame)
        frame.pack(fill="x", padx=5, pady=5)
        
        # ファイル名
        ttk.Label(frame, text="ファイル名:").grid(row=0, column=0, padx=5, pady=2)
        filename_entry = ttk.Entry(frame, width=40)
        filename_entry.grid(row=0, column=1, padx=5, pady=2)
        if initial_file:
            filename_entry.insert(0, initial_file)
        if file_data:
            filename_entry.insert(0, file_data.get("filename", ""))
        
        # URL
        ttk.Label(frame, text="URL:").grid(row=0, column=2, padx=5, pady=2)
        url_entry = ttk.Entry(frame, width=40)
        url_entry.grid(row=0, column=3, padx=5, pady=2)
        if file_data:
            url_entry.insert(0, file_data.get("url", ""))

        # 識別子
        ttk.Label(frame, text="識別子:").grid(row=0, column=4, padx=5, pady=2)
        identifier_entry = ttk.Entry(frame, width=20)
        identifier_entry.grid(row=0, column=5, padx=5, pady=2)
        if file_data:
            identifier_entry.insert(0, file_data.get("identifier", ""))

        # 削除ボタン
        ttk.Button(frame, text="削除", command=lambda: frame.destroy()).grid(row=0, column=7, padx=5, pady=2)
        
        # エントリを保持
        frame.entries = {
            "filename": filename_entry,
            "identifier": identifier_entry,
            "url": url_entry
        }
        
    def select_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Excelファイルを選択",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.excel_path_entry.delete(0, tk.END)
            self.excel_path_entry.insert(0, file_path)
            
    def load_project_from_path(self, file_path: str):
        """指定されたパスのJSONファイルを読み込む"""
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                project_data = json.load(f)
                
            # プロジェクト名を設定
            self.project_name_entry.delete(0, tk.END)
            self.project_name_entry.insert(0, project_data.get("project_name", ""))
            
            # Excelファイルパスを設定
            self.excel_path_entry.delete(0, tk.END)
            self.excel_path_entry.insert(0, project_data.get("excel_path", ""))
            
            # ファイル情報を設定
            for file_data in project_data.get("files", []):
                self.add_file_info(file_data=file_data)
            
        except Exception as e:
            messagebox.showerror("エラー", f"プロジェクトファイルの読み込みに失敗しました: {str(e)}")
            
    def load_project(self):
        # 既存のファイル情報をクリア
        for widget in self.files_frame.winfo_children():
            widget.destroy()
            
        # JSONファイルを選択
        file_path = filedialog.askopenfilename(
            title="プロジェクトファイルを選択",
            filetypes=[("JSON files", "*.json")],
            initialdir="projects"
        )
        
        if not file_path:
            return
            
        self.load_project_from_path(file_path)
            
    def sanitize_filename(self, filename: str) -> str:
        """ファイル名として不適切な文字を_に置換する"""
        # Windowsで使用できない文字を定義
        invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
        return re.sub(invalid_chars, '_', filename)

    def save_project(self):
        # プロジェクト名の取得
        project_name = self.project_name_entry.get().strip()
        if not project_name:
            messagebox.showerror("エラー", "プロジェクト名称を入力してください")
            return
            
        # ファイル情報の取得と空の行の削除
        files = []
        frames_to_remove = []
        for frame in self.files_frame.winfo_children():
            if isinstance(frame, ttk.Frame):
                file_info = {
                    "filename": frame.entries["filename"].get().strip(),
                    "identifier": frame.entries["identifier"].get().strip(),
                    "url": frame.entries["url"].get().strip()
                }
                # ファイル名とURLが必須
                if file_info["filename"] and file_info["url"]:
                    files.append(file_info)
                else:
                    frames_to_remove.append(frame)
                
        # 空の行の入力欄を削除
        if len(frames_to_remove) > 0:
            # 最後の1つを残して削除
            for frame in frames_to_remove[:-1]:
                frame.destroy()
                
        if not files:
            messagebox.showerror("エラー", "少なくとも1つのファイル情報を入力してください（ファイル名とURLは必須です）")
            return
            
        # Excelファイルパスの取得
        excel_path = self.excel_path_entry.get().strip()
            
        # プロジェクトデータの作成
        self.project_data = {
            "project_name": project_name,
            "files": files,
            "excel_path": excel_path
        }
        
        # projectsディレクトリの作成
        projects_dir = Path("projects")
        projects_dir.mkdir(exist_ok=True)
        
        # ファイル名のサニタイズ
        safe_project_name = self.sanitize_filename(project_name)
        
        # JSONファイルの保存
        json_path = projects_dir / f"{safe_project_name}.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(self.project_data, f, ensure_ascii=False, indent=2)
            
        messagebox.showinfo("成功", f"プロジェクト情報を保存しました: {json_path}")
        self.file_saved = True
        
    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()

    def on_closing(self):
        if self.file_saved:
            messagebox.showinfo(
                "保存完了",
                "プロジェクトファイルが保存されました。\n設定を反映するには、File > ファイルを再読み込み を実行してください。"
            )
        self.root.destroy() 

if __name__ == "__main__":
    # 引数の解析
    args = sys.argv[1:]
    initial_files = []
    initial_json_path = None
    
    for arg in args:
        if arg.endswith('.json'):
            initial_json_path = arg
        else:
            initial_files.append(arg)
            
    app = ProjectEditorApp(initial_files=initial_files if initial_files else None,
                         initial_json_path=initial_json_path)
    app.run()
