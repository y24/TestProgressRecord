import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import sys
from pathlib import Path
from typing import List, Dict, Any, Callable
import re

class ProjectEditorApp:
    def __init__(self, parent=None, callback: Callable[[Dict[str, Any]], None] = None, 
                 initial_files: List[str] = None, project_path: str = None):
        self.parent = parent
        self.callback = callback
        if parent is None:
            self.root = tk.Tk()
            self.root.title("プロジェクトの作成")
            self.root.geometry("900x400")
        else:
            self.root = tk.Toplevel(parent)
            self.root.withdraw()  # まず非表示
            self.root.title("プロジェクト情報編集")
            self.root.geometry("1000x400")
            self.root.transient(parent)  # 親ウィンドウに対してモーダルに
            self.root.grab_set()  # フォーカスを保持
            self.root.update_idletasks()
            parent_x = parent.winfo_x()
            parent_y = parent.winfo_y()
            self.root.geometry(f"1000x400+{parent_x+50}+{parent_y+50}")
            self.root.deiconify()  # ここで表示
        
        self.file_saved = False
        self.current_project_name = ""
        
        self.project_data = {
            "project": {
                "project_name": "",
                "files": [],
                "excel_path": ""
            }
        }
        
        self.initial_files = initial_files or []
        self.project_loaded = False
        if project_path:
            self.project_loaded = True
        self.create_widgets()
        
        if project_path:
            self.load_project_from_path(project_path)
        
    def create_widgets(self):
        # メニューバーの作成
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # ファイルメニュー
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="プロジェクトファイルを読込", command=self.load_project)
        
        # プロジェクト名称
        ttk.Label(self.root, text="プロジェクト名称:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.project_name_entry = ttk.Entry(self.root, width=50)
        self.project_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # ファイル情報フレーム
        self.files_frame = ttk.LabelFrame(self.root, text="取得元情報")
        self.files_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        
        # ファイル情報追加ボタン
        ttk.Button(self.files_frame, text="Add", command=self.add_file_info).pack(padx=5, pady=5, anchor="w")
        
        # 集計データ書込先
        ttk.Label(self.root, text="集計データ書込先:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.write_path_entry = ttk.Entry(self.root, width=80)
        self.write_path_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.root, text="参照", command=self.select_excel_file).grid(row=3, column=2, padx=5, pady=5)
        
        # ボタンフレーム
        button_frame = ttk.Frame(self.root)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        # 保存ボタン
        ttk.Button(button_frame, text="OK", command=self.save_project).pack(side="left", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=self.root.destroy).pack(side="left", padx=5)
        
        # 初期ファイルがある場合は追加
        if self.initial_files:
            for file in self.initial_files:
                self.add_file_info(file)
        elif not self.project_loaded:
            self.add_file_info()
        
        # グリッドの設定
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        
    def add_file_info(self, initial_file: str = None, file_data: Dict[str, Any] = None):
        frame = ttk.Frame(self.files_frame)
        frame.pack(fill="x", padx=5, pady=5)
        
        # 削除ボタン
        ttk.Button(frame, text="×", width=2, command=lambda: frame.destroy()).grid(row=0, column=0, padx=5, pady=2)

        # 識別子
        ttk.Label(frame, text="識別子:").grid(row=0, column=1, padx=5, pady=2)
        identifier_entry = ttk.Entry(frame, width=20)
        identifier_entry.grid(row=0, column=2, padx=5, pady=2)
        if file_data:
            identifier_entry.insert(0, file_data.get("identifier", ""))

        # URL
        ttk.Label(frame, text="URL:").grid(row=0, column=3, padx=5, pady=2)
        url_entry = ttk.Entry(frame, width=100)
        url_entry.grid(row=0, column=4, padx=5, pady=2)
        if file_data:
            url_entry.insert(0, file_data.get("url", ""))
        
        # エントリを保持
        frame.entries = {
            "identifier": identifier_entry,
            "url": url_entry
        }
        
    def select_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Excelファイルを選択",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.write_path_entry.delete(0, tk.END)
            self.write_path_entry.insert(0, file_path)
            
    def load_project_from_path(self, file_path: str):
        """指定されたパスのJSONファイルを読み込む"""
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                project_data = data.get("project", {})
                
            # プロジェクト名を設定
            self.project_name_entry.delete(0, tk.END)
            self.project_name_entry.insert(0, project_data.get("project_name", ""))
            self.current_project_name = project_data.get("project_name", "")  # 現在のプロジェクト名を保存
            
            # Excelファイルパスを設定
            self.write_path_entry.delete(0, tk.END)
            self.write_path_entry.insert(0, project_data.get("excel_path", ""))
            
            # 既存のファイル行をクリア
            for widget in self.files_frame.winfo_children():
                if isinstance(widget, ttk.Frame):
                    widget.destroy()
            
            # ファイル情報を設定
            for file_data in project_data.get("files", []):
                self.add_file_info(file_data=file_data)
            
            self.project_loaded = True
            
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
            
        # プロジェクト名が変更された場合の確認
        if self.current_project_name and project_name != self.current_project_name:
            if not messagebox.askokcancel("確認", "プロジェクト名が変更されています。\n新たにプロジェクトファイルが作成されますが、よろしいですか？"):
                return
            
        # ファイル情報の取得と空の行の削除
        files = []
        frames_to_remove = []
        for frame in self.files_frame.winfo_children():
            if isinstance(frame, ttk.Frame):
                file_info = {
                    "identifier": frame.entries["identifier"].get().strip(),
                    "url": frame.entries["url"].get().strip()
                }
                # ファイル名とURLが必須
                if file_info["url"]:
                    files.append(file_info)
                else:
                    frames_to_remove.append(frame)
                
        # 空の行の入力欄を削除
        if len(frames_to_remove) > 0:
            # 最後の1つを残して削除
            for frame in frames_to_remove[:-1]:
                frame.destroy()
            
        # Excelファイルパスの取得
        excel_path = self.write_path_entry.get().strip()

        # projectsディレクトリの作成
        projects_dir = Path("projects")
        projects_dir.mkdir(exist_ok=True)
        
        # ファイル名のサニタイズ
        safe_project_name = self.sanitize_filename(project_name)
        
        # JSONファイルの保存
        json_path = projects_dir / f"{safe_project_name}.json"
        
        # 既存のJSONファイルがある場合は読み込む
        existing_data = {}
        if json_path.exists():
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    existing_data = json.load(f)
            except Exception as e:
                messagebox.showerror("エラー", f"既存のプロジェクトファイルの読み込みに失敗しました: {str(e)}")
                return

        # JSONファイルに保存するプロジェクトデータ
        json_project_data = {
            "project_name": project_name,
            "files": files,
            "excel_path": excel_path
        }

        # 既存のデータを保持しつつ、projectキーのデータを更新
        existing_data["project"] = json_project_data
        
        # JSONファイルの保存
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(existing_data, f, ensure_ascii=False, indent=2)
            
        messagebox.showinfo("成功", f"保存しました: {safe_project_name}")
        self.file_saved = True
        self.current_project_name = project_name  # 現在のプロジェクト名を更新

        if self.callback:
            try:
                # データのバリデーション
                if not project_name:
                    raise ValueError("プロジェクト名が入力されていません")
                
                # コールバック用のデータ（project_pathを含む）
                callback_data = {
                    **json_project_data,
                    "project_path": str(json_path)  # コールバックでのみ渡す
                }
                
                # コールバックでデータを返す
                self.callback(callback_data)
            except Exception as e:
                messagebox.showerror("エラー", f"データの保存に失敗しました: {str(e)}")
                return

        # 保存後にウインドウを閉じる
        self.root.destroy()
        
    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()

    def on_closing(self):
        if self.file_saved and self.callback:
            try:
                # プロジェクトデータを収集
                project_data = {
                    "project_name": self.project_name_entry.get().strip(),
                    "files": self.get_file_info(),
                    "excel_path": self.write_path_entry.get().strip()
                }
                
                # データのバリデーション
                if not project_data["project_name"]:
                    raise ValueError("プロジェクト名が入力されていません")
                
                # コールバックでデータを返す
                self.callback(project_data)
                
            except Exception as e:
                messagebox.showerror("エラー", f"データの保存に失敗しました: {str(e)}")
                return
                
        self.root.destroy()
        
    def get_file_info(self) -> List[Dict[str, str]]:
        """現在のファイル情報を取得する"""
        files = []
        for frame in self.files_frame.winfo_children():
            if isinstance(frame, ttk.Frame):
                file_info = {
                    "identifier": frame.entries["identifier"].get().strip(),
                    "url": frame.entries["url"].get().strip()
                }
                if file_info["url"]:  # URLが入力されている場合のみ追加
                    files.append(file_info)
        return files

if __name__ == "__main__":
    app = ProjectEditorApp()
    app.run()
