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
                 initial_files: List[Dict[str, Any]] = None, project_path: str = None,
                 gathered_data: List[Dict[str, Any]] = None):
        self.parent = parent
        self.callback = callback
        window_title = "プロジェクト設定"
        window_size = "1040x600"
        if parent is None:
            self.root = tk.Tk()
            self.root.title(window_title)
            self.root.geometry(window_size)
        else:
            self.root = tk.Toplevel(parent)
            self.root.withdraw()  # まず非表示
            self.root.title(window_title)
            self.root.geometry(window_size)
            self.root.transient(parent)  # 親ウィンドウに対してモーダルに
            self.root.grab_set()  # フォーカスを保持
            self.root.update_idletasks()
            parent_x = parent.winfo_x()
            parent_y = parent.winfo_y()
            self.root.geometry(f"{window_size}+{parent_x+50}+{parent_y+50}")
            self.root.deiconify()  # ここで表示
        
        self.file_saved = False
        self.project_path = project_path
        self.gathered_data = gathered_data
        
        self.project_data = {
            "project": {
                "project_name": "",
                "files": [],
                "write_path": "",
                "data_sheet_name": ""
            }
        }
        
        self.initial_files = initial_files or []
        self.project_loaded = False
        self.create_widgets()
        
        if project_path:
            # プロジェクトファイルを開いている場合はロード
            self.load_project_from_path(project_path)
        elif gathered_data:
            # ファイル指定で開いている(名称未設定)場合は集計データからファイル情報を追加
            for data in gathered_data:
                if "filepath" in data:
                    self.add_file_info(file_data={
                        "type": "local",
                        "identifier": data.get("file", ""),
                        "path": data["filepath"]
                    })
        
    def create_widgets(self):
        # メニューバーの作成
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # ファイルメニュー
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="プロジェクトファイルを読込", command=self.load_project)
        
        # プロジェクト名称
        ttk.Label(self.root, text="プロジェクト名称(*):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.project_name_entry = ttk.Entry(self.root, width=60)
        self.project_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # ファイル情報フレーム
        self.files_frame = ttk.LabelFrame(self.root, text="取得元ファイルパス/URL")
        self.files_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        
        # ファイル情報追加ボタン
        ttk.Button(self.files_frame, text="Add", command=self.add_file_info).pack(padx=5, pady=5, anchor="w")

        # ボタンフレーム
        button_frame = ttk.Frame(self.root)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        # 保存ボタン
        ttk.Button(button_frame, text="保存", command=self.save_project).pack(side="left", padx=5)
        ttk.Button(button_frame, text="キャンセル", command=self.root.destroy).pack(side="left", padx=5)
        
        # 初期ファイルがある場合は追加
        if self.initial_files:
            for file in self.initial_files:
                self.add_file_info(file)
        
        # グリッドの設定
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        
    def add_file_info(self, file_data: Dict[str, Any] = None):
        frame = ttk.Frame(self.files_frame)
        frame.pack(fill="x", padx=5, pady=5)
        
        # 削除ボタン
        ttk.Button(frame, text="×", width=2, command=lambda: frame.destroy()).grid(row=0, column=0, padx=5, pady=2)

        # 識別子
        ttk.Label(frame, text="識別子:").grid(row=0, column=1, padx=4, pady=2)
        identifier_entry = ttk.Entry(frame, width=20)
        identifier_entry.grid(row=0, column=2, padx=2, pady=2)
        if file_data:
            identifier_entry.insert(0, file_data.get("identifier", ""))

        # ファイルタイプ選択
        file_type_var = tk.StringVar()  # 初期値は空でOK
        file_type_combo = ttk.Combobox(frame, textvariable=file_type_var, width=10, state="readonly")
        file_type_combo["values"] = ("local", "sharepoint")
        file_type_combo.grid(row=0, column=3, padx=(5,2), pady=0)

        # ファイルパス/URL
        ttk.Label(frame, text="パスまたはURL:").grid(row=0, column=4, padx=2, pady=2)
        path_entry = ttk.Entry(frame, width=80)
        path_entry.grid(row=0, column=5, padx=2, pady=2)
        if file_data:
            path_entry.insert(0, file_data.get("path", ""))

        # ファイル選択ボタン
        def select_file():
            file_path = filedialog.askopenfilename(
                title="Excelファイルを選択",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if file_path:
                path_entry.delete(0, tk.END)
                path_entry.insert(0, file_path)

        def update_file_button_visibility(*args):
            if file_type_var.get() == "local":
                file_button.grid()
            else:
                file_button.grid_remove()

        # ファイル選択ボタン
        file_button = ttk.Button(frame, text="...", width=3, command=select_file)
        file_button.grid(row=0, column=6, padx=2, pady=2)

        # ファイルタイプの初期値をセット
        if file_data and "type" in file_data:
            file_type_var.set(file_data["type"])
        else:
            # 初期値はローカル
            file_type_var.set("local")

        # 初期表示時のボタンの表示/非表示を設定
        update_file_button_visibility()
        
        # ラジオボタンの選択が変更された時のイベントハンドラを設定
        file_type_var.trace_add("write", update_file_button_visibility)
        
        # エントリを保持
        frame.entries = {
            "type": file_type_var,
            "identifier": identifier_entry,
            "path": path_entry
        }
            
    def load_project_from_path(self, file_path: str):
        """指定されたパスのJSONファイルを読み込む"""
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                project_data = data.get("project", {})
                
            # プロジェクト名を設定
            self.project_name_entry.delete(0, tk.END)
            self.project_name_entry.insert(0, project_data.get("project_name", ""))
            
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
        # JSONファイルを選択
        file_path = filedialog.askopenfilename(
            title="プロジェクトファイルを選択",
            filetypes=[("JSON files", "*.json")],
            initialdir="projects"
        )
        
        # キャンセル時は中断
        if not file_path: return

        # 既存のファイル情報をクリア
        for widget in self.files_frame.winfo_children():
            widget.destroy()

        # プロジェクトファイルを読み込む
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
                    "type": frame.entries["type"].get(),
                    "identifier": frame.entries["identifier"].get().strip(),
                    "path": frame.entries["path"].get().strip()
                }
                # ファイルパスが必須
                if file_info["path"]:
                    files.append(file_info)
                else:
                    frames_to_remove.append(frame)
                
        # 空の行の入力欄を削除
        if len(frames_to_remove) > 0:
            # 最後の1つを残して削除
            for frame in frames_to_remove[:-1]:
                frame.destroy()

        # 保存先のパスを決定
        if self.project_path:
            json_path = Path(self.project_path)
        else:
            # projectsディレクトリの作成
            projects_dir = Path("projects")
            projects_dir.mkdir(exist_ok=True)
            
            # ファイル名のサニタイズ
            safe_project_name = self.sanitize_filename(project_name)
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
            "files": files
        }

        # 既存のデータを保持しつつ、projectキーのデータを更新
        existing_data["project"] = json_project_data
        
        # 親ウインドウの集計データが渡されている場合は一緒に保存
        if self.gathered_data is not None:
            existing_data["gathered_data"] = self.gathered_data
        
        # JSONファイルの保存
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(existing_data, f, ensure_ascii=False, indent=2)
            
        messagebox.showinfo("成功", f"プロジェクトファイルを保存しました。")
        self.file_saved = True

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
                    "files": self.get_file_info()
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
                    "type": frame.entries["type"].get(),
                    "identifier": frame.entries["identifier"].get().strip(),
                    "path": frame.entries["path"].get().strip()
                }
                if file_info["path"]:  # ファイルパスが入力されている場合のみ追加
                    files.append(file_info)
        return files

if __name__ == "__main__":
    app = ProjectEditorApp()
    app.run()
