import tkinter as tk
from tkinter import ttk
import os
import subprocess
import argparse
from libs import AppConfig

class LauncherApp:
    def __init__(self, args=None):
        self.root = tk.Tk()
        self.root.title("TestTraQ - Launcher")
        self.root.geometry("400x300")
        self.root.resizable(False, False)
        
        # 設定の読み込み
        self.settings = AppConfig.load_settings()
        
        # ウィンドウ位置の復元
        self.window_size = "400x300"
        geometry = f"{self.window_size}{self.settings['app']['window_position']}"
        self.root.geometry(geometry)
        
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # プロジェクトリスト
        ttk.Label(main_frame, text="プロジェクトを選択").pack(anchor=tk.W)
        self.project_list = tk.Listbox(main_frame, selectmode=tk.SINGLE)
        self.project_list.pack(fill=tk.BOTH, expand=True, pady=5)
        self.project_list.bind('<<ListboxSelect>>', self.on_selection_changed)
        self.project_list.bind('<Double-Button-1>', self.on_double_click)
        self.root.bind('<Return>', self.on_enter_pressed)  # Enterキー押下時、選択しているプロジェクトを開く
        self.root.bind('<Up>', self.on_up_pressed)  # 上矢印キー押下時、選択を上に移動
        self.root.bind('<Down>', self.on_down_pressed)  # 下矢印キー押下時、選択を下に移動
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        # 新規プロジェクトボタン
        self.new_project_btn = ttk.Button(
            button_frame,
            text="新規プロジェクト",
            command=self.open_new_project
        )
        self.new_project_btn.pack(side=tk.LEFT, padx=0)
        
        # プロジェクトを開くボタン
        self.open_project_btn = ttk.Button(
            button_frame,
            text="選択したプロジェクトを開く",
            command=self.open_selected_project,
            state=tk.DISABLED
        )
        self.open_project_btn.pack(side=tk.LEFT, padx=5)
        
        # プロジェクトリストの更新
        self.update_project_list()
        
        # ウィンドウ終了時のイベントを設定
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 引数でファイルが指定された場合の処理
        if args and (args.data_files or args.project):
            self.process_command_line_args(args)
            return  # ランチャーウインドウを表示せずに終了

    def process_command_line_args(self, args):
        """コマンドライン引数の処理"""
        # プロジェクトファイルのパスが指定されている場合
        if args.project:
            project_file = args.project
            if os.path.exists(project_file):
                self.open_project_with_args(project_file, args.data_files, args.on_reload)
                return
        
        # データファイルが指定されている場合
        for file in args.data_files:
            if file.endswith('.json'):
                self.open_project_with_args(file, args.data_files, args.on_reload)
                return
        
        # プロジェクトファイルが指定されていない場合は新規プロジェクトとして開く
        self.open_new_project_with_files(args.data_files, args.on_reload)

    def open_project_with_args(self, project_file, data_files, on_reload=False):
        """引数で指定されたプロジェクトファイルを開く"""
        cmd = ["python", "StartProcess.py", "--project", project_file]
        if on_reload:
            cmd.append("--on_reload")
        if data_files:
            cmd.extend(data_files)
        subprocess.Popen(cmd)
        self.root.quit()

    def open_new_project_with_files(self, data_files, on_reload=False):
        """新規プロジェクトとしてファイルを開く"""
        cmd = ["python", "StartProcess.py"]
        if on_reload:
            cmd.append("--on_reload")
        if data_files:
            cmd.extend(data_files)
        subprocess.Popen(cmd)
        self.root.quit()

    def save_window_position(self):
        # ウインドウの位置情報を保存
        geometry = self.root.geometry()
        # geometryから位置情報のみを取得 (例: "400x300+200+100" -> "+200+100")
        position = geometry.split(self.window_size)[1]
        self.settings["app"]["window_position"] = position
        AppConfig.save_settings(settings=self.settings)

    def on_closing(self):
        """ウィンドウ終了時の処理"""
        self.save_window_position()
        self.root.quit()
    
    def update_project_list(self):
        """projectsフォルダ内のJSONファイルをリストに表示"""
        self.project_list.delete(0, tk.END)
        projects_dir = "projects"
        if os.path.exists(projects_dir):
            for file in os.listdir(projects_dir):
                if file.endswith('.json'):
                    self.project_list.insert(tk.END, file)
            
            # 最後に開いたプロジェクトを選択
            last_project = self.settings["app"]["last_opened_project"]
            if last_project:
                try:
                    index = self.project_list.get(0, tk.END).index(last_project)
                    self.project_list.selection_set(index)
                    self.open_project_btn.config(state=tk.NORMAL)
                except ValueError:
                    pass  # 最後に開いたプロジェクトが存在しない場合は何もしない
    
    def on_selection_changed(self, event):
        """プロジェクト選択時の処理"""
        self.open_project_btn.config(
            state=tk.NORMAL if self.project_list.curselection() else tk.DISABLED
        )
    
    def on_double_click(self, event):
        """ダブルクリック時の処理"""
        self.open_selected_project()
    
    def on_enter_pressed(self, event):
        """Enterキー押下時の処理"""
        if self.project_list.curselection():
            self.open_selected_project()
    
    def open_new_project(self):
        """新規プロジェクトを開く"""
        subprocess.Popen(["python", "StartProcess.py"])
        self.root.quit()
    
    def open_selected_project(self):
        """選択したプロジェクトを開く"""
        selection = self.project_list.curselection()
        if selection:
            project_file = os.path.join("projects", self.project_list.get(selection[0]))
            # 最後に開いたプロジェクトを保存
            self.settings["app"]["last_opened_project"] = self.project_list.get(selection[0])
            AppConfig.save_settings(settings=self.settings)
            # StartProcess.pyを起動
            subprocess.Popen(["python", "StartProcess.py", project_file])
            self.root.quit()
    
    def on_up_pressed(self, event):
        """上矢印キーが押されたときの処理"""
        current = self.project_list.curselection()
        if current:
            if current[0] > 0:
                self.project_list.selection_clear(current[0])
                self.project_list.selection_set(current[0] - 1)
                self.project_list.see(current[0] - 1)
        else:
            # 何も選択されていない場合は最後の項目を選択
            last_index = self.project_list.size() - 1
            if last_index >= 0:
                self.project_list.selection_set(last_index)
                self.project_list.see(last_index)
        self.on_selection_changed(None)

    def on_down_pressed(self, event):
        """下矢印キーが押されたときの処理"""
        current = self.project_list.curselection()
        if current:
            if current[0] < self.project_list.size() - 1:
                self.project_list.selection_clear(current[0])
                self.project_list.selection_set(current[0] + 1)
                self.project_list.see(current[0] + 1)
        else:
            # 何も選択されていない場合は最初の項目を選択
            if self.project_list.size() > 0:
                self.project_list.selection_set(0)
                self.project_list.see(0)
        self.on_selection_changed(None)

    def run(self):
        """アプリケーションを実行"""
        self.root.mainloop()

def main():
    # コマンドライン引数の設定
    parser = argparse.ArgumentParser(description="TestTraQ Launcher")
    parser.add_argument("--project", help="プロジェクトファイルのパス")
    parser.add_argument("--on_reload", action="store_true", help="データ再集計時のフラグ")
    parser.add_argument("data_files", nargs="*", help="処理するファイルのパス")
    args = parser.parse_args()

    # 引数がある場合は直接StartProcess.pyを実行
    if args.data_files or args.project:
        cmd = ["python", "StartProcess.py"]
        if args.project:
            cmd.extend(["--project", args.project])
        if args.data_files:
            cmd.extend(args.data_files)
        subprocess.Popen(cmd)
        return

    # 引数がない場合はランチャーを表示
    app = LauncherApp(args)
    app.run()

if __name__ == "__main__":
    main() 