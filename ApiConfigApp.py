import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from libs.Win32Credential import save_api_key, delete_api_key

class ApiConfigApp:
    def __init__(self, root):
        self.root = root
        self.root.title("APIキー設定")
        self.root.geometry("370x120")
        
        # メインフレーム
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # APIキー入力ラベル
        api_key_label = ttk.Label(main_frame, text="APIキー:")
        api_key_label.grid(row=0, column=0, sticky=tk.W)
        
        # APIキー入力フィールド（パスワード表示）
        self.api_key_entry = ttk.Entry(main_frame, show="*", width=40)
        self.api_key_entry.grid(row=0, column=1, padx=5, pady=5)
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=20)
        
        # OKボタン
        ok_button = ttk.Button(button_frame, text="保存", command=self.save_key)
        ok_button.grid(row=0, column=0, padx=5)
        
        # キャンセルボタン
        cancel_button = ttk.Button(button_frame, text="キャンセル", command=self.root.destroy)
        cancel_button.grid(row=0, column=1, padx=5)
        
        # 削除ボタン
        delete_button = ttk.Button(button_frame, text="APIキー削除", command=self.delete_key)
        delete_button.grid(row=0, column=2, padx=5)
    
    def save_key(self):
        api_key = self.api_key_entry.get()
        if not api_key:
            messagebox.showwarning("必須項目", "APIキーを入力してください。")
            return
        
        try:
            save_api_key(api_key)
            messagebox.showinfo("成功", "APIキーを保存しました。")
            self.root.destroy()
        except Exception as e:
            messagebox.showerror("エラー", f"APIキーの保存に失敗しました: {str(e)}")
    
    def delete_key(self):
        try:
            delete_api_key()
            messagebox.showinfo("成功", "APIキーを削除しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"APIキーの削除に失敗しました: {str(e)}")

def run():
    root = tk.Tk()
    app = ApiConfigApp(root)
    root.mainloop()

if __name__ == "__main__":
    run()
