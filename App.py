import tkinter as tk
from tkinter import ttk, filedialog
import subprocess
import csv, os, pprint
from datetime import datetime

import WriteData
from libs import Utility
from libs import AppConfig
from libs import Dialog

def create_treeview(parent, data, structure, file_name):
    columns = []
    if structure == 'by_env':
        all_keys = set()
        for dates in data.values():
            for values in dates.values():
                all_keys.update(values.keys())
        columns = ["環境名", "日付"] + Utility.sort_by_master(master_list=settings["common"]["results"]+["completed"], input_list=all_keys)
        data = dict(Utility.sort_nested_dates_desc(data))
    elif structure == 'by_name':
        columns = ["日付", "担当者", "Completed"]
        data = dict(sorted(data.items(), reverse=True))
    elif structure == 'total':
        all_keys = set()
        for values in data.values():
            all_keys.update(values.keys())
        columns = ["日付"] + Utility.sort_by_master(master_list=settings["common"]["results"]+["completed"], input_list=all_keys)
        data = dict(sorted(data.items(), reverse=True))

    frame = ttk.Frame(parent)
    frame.pack(fill=tk.BOTH, expand=True)
    
    tree_frame = ttk.Frame(frame)
    tree_frame.pack(fill=tk.BOTH, expand=True)

    tree = ttk.Treeview(tree_frame, columns=columns, show='headings')

    scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar_y.set)
    
    scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
    tree.pack(fill=tk.BOTH, expand=True)

    for col in columns:
        tree.heading(col, text=col)
        # 列幅
        if col == "日付":
            col_w = 100
        elif col == "環境名":
            col_w = 120
        else:
            col_w = 70
        tree.column(col, anchor='center', width=col_w)
    
    today = datetime.today().strftime("%Y-%m-%d")
    row_colors = {}
    alternating_colors = ["#ffffff", "#f0f0f0"]
    highlight_color = "#d0f0ff"
    
    if structure == 'total':
        for index, (date, values) in enumerate(data.items()):
            bg_color = alternating_colors[index % 2]
            row = [date] + [values.get(k, 0) for k in Utility.sort_by_master(master_list=settings["common"]["results"]+["completed"], input_list=all_keys)]
            item_id = tree.insert('', 'end', values=row, tags=(date,))
            tree.tag_configure(date, background=highlight_color if date == today else bg_color)
    elif structure == 'by_env':
        if not data:
            tree.insert('', 'end', values=["取得失敗", "-"] + ["-" for _ in Utility.sort_by_master(master_list=settings["common"]["results"]+["completed"], input_list=all_keys)])
        else:
            for index, (env, dates) in enumerate(data.items()):
                bg_color = alternating_colors[index % 2]
                row_colors[env] = bg_color
                for date, values in dates.items():
                    row = [env, date] + [values.get(k, 0) for k in Utility.sort_by_master(master_list=settings["common"]["results"]+["completed"], input_list=all_keys)]
                    item_id = tree.insert('', 'end', values=row, tags=(date,))
                    tree.tag_configure(date, background=highlight_color if date == today else bg_color)
    elif structure == 'by_name':
        for index, (date, names) in enumerate(data.items()):
            bg_color = alternating_colors[index % 2]
            row_colors[date] = bg_color
            for name, count in names.items():
                item_id = tree.insert('', 'end', values=(date, name, count), tags=(date,))
                tree.tag_configure(date, background=highlight_color if date == today else bg_color)
    
    tree.pack(fill=tk.BOTH, expand=True)
    
    save_button = ttk.Button(frame, text="CSV保存", command=lambda: save_to_csv(tree, columns, file_name, structure))
    save_button.pack(pady=5)
    
    return tree

def save_to_csv(tree, columns, file_name, structure):
    # デフォルトファイル名
    basename = os.path.splitext(file_name)[0]
    tab = settings["write"]["structures"][structure]

    # 保存先の選択
    file_path = filedialog.asksaveasfilename(initialfile=f"{basename}_{tab}", defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if not file_path:
        return

    # 保存
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(columns)
        for item in tree.get_children():
            writer.writerow(tree.item(item)['values'])

    # 保存完了
    response = Dialog.ask(title="保存完了", message=f"CSVデータを保存しました。\n{file_path}\n\nファイルを開きますか？")
    if response == "yes":
        open_file(file_path=file_path)

def update_display(selected_file):
    current_tab = notebook.index(notebook.select()) if notebook.tabs() else 0
    
    for widget in notebook.winfo_children():
        widget.destroy()
    
    data = next(item for item in input_data if item['selector_label'] == selected_file)

    frame_total = ttk.Frame(notebook)
    notebook.add(frame_total, text=settings["write"]["structures"]["total"])
    create_treeview(frame_total, data['total'], 'total', data["file"])

    frame_env = ttk.Frame(notebook)
    notebook.add(frame_env, text=settings["write"]["structures"]["by_env"])
    create_treeview(frame_env, data['by_env'], 'by_env', data["file"])
    
    frame_name = ttk.Frame(notebook)
    notebook.add(frame_name, text=settings["write"]["structures"]["by_name"])
    create_treeview(frame_name, data['by_name'], 'by_name', data["file"])
    notebook.select(current_tab)

def write_data(field_data):    
    file_path = field_data["filepath"].get()
    table_name = field_data["table_name"].get()

    # ファイルパス未入力
    if not file_path:
        Dialog.show_warning("Error", f"書込先のファイルを選択してください。")
        return

    # フィールドの設定値をグローバルに反映
    settings["write"]["filepath"] = file_path
    settings["write"]["table_name"] = table_name

    # データ書込
    try:
        result = WriteData.execute(input_data, file_path, table_name)
    except Exception as e:
        Dialog.show_warning(title="Error", message=f"保存失敗：ファイルが読み取り専用の可能性があります。\n{e}")
        return

    # 成功したら設定を保存
    if result:
        AppConfig.save_settings(settings)
        response = Dialog.ask(title="保存完了", message=f"テーブル'{table_name}'にデータを書き込みました。\n{file_path}\n\nファイルを開きますか？")
        if response == "yes":
            open_file(file_path=file_path)

def select_write_file(entry):
    filepath = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel file", "*.xlsx")])
    if filepath:  # キャンセルで空文字が返ってきたときは変更しない
        entry.delete(0, tk.END)  # 既存の内容をクリア
        entry.insert(0, filepath)  # 新しいファイルパスをセット

def open_file(file_path):
    if os.path.isfile(file_path):
        try:
            os.startfile(file_path)  # Windows
        except AttributeError:
            subprocess.run(["xdg-open", file_path])  # Linux/Mac
    else:
        Dialog.showerror("エラー", "指定されたファイルが見つかりません")

def create_input_area(parent, settings):
    input_frame = ttk.LabelFrame(parent, text="データ書込")
    input_frame.pack(fill=tk.X, padx=5, pady=5)
    
    ttk.Label(input_frame, text="書込先:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=3)
    file_path_entry = ttk.Entry(input_frame, width=50)
    file_path_entry.insert(0, settings["write"]["filepath"])
    file_path_entry.grid(row=0, column=1, padx=5, pady=5)
    ttk.Button(input_frame, text="選択", command=lambda: select_write_file(file_path_entry)).grid(row=0, column=2, padx=5, pady=2)
    
    field_frame = ttk.Frame(input_frame)
    field_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=2, sticky=tk.W)

    ttk.Label(field_frame, text="データシート名:").pack(side=tk.LEFT, pady=3)
    table_name_entry = ttk.Entry(field_frame, width=20)
    table_name_entry.insert(0, settings["write"]["table_name"])
    table_name_entry.pack(side=tk.LEFT, padx=5)

    field_data = {"filepath": file_path_entry, "table_name": table_name_entry}

    submit_frame = ttk.Frame(input_frame)
    submit_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=2, sticky=tk.W)
    ttk.Button(submit_frame, text="データ書込", command=lambda: write_data(field_data)).pack(side=tk.LEFT, pady=5)
    ttk.Button(submit_frame, text="開く", command=lambda: open_file(field_data["filepath"].get())).pack(side=tk.LEFT, padx=5, pady=5)

def load_data(data):
    global notebook, file_selector, input_data, settings

    input_data = data

    root = tk.Tk()
    root.title("TestProgressTracker")
    root.geometry("780x500")

    # 設定読み込み
    settings = AppConfig.load_settings()

    # ファイル書き込みエリア
    create_input_area(root, settings)

    # ファイル選択プルダウン
    file_selector = ttk.Combobox(root, values=[file["selector_label"] for file in input_data], state="readonly")
    file_selector.pack(fill=tk.X, padx=5, pady=5)
    file_selector.bind("<<ComboboxSelected>>", lambda event: update_display(file_selector.get()))
    
    # グリッド表示部分
    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True)
    
    if input_data:
        file_selector.current(0)
        update_display(input_data[0]['selector_label'])

    root.protocol("WM_DELETE_WINDOW", root.quit)  # アプリ終了時に後続処理を継続
    root.mainloop()
    root.destroy()
