import tkinter as tk
from tkinter import ttk, filedialog
import csv, pprint
from datetime import datetime

import WriteData
from libs import Utility
from libs import AppConfig
from libs import Dialog

SETTINGS = {"filepath": "", "table_name": "DATA"}
master_result = ["Pass", "Fixed", "Suspend", "N/A", "Completed"]

def create_treeview(parent, data, structure):
    columns = []
    if structure == 'by_env':
        all_keys = set()
        for dates in data.values():
            for values in dates.values():
                all_keys.update(values.keys())
        columns = ["環境名", "日付"] + Utility.sort_by_master(master_list=master_result, input_list=all_keys)
        data = dict(Utility.sort_nested_dates_desc(data))
    elif structure == 'by_name':
        columns = ["日付", "担当者", "実施数"]
        data = dict(sorted(data.items(), reverse=True))
    elif structure == 'total':
        all_keys = set()
        for values in data.values():
            all_keys.update(values.keys())
        columns = ["日付"] + Utility.sort_by_master(master_list=master_result, input_list=all_keys)
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
        tree.column(col, anchor='center', width=100)
    
    today = datetime.today().strftime("%Y-%m-%d")
    row_colors = {}
    alternating_colors = ["#ffffff", "#f0f0f0"]
    highlight_color = "#d0f0ff"
    
    if structure == 'total':
        for index, (date, values) in enumerate(data.items()):
            bg_color = alternating_colors[index % 2]
            row = [date] + [values.get(k, 0) for k in sorted(all_keys)]
            item_id = tree.insert('', 'end', values=row, tags=(date,))
            tree.tag_configure(date, background=highlight_color if date == today else bg_color)
    elif structure == 'by_env':
        if not data:
            tree.insert('', 'end', values=["取得失敗", "-"] + ["-" for _ in sorted(all_keys)])
        else:
            for index, (env, dates) in enumerate(data.items()):
                bg_color = alternating_colors[index % 2]
                row_colors[env] = bg_color
                for date, values in dates.items():
                    row = [env, date] + [values.get(k, 0) for k in sorted(all_keys)]
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
    
    save_button = ttk.Button(frame, text="CSV保存", command=lambda: save_to_csv(tree, columns))
    save_button.pack(pady=5)
    
    return tree

def save_to_csv(tree, columns):
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if not file_path:
        return
    
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(columns)
        for item in tree.get_children():
            writer.writerow(tree.item(item)['values'])

def update_display(selected_file, data_files):
    current_tab = notebook.index(notebook.select()) if notebook.tabs() else 0
    
    for widget in notebook.winfo_children():
        widget.destroy()
    
    data = next(item for item in data_files if item['file'] == selected_file)

    frame_total = ttk.Frame(notebook)
    notebook.add(frame_total, text="合計")
    create_treeview(frame_total, data['total'], 'total')

    frame_env = ttk.Frame(notebook)
    notebook.add(frame_env, text="環境別")
    create_treeview(frame_env, data['by_env'], 'by_env')
    
    frame_name = ttk.Frame(notebook)
    notebook.add(frame_name, text="担当者別")
    create_treeview(frame_name, data['by_name'], 'by_name')
    
    notebook.select(current_tab)

def convert_to_2d_array(data):
    header = ["ファイル名", "環境名", "日付", "Completed"]
    result = [header]
    for entry in data:
        file_name = entry.get("file", "")
        if entry["by_env"]:
            for env, env_data in entry.get("by_env", {}).items():
                for date, values in env_data.items():
                    result.append([file_name, env, date, values.get("Completed", 0)])
        else:
            # 環境別データがない場合は環境名を !取得失敗! として出力
            for date, values in entry.get("total", {}).items():
                result.append([file_name, "!取得失敗!", date, values.get("Completed", 0)])
    return result

def write_data(field_data):
    SETTINGS = {
        "filepath": field_data["filepath"].get(),
        "table_name": field_data["table_name"].get()
    }

    # データを書込用に変換
    converted_data = convert_to_2d_array(INPUT_DATA)

    # データ書込
    try:
        WriteData.update_table(converted_data, SETTINGS["filepath"], SETTINGS["table_name"])
    except Exception as e:
        Dialog.show_warning("Error", f"保存失敗：ファイルが読み取り専用の可能性があります。\n{e}")

    # 設定を保存
    AppConfig.save_settings(SETTINGS)

def select_write_file(entry):
    filepath = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if filepath:  # キャンセルで空文字が返ってきたときは変更しない
        entry.delete(0, tk.END)  # 既存の内容をクリア
        entry.insert(0, filepath)  # 新しいファイルパスをセット

def create_input_area(parent, settings):
    input_frame = ttk.LabelFrame(parent, text="データ書込設定")
    input_frame.pack(fill=tk.X, padx=5, pady=5)
    
    ttk.Label(input_frame, text="ファイルパス:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
    file_path_entry = ttk.Entry(input_frame, width=50)
    file_path_entry.insert(0, settings["filepath"])
    file_path_entry.grid(row=0, column=1, padx=5, pady=2)
    ttk.Button(input_frame, text="選択", command=lambda: select_write_file(file_path_entry)).grid(row=0, column=2, padx=5, pady=2)
    
    field_frame = ttk.Frame(input_frame)
    field_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=2, sticky=tk.W)

    ttk.Label(field_frame, text="データシート名:").pack(side=tk.LEFT, pady=5)
    table_name_entry = ttk.Entry(field_frame, width=20)
    table_name_entry.insert(0, settings["table_name"])
    table_name_entry.pack(side=tk.LEFT)

    field_data = {"filepath": file_path_entry, "table_name": table_name_entry}

    submit_frame = ttk.Frame(parent)
    submit_frame.pack(fill=tk.BOTH)
    ttk.Button(submit_frame, text="書き込み", command=lambda: write_data(field_data)).pack(pady=(0,5))

def load_data(data_files):
    global notebook, file_selector, INPUT_DATA

    INPUT_DATA = data_files

    root = tk.Tk()
    root.title("Viwer")
    root.geometry("600x500")
    
    # ファイル選択プルダウン
    file_selector = ttk.Combobox(root, values=[file['file'] for file in data_files], state="readonly")
    file_selector.pack(fill=tk.X, padx=5, pady=5)
    file_selector.bind("<<ComboboxSelected>>", lambda event: update_display(file_selector.get(), data_files))
    
    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True)
    
    if data_files:
        file_selector.current(0)
        update_display(data_files[0]['file'], data_files)
    
    # 設定読み込み
    settings = AppConfig.load_settings() or SETTINGS
    # ファイル書き込みエリア
    create_input_area(root, settings)

    root.protocol("WM_DELETE_WINDOW", root.quit)  # アプリ終了時に後続処理を継続
    root.mainloop()
    root.destroy()
