import tkinter as tk
from tkinter import ttk, filedialog
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
        columns = ["環境名", "日付"] + Utility.sort_by_master(master_list=settings["common"]["completed_results"]+["completed"], input_list=all_keys)
        data = dict(Utility.sort_nested_dates_desc(data))
    elif structure == 'by_name':
        columns = ["日付", "担当者", "Completed"]
        data = dict(sorted(data.items(), reverse=True))
    elif structure == 'total':
        all_keys = set()
        for values in data.values():
            all_keys.update(values.keys())
        columns = ["日付"] + Utility.sort_by_master(master_list=settings["common"]["completed_results"]+["completed"], input_list=all_keys)
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
            row = [date] + [values.get(k, 0) for k in Utility.sort_by_master(master_list=settings["common"]["completed_results"]+["completed"], input_list=all_keys)]
            item_id = tree.insert('', 'end', values=row, tags=(date,))
            tree.tag_configure(date, background=highlight_color if date == today else bg_color)
    elif structure == 'by_env':
        if not data:
            tree.insert('', 'end', values=["取得失敗", "-"] + ["-" for _ in Utility.sort_by_master(master_list=settings["common"]["completed_results"]+["completed"], input_list=all_keys)])
        else:
            for index, (env, dates) in enumerate(data.items()):
                bg_color = alternating_colors[index % 2]
                row_colors[env] = bg_color
                for date, values in dates.items():
                    row = [env, date] + [values.get(k, 0) for k in Utility.sort_by_master(master_list=settings["common"]["completed_results"]+["completed"], input_list=all_keys)]
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
    basename = os.path.splitext(file_name)[0]
    tab = settings["write"]["structures"][structure]
    file_path = filedialog.asksaveasfilename(initialfile=f"{basename}_{tab}", defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if not file_path:
        return
    
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(columns)
        for item in tree.get_children():
            writer.writerow(tree.item(item)['values'])

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

def convert_to_2d_array(data):
    results = settings["common"]["results"] + ["Completed"]
    header = ["ファイル名", "環境名", "日付"] + results
    out_arr = [header]
    for entry in data:
        file_name = entry.get("file", "")
        if entry["by_env"]:
            for env, env_data in entry.get("by_env", {}).items():
                for date, values in env_data.items():
                    out_arr.append([file_name, env, date] + [values.get(v, 0) for v in results])
        else:
            # 環境別データがない場合は合計データを使用して、環境名は空で出力
            for date, values in entry.get("total", {}).items():
                out_arr.append([file_name, "", date] + [values.get(v, 0) for v in results])
    return out_arr

def write_data(field_data):
    filepath = field_data["filepath"].get()
    if not filepath:
        Dialog.show_warning("Error", f"書込先のファイルを選択してください。")
        return

    # フィールドの値を取得
    settings["write"]["filepath"] = filepath
    settings["write"]["table_name"] = field_data["table_name"].get()

    # データを書込用に変換
    converted_data = convert_to_2d_array(input_data)

    # データ書込
    try:
        WriteData.update_table(converted_data, settings["write"]["filepath"], settings["write"]["table_name"])
    except Exception as e:
        Dialog.show_warning("Error", f"保存失敗：ファイルが読み取り専用の可能性があります。\n{e}")
        return

    # 設定を保存
    AppConfig.save_settings(settings)

def select_write_file(entry):
    filepath = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel file", "*.xlsx")])
    if filepath:  # キャンセルで空文字が返ってきたときは変更しない
        entry.delete(0, tk.END)  # 既存の内容をクリア
        entry.insert(0, filepath)  # 新しいファイルパスをセット

def create_input_area(parent, settings):
    input_frame = ttk.LabelFrame(parent, text="進捗データ書込")
    input_frame.pack(fill=tk.X, padx=5, pady=5)
    
    ttk.Label(input_frame, text="書込先:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
    file_path_entry = ttk.Entry(input_frame, width=50)
    file_path_entry.insert(0, settings["write"]["filepath"])
    file_path_entry.grid(row=0, column=1, padx=5, pady=5)
    ttk.Button(input_frame, text="選択", command=lambda: select_write_file(file_path_entry)).grid(row=0, column=2, padx=5, pady=2)
    
    field_frame = ttk.Frame(input_frame)
    field_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=2, sticky=tk.W)

    ttk.Label(field_frame, text="データシート名:").pack(side=tk.LEFT, padx=5, pady=5)
    table_name_entry = ttk.Entry(field_frame, width=20)
    table_name_entry.insert(0, settings["write"]["table_name"])
    table_name_entry.pack(side=tk.LEFT)

    field_data = {"filepath": file_path_entry, "table_name": table_name_entry}

    submit_frame = ttk.Frame(parent)
    submit_frame.pack(fill=tk.BOTH)
    ttk.Button(submit_frame, text="書き込み", command=lambda: write_data(field_data)).pack(pady=(0,5))

def load_data(data):
    global notebook, file_selector, input_data, settings

    input_data = data

    root = tk.Tk()
    root.title("TestProgressRecord")
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
