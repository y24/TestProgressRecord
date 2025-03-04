import tkinter as tk
from tkinter import ttk, filedialog
import csv, json
from datetime import datetime

from libs import Utility

# 設定ファイルのパス
CONFIG_FILE = "appconfig.json"

def create_treeview(parent, data, structure):
    columns = []
    if structure == 'by_env':
        all_keys = set()
        for dates in data.values():
            for values in dates.values():
                all_keys.update(values.keys())
        columns = ["環境名", "日付"] + sorted(all_keys)
        data = dict(Utility.sort_nested_dates_desc(data))
    elif structure == 'by_name':
        columns = ["日付", "担当者", "実施数"]
        data = dict(sorted(data.items(), reverse=True))
    elif structure == 'total':
        all_keys = set()
        for values in data.values():
            all_keys.update(values.keys())
        columns = ["日付"] + sorted(all_keys)
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

def write_data(write_settings):
    print(write_settings)

def create_input_area(parent):
    frame = ttk.LabelFrame(parent, text="データ書込設定")
    frame.pack(fill=tk.X, padx=5, pady=5)
    
    ttk.Label(frame, text="ファイルパス:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
    file_path_entry = ttk.Entry(frame, width=50)
    file_path_entry.grid(row=0, column=1, padx=5, pady=2)
    ttk.Button(frame, text="選択", command=lambda: file_path_entry.insert(0, filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]))).grid(row=0, column=2, padx=5, pady=2)
    
    field_frame = ttk.Frame(frame)
    field_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=2, sticky=tk.W)
    
    ttk.Label(field_frame, text="ヘッダー列:").pack(side=tk.LEFT, padx=5)
    header_col_entry = ttk.Entry(field_frame, width=5)
    header_col_entry.insert(0, "H")
    header_col_entry.pack(side=tk.LEFT)
    
    ttk.Label(field_frame, text="ファイル名列:").pack(side=tk.LEFT, padx=5)
    filename_col_entry = ttk.Entry(field_frame, width=5)
    filename_col_entry.insert(0, "B")
    filename_col_entry.pack(side=tk.LEFT)
    
    ttk.Label(field_frame, text="日付キー:").pack(side=tk.LEFT, padx=5)
    date_key_entry = ttk.Entry(field_frame, width=10)
    date_key_entry.insert(0, "日付")
    date_key_entry.pack(side=tk.LEFT)
    
    ttk.Label(field_frame, text="実績キー:").pack(side=tk.LEFT, padx=5)
    actual_key_entry = ttk.Entry(field_frame, width=10)
    actual_key_entry.insert(0, "実績")
    actual_key_entry.pack(side=tk.LEFT)
    
    date_frame = ttk.Frame(frame)
    date_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=2, sticky=tk.W+tk.E)
    ttk.Label(date_frame, text="開始日:").pack(side=tk.LEFT, padx=5)
    start_date_entry = ttk.Entry(date_frame, width=15)
    start_date_entry.insert(0, "2000-01-01")
    start_date_entry.pack(side=tk.LEFT)
    ttk.Label(date_frame, text="終了日:").pack(side=tk.LEFT, padx=5)
    end_date_entry = ttk.Entry(date_frame, width=15)
    end_date_entry.insert(0, datetime.today().strftime("%Y-%m-%d"))
    end_date_entry.pack(side=tk.LEFT)

    until_today_var = tk.BooleanVar()
    ttk.Checkbutton(date_frame, text="今日まで", variable=until_today_var).pack(side=tk.LEFT, padx=10)


    write_settings = {
        "filepath": file_path_entry.get(),
        "header_row": header_col_entry.get(),
        "filename_row": filename_col_entry.get(),
        "date_row": date_key_entry.get(),
        "actual_row": actual_key_entry.get(),
        "until_today": until_today_var.get(),
        "start_date": start_date_entry.get(),
        "end_date": end_date_entry.get()
    }

    submit_frame = ttk.Frame(parent)
    submit_frame.pack(fill=tk.BOTH)
    ttk.Button(submit_frame, text="書き込み", command=lambda: write_data(write_settings)).pack(pady=(0,5))

def load_data(data_files):
    global notebook, file_selector
    root = tk.Tk()
    root.title("Viwer")
    root.geometry("600x500")
    
    file_selector = ttk.Combobox(root, values=[file['file'] for file in data_files], state="readonly")
    file_selector.pack(fill=tk.X, padx=5, pady=5)
    file_selector.bind("<<ComboboxSelected>>", lambda event: update_display(file_selector.get(), data_files))
    
    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True)
    
    if data_files:
        file_selector.current(0)
        update_display(data_files[0]['file'], data_files)
    
    create_input_area(root)

    root.protocol("WM_DELETE_WINDOW", root.quit)  # アプリ終了時に後続処理を継続
    root.mainloop()
    root.destroy()
