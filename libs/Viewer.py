import tkinter as tk
from tkinter import ttk, filedialog
import csv

def create_treeview(parent, data, structure):
    columns = []
    if structure == 'by_env':
        all_keys = set()
        for dates in data.values():
            for values in dates.values():
                all_keys.update(values.keys())
        columns = ["環境名", "日付"] + sorted(all_keys)
    elif structure == 'by_name':
        columns = ["日付", "名前", "数"]
    elif structure == 'total':
        all_keys = set()
        for values in data.values():
            all_keys.update(values.keys())
        columns = ["日付"] + sorted(all_keys)
    
    frame = ttk.Frame(parent)
    frame.pack(fill=tk.BOTH, expand=True)
    
    tree = ttk.Treeview(frame, columns=columns, show='headings')
    
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, anchor='center', width=100)
    
    row_colors = {}
    alternating_colors = ["#ffffff", "#f0f0f0"]
    
    if structure == 'total':
        for index, (date, values) in enumerate(data.items()):
            bg_color = alternating_colors[index % 2]
            row = [date] + [values.get(k, 0) for k in sorted(all_keys)]
            item_id = tree.insert('', 'end', values=row, tags=(date,))
            tree.tag_configure(date, background=bg_color)
    elif structure == 'by_env':
        for index, (env, dates) in enumerate(data.items()):
            bg_color = alternating_colors[index % 2]
            row_colors[env] = bg_color
            for date, values in dates.items():
                row = [env, date] + [values.get(k, 0) for k in sorted(all_keys)]
                item_id = tree.insert('', 'end', values=row, tags=(env,))
                tree.tag_configure(env, background=bg_color)
    elif structure == 'by_name':
        for index, (date, names) in enumerate(data.items()):
            bg_color = alternating_colors[index % 2]
            row_colors[date] = bg_color
            for name, count in names.items():
                item_id = tree.insert('', 'end', values=(date, name, count), tags=(date,))
                tree.tag_configure(date, background=bg_color)
    
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

def load_data(data_files):
    global notebook, file_selector
    root = tk.Tk()
    root.title("Viwer")
    root.geometry("600x450")
    
    file_selector = ttk.Combobox(root, values=[file['file'] for file in data_files], state="readonly")
    file_selector.pack(fill=tk.X, padx=5, pady=5)
    file_selector.bind("<<ComboboxSelected>>", lambda event: update_display(file_selector.get(), data_files))
    
    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True)
    
    if data_files:
        file_selector.current(0)
        update_display(data_files[0]['file'], data_files)
    
    root.protocol("WM_DELETE_WINDOW", root.quit)  # アプリ終了時に後続処理を継続
    root.mainloop()
    root.destroy()
