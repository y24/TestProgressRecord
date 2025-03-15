import tkinter as tk
from tkinter import ttk, filedialog
import subprocess
import csv, os, sys, pprint
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
from collections import defaultdict

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
        columns = ["環境名", "日付"] + Utility.sort_by_master(master_list=settings["common"]["results"]+[settings["common"]["completed"]], input_list=all_keys)
        data = dict(Utility.sort_nested_dates_desc(data))
    elif structure == 'by_name':
        columns = ["日付", "担当者", "Completed"]
        data = dict(sorted(data.items(), reverse=True))
    elif structure == 'daily':
        all_keys = set()
        for values in data.values():
            all_keys.update(values.keys())
        columns = ["日付"] + Utility.sort_by_master(master_list=settings["common"]["results"]+[settings["common"]["completed"]], input_list=all_keys)
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
    
    if structure == 'daily':
        for index, (date, values) in enumerate(data.items()):
            bg_color = alternating_colors[index % 2]
            row = [date] + [values.get(k, 0) for k in Utility.sort_by_master(master_list=settings["common"]["results"]+["completed"], input_list=all_keys)]
            item_id = tree.insert('', 'end', values=row, tags=(date,))
            tree.tag_configure(date, background=highlight_color if date == today else bg_color)
    elif structure == 'by_env':
        if not data:
            tree.insert('', 'end', values=["取得できませんでした", "-"] + ["-" for _ in Utility.sort_by_master(master_list=settings["common"]["results"]+["completed"], input_list=all_keys)])
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
    
    save_button = ttk.Button(frame, text="CSV保存", command=lambda: save_to_csv_tab(tree, columns, file_name, structure))
    save_button.pack(pady=5)
    
    return tree

def save_to_csv_tab(tree, columns, file_name, structure):
    # デフォルトファイル名
    basename = os.path.splitext(file_name)[0]
    tab = settings["app"]["structures"][structure]
    initial_filename = f"{basename}_{tab}"

    # 保存先の選択
    file_path = filedialog.asksaveasfilename(initialfile=initial_filename, defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
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

def save_to_csv_all():
    # データ変換
    converted_data = WriteData.convert_to_2d_array(data=input_data, settings=settings)

    #  デフォルトファイル名
    initial_filename = f'進捗集計_{datetime.today().strftime("%Y-%m-%d")}'

    # 保存先の選択
    file_path = filedialog.asksaveasfilename(initialfile=initial_filename, defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if not file_path:
        return
    
    # 保存
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerows(converted_data)

    # 保存完了
    response = Dialog.ask(title="保存完了", message=f"CSVデータを保存しました。\n{file_path}\n\nファイルを開きますか？")
    if response == "yes":
        open_file(file_path=file_path)

def update_display(selected_file):
    # 全ファイル集計
    update_info_label(sum_values(input_data, "count"), total_count_label, total_rate_label)
    update_bar_chart(data=sum_values(input_data, "total"), incompleted_count=sum_values(input_data, "count")["incompleted"], ax=total_ax, canvas=total_canvas)

    # タブ切り替え
    current_tab = notebook.index(notebook.select()) if notebook.tabs() else 0
    for widget in notebook.winfo_children():
        widget.destroy()
    
    data = next(item for item in input_data if item['selector_label'] == selected_file)

    # ファイル別集計
    update_info_label(data["count"], file_count_label, file_rate_label)
    update_bar_chart(data=data['total'], incompleted_count=data["count"]["incompleted"], ax=file_ax, canvas=file_canvas)

    # タブ内情報の更新
    frame_total = ttk.Frame(notebook)
    notebook.add(frame_total, text=settings["app"]["structures"]["daily"])
    create_treeview(frame_total, data['total_daily'], 'daily', data["file"])

    frame_env = ttk.Frame(notebook)
    notebook.add(frame_env, text=settings["app"]["structures"]["by_env"])
    create_treeview(frame_env, data['by_env'], 'by_env', data["file"])
    
    frame_name = ttk.Frame(notebook)
    notebook.add(frame_name, text=settings["app"]["structures"]["by_name"])
    create_treeview(frame_name, data['by_name'], 'by_name', data["file"])

    # プルダウン切替時にタブの選択状態を保持
    notebook.select(current_tab)

def write_data(field_data):    
    file_path = field_data["filepath"].get()
    table_name = field_data["table_name"].get()

    # ファイルパス未入力
    if not file_path:
        Dialog.show_messagebox(root, type="warning", title="Error", message=f"書込先のファイルを選択してください。")
        return

    # 確認
    response = Dialog.ask(title="保存確認", message=f'{len(input_data)}件のファイルから取得したデータをすべて書き込みます。よろしいですか？')
    if response == "no":
        return

    # フィールドの設定値をグローバルに反映
    settings["app"]["filepath"] = file_path
    settings["app"]["table_name"] = table_name

    # データ書込
    try:
        result = WriteData.execute(input_data, file_path, table_name)
    except ValueError as e:
        Dialog.show_messagebox(root, type="warning", title="データなし", message=e)
        return
    except PermissionError as e:
        Dialog.show_messagebox(root, type="error", title="保存失敗", message=e)
        return

    # 成功したら設定を保存
    if result:
        AppConfig.save_settings(settings)
        response = Dialog.ask(title="保存完了", message=f'"{table_name}"シートにデータを書き込みました。\n{file_path}\n\nこのアプリを終了して、書込先ファイルを開きますか？')
        if response == "yes":
            open_file(file_path=file_path, exit=True)

def select_write_file(entry):
    filepath = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel file", "*.xlsx")])
    if filepath:  # キャンセルで空文字が返ってきたときは変更しない
        entry.delete(0, tk.END)  # 既存の内容をクリア
        entry.insert(0, filepath)  # 新しいファイルパスをセット

def open_file(file_path, exit:bool=False):
    if os.path.isfile(file_path):
        try:
            os.startfile(file_path)  # Windows
        except AttributeError:
            subprocess.run(["xdg-open", file_path])  # Linux/Mac
        # 終了
        if exit:
            sys.exit()
    else:
        Dialog.show_messagebox(root, type="error", title="Error", message="指定されたファイルが見つかりません")

def create_menubar(parent):
    menubar = tk.Menu(parent)
    parent.config(menu=menubar)
    filemenu = tk.Menu(menubar, tearoff=0)
    filemenu.add_command(label="開く", command=reload_files)
    filemenu.add_separator()
    filemenu.add_command(label="終了", command=parent.quit)
    menubar.add_cascade(label="File", menu=filemenu)

def create_input_area(parent, settings):
    input_frame = ttk.LabelFrame(parent, text="集計データ書込")
    input_frame.pack(fill=tk.X, padx=5, pady=5)
    
    ttk.Label(input_frame, text="書込先:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=3)
    file_path_entry = ttk.Entry(input_frame, width=50)
    file_path_entry.insert(0, settings["app"]["filepath"])
    file_path_entry.grid(row=0, column=1, padx=5, pady=5)
    ttk.Button(input_frame, text="選択", command=lambda: select_write_file(file_path_entry)).grid(row=0, column=2, padx=5, pady=2)
    
    field_frame = ttk.Frame(input_frame)
    field_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=2, sticky=tk.W)

    ttk.Label(field_frame, text="データシート名:").pack(side=tk.LEFT, pady=3)
    table_name_entry = ttk.Entry(field_frame, width=20)
    table_name_entry.insert(0, settings["app"]["table_name"])
    table_name_entry.pack(side=tk.LEFT, padx=5)

    field_data = {"filepath": file_path_entry, "table_name": table_name_entry}

    submit_frame = ttk.Frame(input_frame)
    submit_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=2, sticky=tk.W)
    ttk.Button(submit_frame, text="データ書込", command=lambda: write_data(field_data)).pack(side=tk.LEFT, pady=5)
    ttk.Button(submit_frame, text="書込先を開く", command=lambda: open_file(field_data["filepath"].get())).pack(side=tk.LEFT, padx=5, pady=5)
    ttk.Button(submit_frame, text="CSV保存", command=lambda: save_to_csv_all()).pack(side=tk.LEFT, padx=5, pady=5)

def meke_rate_text(value1, value2):
    if value2:
        rate = (value1 / value2) * 100
        # rate = rate if rate < 100 else 100
        return f"{rate:.1f}%"
    else:
        return "--"

def update_info_label(data, count_label, rate_label):
    # 値
    available = data["available"]
    completed = data["completed"]
    filled = data["filled"]

    # ケース数テキスト
    count = available if available else "--"
    count_text = f'テストケース数: {count} (総数: {data["all"]} / 対象外: {data["excluded"]})'
    # 完了率テキスト
    completed_rate_text = f'完了率: {meke_rate_text(completed, available)} [{completed}/{available}]'
    # 消化率テキスト
    filled_rate_text = f'消化率: {meke_rate_text(filled, available)} [{filled}/{available}]'

    # 表示を更新
    count_label.config(text=count_text)
    rate_label.config(text=f'{completed_rate_text}  |  {filled_rate_text}')

def update_bar_chart(data, incompleted_count, ax, canvas):
    # 表示順を固定
    sorted_labels = settings["common"]["results"]

    # 各ラベルとサイズ
    labels = [label for label in sorted_labels if label in data]
    sizes = [data[label] for label in labels]

    # 未実施数を追加
    labels += [settings["common"]["not_run"]]
    sizes += [incompleted_count]

    # 合計値
    total = sum(sizes)

    # 積み上げ横棒グラフ
    ax.clear()
    left = np.zeros(1)  # 初期の左端位置
    bars = []  # 各バーのオブジェクトを保存

    # ラベルごとの色設定
    color_map = settings["common"]["colors"]

    for size, label in zip(sizes, labels):
        color = color_map.get(label, "gray")  # ラベルの色を取得（デフォルトは灰色）
        bar = ax.barh(0, size, left=left, color=color, label=label)  # 横棒グラフを描画
        bars.append((bar[0], size, label, color))  # バー情報を記録
        left += size  # 次のバーの開始位置を更新

    ax.set_xlim(0, total)  # 左右いっぱい
    ax.set_yticks([])  # y軸ラベルを非表示
    ax.set_xticks([])  # x軸の目盛りを非表示
    ax.set_frame_on(False)  # 枠線を削除
    ax.margins(0) # 余白なし

    # バーの中央にラベルを配置（幅が十分にある場合のみ）
    for bar, size, label, color in bars:
        percentage = (size / total) * 100  # 割合計算
        
        # ラベルフォーマット
        if size > total * 0.15:  # 割合15%以上は%も表示
            label_text = f"{label} ({percentage:.1f}%)"
        elif size > total * 0.08:  # 割合8%以上はラベルのみ
            label_text = label
        else:
            label_text = ""

        # ラベルの色
        if color in color_map["black_labels"]:
            label_color = 'black'
        elif color in color_map["gray_labels"]:
            label_color = 'dimgrey'
        else:
            label_color = 'white'

        ax.text(
            bar.get_x() + bar.get_width() / 2,  # 中央位置
            bar.get_y() + bar.get_height() / 2,  # 中央位置
            label_text,
            ha='center', va='center', fontsize=8, 
            color=label_color
        )

    if canvas:
        canvas.draw()

def sum_values(data, param):
    result = defaultdict(int)
    for entry in data:
        for key, value in entry[param].items():
            result[key] += value
    return result

def save_window_position():
    # ウインドウの位置情報を保存
    settings["app"]["window_position"] = root.geometry()
    AppConfig.save_settings(settings)

def on_closing():
    # ウインドウ終了時
    save_window_position()
    root.quit()

def close_all_dialogs():
    """開いている全てのToplevelウィンドウ（ダイアログ含む）を閉じる"""
    for widget in root.winfo_children():
        if isinstance(widget, tk.Toplevel):
            widget.destroy()

def reload_files():
    python = sys.executable
    close_all_dialogs()
    subprocess.Popen([python, *sys.argv])
    sys.exit()

def launch(data, errors):
    global root, notebook, file_selector, input_data, settings
    global file_count_label, file_rate_label, file_fig, file_ax, file_canvas
    global total_count_label, total_rate_label, total_fig, total_ax, total_canvas

    # 読込データ
    input_data = data

    # 設定のロード
    settings = AppConfig.load_settings()

    # 親ウインドウ生成
    root = tk.Tk()
    root.title("TestProgressTracker")

    # ウインドウ表示位置を復元
    root.geometry(settings["app"]["window_position"])

    # データ抽出に失敗したファイルのリスト
    ers = "\n".join(["  "+ f["error"] for f in errors])

    # 1件もデータがなかった場合は終了
    if not len(data):
        Dialog.show_messagebox(root, type="error", title="抽出エラー", message=f"1件もデータが抽出できませんでした。終了します。\n\nFile(s):\n{ers}")
        sys.exit()

    # 再読み込みボタン
    # ttk.Button(root, text="ファイルを再読込", command=reload_files).pack(fill=tk.X)

    # メニューバー
    create_menubar(parent=root)

    # 全ファイルエリア
    total_frame = ttk.LabelFrame(root, text="全ファイル集計")
    total_frame.pack(fill=tk.X, padx=7, pady=5)

    # テストケース数
    total_count_label = ttk.Label(total_frame, anchor="w")
    total_count_label.pack(fill=tk.X, padx=5)

    # 完了率
    total_rate_label = ttk.Label(total_frame, anchor="w")
    total_rate_label.pack(fill=tk.X, padx=5)

    # グラフ表示(全体)
    total_fig, total_ax = plt.subplots(figsize=(8, 0.3))
    total_canvas = FigureCanvasTkAgg(total_fig, master=total_frame)
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    total_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # ファイル別エリア
    file_frame = ttk.LabelFrame(root, text="ファイル別集計")
    file_frame.pack(fill=tk.X, padx=5, pady=5)

    # ファイル選択プルダウン
    file_selector = ttk.Combobox(file_frame, values=[file["selector_label"] for file in input_data], state="readonly")
    file_selector.pack(fill=tk.X, padx=5, pady=5)
    file_selector.bind("<<ComboboxSelected>>", lambda event: update_display(file_selector.get()))

    # テストケース数
    file_count_label = ttk.Label(file_frame, anchor="w")
    file_count_label.pack(fill=tk.X, padx=5)

    # 完了率
    file_rate_label = ttk.Label(file_frame, anchor="w")
    file_rate_label.pack(fill=tk.X, padx=5)

    # グラフ表示(ファイル別)
    file_fig, file_ax = plt.subplots(figsize=(6, 0.2))
    file_canvas = FigureCanvasTkAgg(file_fig, master=file_frame)
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    file_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

    # タブ表示
    notebook = ttk.Notebook(file_frame, height=300)
    notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # ファイル書き込みエリア
    create_input_area(parent=root, settings=settings)

    # グリッド表示
    if input_data:
        file_selector.current(0)
        update_display(input_data[0]['selector_label'])

    if len(errors):
        Dialog.show_messagebox(root, type="warning", title="一部エラー", message=f"以下のファイルはデータが抽出できませんでした。\n\nFile(s):\n{ers}")

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
    root.destroy()
