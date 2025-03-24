import tkinter as tk
from tkinter import ttk, filedialog
import subprocess
import csv, os, sys, pprint
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np

import WriteData
from libs import Utility
from libs import AppConfig
from libs import Dialog

def create_treeview(parent, data, structure, file_name):
    completed = "completed"
    labels = settings["common"]["labels"]
    columns = []
    if structure == 'by_env':
        all_keys = set()
        for dates in data.values():
            for values in dates.values():
                all_keys.update(values.keys())
        columns = ["環境名", "日付"] + Utility.sort_by_master(master_list=settings["common"]["results"]+[labels[completed]], input_list=all_keys)
        data = dict(Utility.sort_nested_dates_desc(data))
    elif structure == 'by_name':
        columns = ["日付", "担当者", "Completed"]
        data = dict(sorted(data.items(), reverse=True))
    elif structure == 'daily':
        all_keys = set()
        for values in data.values():
            all_keys.update(values.keys())
        columns = ["日付"] + Utility.sort_by_master(master_list=settings["common"]["results"]+[labels[completed]], input_list=all_keys)
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
            row = [date] + [values.get(k, 0) for k in Utility.sort_by_master(master_list=settings["common"]["results"]+[labels[completed]], input_list=all_keys)]
            item_id = tree.insert('', 'end', values=row, tags=(date,))
            tree.tag_configure(date, background=highlight_color if date == today else bg_color)
    elif structure == 'by_env':
        if not data:
            tree.insert('', 'end', values=["取得できませんでした", "-"] + ["-" for _ in Utility.sort_by_master(master_list=settings["common"]["results"]+[completed], input_list=all_keys)])
        else:
            for index, (env, dates) in enumerate(data.items()):
                bg_color = alternating_colors[index % 2]
                row_colors[env] = bg_color
                for date, values in dates.items():
                    row = [env, date] + [values.get(k, 0) for k in Utility.sort_by_master(master_list=settings["common"]["results"]+[completed], input_list=all_keys)]
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
        
    menubutton = ttk.Menubutton(parent, text="エクスポート", direction="below")
    menu = tk.Menu(menubutton, tearoff=0)
    menu.add_command(label="CSVで保存", command=lambda: save_to_csv(treeview_to_array(tree), f'{os.path.splitext(file_name)[0]}_{settings["app"]["structures"][structure]}'))
    menu.add_command(label="クリップボードにコピー", command=lambda: copy_to_clipboard(treeview_to_array(tree)))
    menubutton.config(menu=menu)
    menubutton.pack(side=tk.LEFT, padx=2, pady=5)

    return tree

def treeview_to_array(treeview):
    """
    TkinterのTreeviewのデータを2次元配列に変換する関数

    :param treeview: TkinterのTreeviewウィジェット
    :return: 2次元配列（リストのリスト）
    """
    # ヘッダー取得
    columns = treeview["columns"]
    headers = [treeview.heading(col)["text"] for col in columns]

    # データ取得
    data = []
    for item in treeview.get_children():
        row = [treeview.item(item)["values"][i] for i in range(len(columns))]
        data.append(row)

    return [headers] + data  # ヘッダーとデータを結合

def save_to_csv(data, filename):
    # 保存先の選択
    file_path = filedialog.asksaveasfilename(initialfile=filename, defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if not file_path:
        return
    
    # 保存
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerows(data)

    # 保存完了
    response = Dialog.ask(title="保存完了", message=f"CSVデータを保存しました。\n{file_path}\n\nファイルを開きますか？")
    if response == "yes":
        open_file(file_path=file_path)

def update_display(selected_file, count_label, rate_label, ax, canvas, notebook):
    # タブ切り替え
    current_tab = notebook.index(notebook.select()) if notebook.tabs() else 0
    for widget in notebook.winfo_children():
        widget.destroy()
    
    data = next(item for item in input_data if item['selector_label'] == selected_file)

    # ファイル別結果
    if "error" in data:
        count_total_data = {'all': 0, 'excluded': 0, 'available': 0, 'filled': 0, 'completed': 0, 'incompleted': 0}
        total_data = Utility.initialize_dict(settings["common"]["results"])
        incompleted = 0
        daily_data = {}
        by_env_data = {}
        by_name_data = {}
    else:
        count_total_data = data["count_total"]
        total_data = data['total']
        incompleted = data["count_total"]["incompleted"]
        daily_data = data['daily']
        by_env_data = data['by_env']
        by_name_data = data['by_name']

    update_info_label(data=count_total_data, count_label=count_label, rate_label=rate_label, detail=True)
    update_bar_chart(data=total_data, incompleted_count=incompleted, ax=ax, canvas=canvas, show_label=False)

    # TreeViewの更新
    frame_total = ttk.Frame(notebook)
    notebook.add(frame_total, text=settings["app"]["structures"]["daily"])
    create_treeview(frame_total, daily_data, 'daily', data["file"])

    frame_env = ttk.Frame(notebook)
    notebook.add(frame_env, text=settings["app"]["structures"]["by_env"])
    create_treeview(frame_env, by_env_data, 'by_env', data["file"])
    
    frame_name = ttk.Frame(notebook)
    notebook.add(frame_name, text=settings["app"]["structures"]["by_name"])
    create_treeview(frame_name, by_name_data, 'by_name', data["file"])

    # プルダウン切替時にタブの選択状態を保持
    notebook.select(current_tab)

def create_filelist_area(parent):
    # スタイル設定
    padx = 1
    pady = 3

    # フレーム
    file_frame = ttk.Frame(parent)
    file_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    # ヘッダ
    headers = ["ファイル名", "項目数", "完了数", "完了率", "進捗"]
    for col, text in enumerate(headers):
        ttk.Label(file_frame, text=text, foreground="#444444", background="#e0e0e0", relief="solid").grid(
            row=0, column=col, sticky=tk.W+tk.E, padx=padx, pady=pady
        )

    # 列のリサイズ設定
    file_frame.grid_columnconfigure(0, weight=3)

    # データ行
    for index, file_data in enumerate(input_data, 1):
        if "error" in file_data:
            on_error = True
            total_data = {"error": 0}
            completed = 0
            available = 0
            incompleted = 0
            comp_rate_text = "--"
            error_type = file_data["error"]["type"]
            error_message = file_data["error"]["message"]
        else:
            on_error = False
            total_data = file_data['total']
            completed = file_data['count_total']['completed']
            available = file_data['count_total']['available']
            incompleted = file_data['count_total']['incompleted']
            comp_rate_text = Utility.meke_rate_text(completed, available)
        
        # ファイル名
        filename_label = ttk.Label(file_frame, text=file_data['file'])
        filename_label.grid(row=index, column=0, sticky=tk.W, padx=padx, pady=pady)
        if on_error: filename_label.config(foreground="red")
        
        # 項目数
        ttk.Label(file_frame, text=available).grid(
            row=index, column=1, padx=padx, pady=pady
        )

        # 完了数
        ttk.Label(file_frame, text=completed).grid(
            row=index, column=2, padx=padx, pady=pady
        )

        # 完了率
        ttk.Label(file_frame, text=comp_rate_text).grid(
            row=index, column=3, padx=padx, pady=pady
        )

        if on_error:
            # エラー表示
            message = f'{error_message}[{error_type}]'
            ttk.Label(file_frame, text=message, foreground="red").grid(row=index, column=4, sticky=tk.W, padx=padx, pady=pady)
        else:
            # 進捗グラフ
            fig, ax = plt.subplots(figsize=(3, 0.1))
            canvas = FigureCanvasTkAgg(fig, master=file_frame)
            plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
            canvas.get_tk_widget().grid(row=index, column=4, padx=padx, pady=pady)
            # グラフを更新
            update_bar_chart(data=total_data, incompleted_count=incompleted, ax=ax, canvas=canvas, show_label=False)

    # エクスポート
    exp_frame = ttk.Frame(parent)
    exp_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    menubutton = ttk.Menubutton(exp_frame, text="エクスポート", direction="below")
    menu = tk.Menu(menubutton, tearoff=0)
    # menu.add_command(label="CSVで保存", command=lambda: save_to_csv(treeview_to_array(tree), f'{os.path.splitext(file_name)[0]}_{settings["app"]["structures"][structure]}'))
    menu.add_command(label="クリップボードにコピー", command=lambda: copy_to_clipboard(WriteData.convert_to_2d_array(data=input_data, settings=settings)))
    menubutton.config(menu=menu)
    menubutton.pack(anchor=tk.SW, side=tk.BOTTOM, padx=2, pady=5)

def write_to_excel(field_data):    
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
    filepath = filedialog.askopenfilename(title="書込先ファイルを選択", defaultextension=".xlsx", filetypes=[("Excel file", "*.xlsx")])
    if filepath:  # キャンセルで空文字が返ってきたときは変更しない
        entry.delete(0, tk.END)  # 既存の内容をクリア
        entry.insert(0, filepath)  # 新しいファイルパスをセット

def open_file(file_path, exit:bool=False):
    if os.path.isfile(file_path):
        try:
            os.startfile(file_path)  # Windows
        except AttributeError:
            subprocess.run(["xdg-open", file_path])  # Linux/Mac
        # ファイルを開いたあと終了する
        if exit:
            sys.exit()
    else:
        Dialog.show_messagebox(root, type="error", title="Error", message="指定されたファイルが見つかりません")

def copy_to_clipboard(data):
    # 2次元配列をタブ区切りのテキストに変換
    text = "\n".join(["\t".join(map(str, row)) for row in data])
    # クリップボードにコピー
    root.clipboard_clear()
    root.clipboard_append(text)
    root.update()

def edit_settings():
    Dialog.show_messagebox(root=root, type="info", title="ユーザー設定編集", message=f"ユーザー設定ファイルを開きます。\n編集した設定を反映させるには、File > 再読み込み を実行してください。")
    open_file(file_path="UserConfig.json", exit=False)

def create_menubar(parent):
    menubar = tk.Menu(parent)
    parent.config(menu=menubar)
    # File
    file_menu = tk.Menu(menubar, tearoff=0)
    file_menu.add_command(label="開く", command=load_files)
    file_menu.add_command(label="再読み込み", command=reload_files)
    file_menu.add_separator()
    file_menu.add_command(label="終了", command=parent.quit)
    menubar.add_cascade(label="File", menu=file_menu)
    # Settings
    settings_menu = tk.Menu(menubar, tearoff=0)
    settings_menu.add_command(label="設定ファイル編集", command=edit_settings)
    menubar.add_cascade(label="Settings", menu=settings_menu)

def create_input_area(parent, settings):
    input_frame = ttk.LabelFrame(parent, text="集計データ出力")
    input_frame.pack(fill=tk.X, padx=5, pady=3)
    
    ttk.Label(input_frame, text="書込先:").grid(row=0, column=0, sticky=tk.W, padx=2, pady=3)
    file_path_entry = ttk.Entry(input_frame, width=50)
    file_path_entry.insert(0, settings["app"]["filepath"])
    file_path_entry.grid(row=0, column=1)
    ttk.Button(input_frame, text="...", width=3, command=lambda: select_write_file(file_path_entry)).grid(row=0, column=2, padx=2, pady=3)

    ttk.Label(input_frame, text="データシート名:").grid(row=0, column=3, padx=(4,2))
    table_name_entry = ttk.Entry(input_frame, width=20)
    table_name_entry.insert(0, settings["app"]["table_name"])
    table_name_entry.grid(row=0, column=4, padx=2)

    field_data = {"filepath": file_path_entry, "table_name": table_name_entry}

    submit_frame = ttk.Frame(input_frame)
    submit_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=2, sticky=tk.W)
    ttk.Button(submit_frame, text="Excelへ書込", command=lambda: write_to_excel(field_data)).pack(side=tk.LEFT, padx=2, pady=(0,2))
    # ttk.Button(submit_frame, text="書込先を開く", command=lambda: open_file(field_data["filepath"].get())).pack(side=tk.LEFT, padx=2, pady=5)

    ttk.Button(submit_frame, text="CSV保存", command=lambda: save_to_csv(WriteData.convert_to_2d_array(data=input_data, settings=settings), f'進捗集計_{datetime.today().strftime("%Y-%m-%d")}')).pack(side=tk.LEFT, padx=2, pady=(0,2))
    ttk.Button(submit_frame, text="クリップボードにコピー", command=lambda: copy_to_clipboard(WriteData.convert_to_2d_array(data=input_data, settings=settings))).pack(side=tk.LEFT, padx=2, pady=(0,2))

    # menubutton = ttk.Menubutton(submit_frame, text="エクスポート", direction="below")
    # menu = tk.Menu(menubutton, tearoff=0)
    # menu.add_command(label="CSVで保存", command=lambda: save_to_csv(WriteData.convert_to_2d_array(data=input_data, settings=settings), f'進捗集計_{datetime.today().strftime("%Y-%m-%d")}'))
    # menu.add_command(label="クリップボードにコピー", command=lambda: copy_to_clipboard(WriteData.convert_to_2d_array(data=input_data, settings=settings)))
    # menubutton.config(menu=menu)
    # menubutton.pack(side=tk.LEFT, padx=2)

def update_info_label(data, count_label, rate_label, detail=True):
    # 値
    if "error" in data:
        all = None
        available  = None
        completed = None
        filled = None
        excluded = None
    else:
        all = data["all"]
        available = data["available"]
        completed = data["completed"]
        filled = data["filled"]
        excluded = data["excluded"]

    # ケース数テキスト
    count = available if available else "--"
    count_text = f'テストケース数: {count}'
    if detail:
        count_text += f' (総数: {all or "-"}/ 対象外: {excluded or "-"})'
    # 完了率テキスト
    completed_rate_text = f'完了率: {Utility.meke_rate_text(completed, available)} [{completed or "-"}/{available or "-"}]'
    # 消化率テキスト
    filled_rate_text = f'消化率: {Utility.meke_rate_text(filled, available)} [{filled or "-"}/{available or "-"}]'

    # 表示を更新
    count_label.config(text=count_text)
    rate_label.config(text=f'{completed_rate_text}  |  {filled_rate_text}')

def update_bar_chart(data, incompleted_count, ax, canvas, show_label=True):
    # 表示順を固定
    sorted_labels = settings["common"]["results"]

    # 各ラベルとサイズ
    labels = [label for label in sorted_labels if label in data]
    sizes = [data[label] for label in labels]

    # 未実施数を追加
    labels += [settings["common"]["labels"]["not_run"]]
    sizes += [incompleted_count]

    # 合計値
    total = sum(sizes)

    # 積み上げ横棒グラフ
    ax.clear()
    left = np.zeros(1)  # 初期の左端位置
    bars = []  # 各バーのオブジェクトを保存

    # ラベルごとの色設定
    bar_color_map = settings["app"]["colors"]["bar"]
    label_color_map = settings["app"]["colors"]["label"]

    for size, label in zip(sizes, labels):
        color = bar_color_map.get(label, "gray")  # ラベルの色を取得（デフォルトは灰色）
        bar = ax.barh(0, size, left=left, color=color, label=label)  # 横棒グラフを描画
        bars.append((bar[0], size, label, color))  # バー情報を記録
        left += size  # 次のバーの開始位置を更新

    if total:
        ax.set_xlim(0, total)  # 左右いっぱい
    ax.set_yticks([])  # y軸ラベルを非表示
    ax.set_xticks([])  # x軸の目盛りを非表示
    ax.set_frame_on(False)  # 枠線を削除
    ax.margins(0) # 余白なし

    # バーの中央にラベルを配置
    if show_label:
        for bar, size, label, color in bars:
            # 割合計算
            percentage = Utility.safe_divide(size, total)
            if percentage:
                # 表示用の値
                percentage = (percentage) * 100
                # ラベルフォーマット
                if size > total * 0.15:  # 割合15%以上は%も表示
                    label_text = f"{label} ({percentage:.1f}%)"
                elif size > total * 0.08:  # 割合8%以上はラベルのみ
                    label_text = label
                else:
                    label_text = ""
            else:
                label_text = ""

            # ラベルの色
            if color in label_color_map["black"]:
                label_color = 'black'
            elif color in label_color_map["gray"]:
                label_color = 'dimgrey'
            elif total == 0:
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

def new_process(inputs):
    close_all_dialogs()
    python = sys.executable
    subprocess.Popen([python, sys.argv[0]] + inputs)
    sys.exit()

def load_files():
    inputs = Dialog.select_files(("Excel/Zipファイル", "*.xlsx;*.zip"))
    if inputs: new_process(inputs=list(inputs))

def reload_files():
    response = Dialog.ask(title="確認", message="全てのファイルを再読み込みします。よろしいですか？")
    if response == "yes": new_process(inputs=list(input_args))

def create_global_tab(parent):
    # 全体タブ
    nb = ttk.Notebook(parent)
    tab1 = tk.Frame(nb)
    tab2 = tk.Frame(nb)
    nb.add(tab1, text=' 集計結果 ')
    nb.add(tab2, text=' ファイル別 ')
    nb.pack(fill=tk.BOTH, expand=True)
    return tab1, tab2

def create_total_tab(parent):
    # 全体集計結果にはエラーを含めない
    filtered_data = Utility.filter_objects(input_data, exclude_key="error")

    # 集計結果タブ
    total_frame = ttk.Frame(parent)
    total_frame.pack(fill=tk.X, padx=5, pady=5)

    # テストケース数、完了率(全体)
    total_count_label = ttk.Label(total_frame, anchor="w")
    total_count_label.pack(fill=tk.X, padx=5)
    total_rate_label = ttk.Label(total_frame, anchor="w")
    total_rate_label.pack(fill=tk.X, padx=5)

    # テストケース数、完了率を更新
    update_info_label(data=Utility.sum_values(filtered_data, "count_total"), count_label=total_count_label, rate_label=total_rate_label, detail=True)

    # グラフ表示(全体)
    total_fig, total_ax = plt.subplots(figsize=(8, 0.25))
    total_canvas = FigureCanvasTkAgg(total_fig, master=total_frame)
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    total_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # グラフを更新
    update_bar_chart(data=Utility.sum_values(filtered_data, "total"), incompleted_count=Utility.sum_values(filtered_data, "count_total")["incompleted"], ax=total_ax, canvas=total_canvas, show_label=True)

    # ファイル別グラフ
    create_filelist_area(parent=parent)

    # ファイル書き込みエリア
    create_input_area(parent=parent, settings=settings)

def create_byfile_tab(parent):
    file_frame = ttk.Frame(parent)
    file_frame.pack(fill=tk.X, padx=5, pady=5)

    # ファイル選択プルダウン
    file_selector = ttk.Combobox(file_frame, values=[file["selector_label"] for file in input_data], state="readonly")
    file_selector.pack(fill=tk.X, padx=20, pady=5)
    file_selector.bind("<<ComboboxSelected>>", lambda event: update_display(file_selector.get(), count_label=file_count_label, rate_label=file_rate_label, ax=file_ax, canvas=file_canvas, notebook=notebook))

    # テストケース数(ファイル別)
    file_count_label = ttk.Label(file_frame, anchor="w")
    file_count_label.pack(fill=tk.X, padx=20)

    # 完了率(ファイル別)
    file_rate_label = ttk.Label(file_frame, anchor="w")
    file_rate_label.pack(fill=tk.X, padx=20)

    # グラフ表示(ファイル別)
    file_fig, file_ax = plt.subplots(figsize=(6, 0.1))
    file_canvas = FigureCanvasTkAgg(file_fig, master=file_frame)
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    file_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

    # タブ表示
    notebook_height = 300 if len(input_data) > 1 else 355
    notebook = ttk.Notebook(file_frame, height=notebook_height)
    notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # グリッド表示
    if input_data:
        file_selector.current(0)
        update_display(input_data[0]['selector_label'], count_label=file_count_label, rate_label=file_rate_label, ax=file_ax, canvas=file_canvas, notebook=notebook)

def launch(data, args):
    global root, input_data, settings, input_args
    
    # 読込データ
    input_data = data

    # # 正常データ
    # input_data = [r for r in data]
    # # エラーデータ
    # errors = [r for r in data if "error" in r]

    # 起動時に指定したファイルパス（再読込用）
    input_args = args
    # 設定のロード
    settings = AppConfig.load_settings()

    # 親ウインドウ生成
    root = tk.Tk()
    root.title("TestTraQ")
    # ウインドウ表示位置を復元
    root.geometry(settings["app"]["window_position"])

    # メニューバー
    create_menubar(parent=root)
    # グローバルタブ
    tab1, tab2 = create_global_tab(parent=root)

    # タブ1：全体集計タブ
    create_total_tab(tab1)
    # タブ2：ファイル別集計タブ
    create_byfile_tab(tab2)

    # データ抽出に失敗したファイルのリスト
    errors = [r for r in data if "error" in r]
    ers = "\n".join(["  "+ err["file"] for err in errors])

    # 1件もデータがなかった場合は終了
    if not len(data):
        Dialog.show_messagebox(root, type="error", title="抽出エラー", message=f"1件もデータが抽出できませんでした。終了します。\n\nFile(s):\n{ers}")
        sys.exit()

    # 一部ファイルにデータがなかった場合はアラート
    if len(errors):
        Dialog.show_messagebox(root, type="warning", title="一部エラー", message=f"以下のファイルはデータが抽出できませんでした。\n\nFile(s):\n{ers}")

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
    root.destroy()
