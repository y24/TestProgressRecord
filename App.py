import tkinter as tk
from tkinter import ttk, filedialog
from tktooltip import ToolTip
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

def _create_row_data(structure: str, values: dict, settings: dict, all_keys: set) -> list:
    """行データを生成する

    Args:
        structure: データ構造タイプ ('daily', 'by_env', 'by_name')
        values: 表示データ
        settings: 設定情報
        all_keys: 全キーのセット

    Returns:
        list: 表示用の行データ
    """
    completed = settings["test_status"]["labels"]["completed"]
    
    if structure == 'daily':
        return [values[0]] + [values[1].get(k, 0) for k in Utility.sort_by_master(
            master_list=settings["test_status"]["results"] + [completed],
            input_list=all_keys
        )]
    elif structure == 'by_env':
        env, date, data = values
        return [env, date] + [data.get(k, 0) for k in Utility.sort_by_master(
            master_list=settings["test_status"]["results"] + [completed],
            input_list=all_keys
        )]
    elif structure == 'by_name':
        date, name, count = values
        return (date, name, count)

def _insert_tree_rows(tree: ttk.Treeview, structure: str, data: dict, settings: dict, 
                     all_keys: set) -> None:
    """Treeviewに行データを挿入する

    Args:
        tree: 対象のTreeviewウィジェット
        structure: データ構造タイプ
        data: 表示データ
        settings: 設定情報
        all_keys: 全キーのセット
    """
    today = datetime.today().strftime("%Y-%m-%d")
    alternating_colors = ["#ffffff", "#f0f0f0"]
    highlight_color = "#d0f0ff"
    
    # データが空の場合（by_envのみ）
    if structure == 'by_env' and not data:
        dummy_row = ["環境名を取得できませんでした", "-"] + ["-" for _ in Utility.sort_by_master(
            master_list=settings["test_status"]["results"] + [settings["test_status"]["labels"]["completed"]],
            input_list=all_keys
        )]
        tree.insert('', 'end', values=dummy_row)
        return

    # データの挿入
    for index, item in enumerate(data.items()):
        bg_color = alternating_colors[index % 2]
        
        if structure == 'daily':
            date, values = item
            row_data = _create_row_data(structure, (date, values), settings, all_keys)
            tree.insert('', 'end', values=row_data, tags=(date,))
            tree.tag_configure(date, background=highlight_color if date == today else bg_color)
            
        elif structure == 'by_env':
            env, dates = item
            for date, values in dates.items():
                row_data = _create_row_data(structure, (env, date, values), settings, all_keys)
                tree.insert('', 'end', values=row_data, tags=(date,))
                tree.tag_configure(date, background=highlight_color if date == today else bg_color)
                
        elif structure == 'by_name':
            date, names = item
            for name, count in names.items():
                row_data = _create_row_data(structure, (date, name, count), settings, all_keys)
                tree.insert('', 'end', values=row_data, tags=(date,))
                tree.tag_configure(date, background=highlight_color if date == today else bg_color)

def _get_all_keys(data: dict, structure: str) -> set:
    """データ構造に応じて全てのキーを収集する

    Args:
        data: 表示データ
        structure: データ構造タイプ ('daily', 'by_env', 'by_name')

    Returns:
        set: 収集されたキーの集合
    """
    all_keys = set()
    
    if structure == 'by_env':
        for dates in data.values():
            for values in dates.values():
                all_keys.update(values.keys())
    elif structure == 'daily':
        for values in data.values():
            all_keys.update(values.keys())
    # by_nameの場合は空セットのまま
    
    return all_keys

def _get_columns(structure: str, settings: dict, all_keys: set) -> list:
    """データ構造に応じて列定義を生成する

    Args:
        structure: データ構造タイプ
        settings: 設定情報
        all_keys: 全キーの集合

    Returns:
        list: 列名のリスト
    """
    completed = settings["test_status"]["labels"]["completed"]
    
    column_definitions = {
        'by_env': ["環境名", "日付"],
        'by_name': ["日付", "担当者", "Completed"],
        'daily': ["日付"]
    }
    
    base_columns = column_definitions.get(structure, [])
    
    # by_nameの場合は基本列のみ返す
    if structure == 'by_name':
        return base_columns
        
    # その他の場合は結果列を追加
    result_columns = Utility.sort_by_master(
        master_list=settings["test_status"]["results"] + [completed],
        input_list=all_keys
    )
    
    return base_columns + result_columns

def _sort_data(data: dict, structure: str) -> dict:
    """データ構造に応じてデータをソートする

    Args:
        data: ソート対象のデータ
        structure: データ構造タイプ

    Returns:
        dict: ソート済みデータ
    """
    if structure == 'by_env':
        return dict(Utility.sort_nested_dates_desc(data))
    else:  # daily, by_name
        return dict(sorted(data.items(), reverse=True))

def create_treeview(parent, data, structure, file_name):
    """Treeviewウィジェットを作成する"""

    # 全キーの収集
    all_keys = _get_all_keys(data, structure)
    
    # 列定義の生成
    columns = _get_columns(structure, settings, all_keys)
    
    # データのソート
    data = _sort_data(data, structure)
    
    # Treeviewの作成（以下は既存のコード）
    frame = ttk.Frame(parent)
    frame.pack(fill=tk.BOTH, expand=True)
    
    tree_frame = ttk.Frame(frame)
    tree_frame.pack(fill=tk.BOTH, expand=True)

    # スクロール可能なTreeviewの作成
    tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
    scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar_y.set)
    
    scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
    tree.pack(fill=tk.BOTH, expand=True)

    # 列の設定（ヘッダーテキストと幅）
    for col in columns:
        tree.heading(col, text=col)
        # 列の種類に応じて幅を設定
        if col == "日付":
            col_w = 100
        elif col == "環境名":
            col_w = 120
        else:
            col_w = 70
        tree.column(col, anchor='center', width=col_w)
    
    # データ構造タイプに応じてTreeviewにデータを挿入
    _insert_tree_rows(tree, structure, data, settings, all_keys)
    
    tree.pack(fill=tk.BOTH, expand=True)
    
    # エクスポートメニューの作成
    menubutton = ttk.Menubutton(parent, text="エクスポート", direction="below")
    menu = tk.Menu(menubutton, tearoff=0)
    menu.add_command(label="CSVで保存", 
                    command=lambda: save_to_csv(
                        treeview_to_array(tree), 
                        f'{os.path.splitext(file_name)[0]}_{settings["app"]["structures"][structure]}'
                    ))
    menu.add_command(label="クリップボードにコピー", 
                    command=lambda: copy_to_clipboard(treeview_to_array(tree)))
    menubutton.config(menu=menu)
    menubutton.pack(side=tk.LEFT, padx=2, pady=5)

    return tree

def treeview_to_array(treeview):
    """
    TkinterのTreeviewのデータを2次元配列に変換する関数

    :param treeview: TkinterのTreeviewウィジェット
    :return: 2次元配列（リストのリスト）
    """
    # Treeviewのヘッダー行を取得
    columns = treeview["columns"]
    headers = [treeview.heading(col)["text"] for col in columns]

    # Treeviewのデータを取得
    data = []
    for item in treeview.get_children():
        row = [treeview.item(item)["values"][i] for i in range(len(columns))]
        data.append(row)

    return [headers] + data  # ヘッダーとデータを結合

# 2次元配列をCSVファイルに保存
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

def update_byfile_tab(selected_file, count_label, rate_label, ax, canvas, notebook):
    # タブ切り替え
    current_tab = notebook.index(notebook.select()) if notebook.tabs() else 0
    for widget in notebook.winfo_children():
        widget.destroy()
    
    data = next(item for item in input_data if item['selector_label'] == selected_file)

    # ファイル別結果
    if "error" in data:
        stats_data = {'all': 0, 'excluded': 0, 'available': 0, 'filled': 0, 'completed': 0, 'incompleted': 0}
        total_data = Utility.initialize_dict(settings["test_status"]["results"])
        incompleted = 0
        daily_data = {}
        by_env_data = {}
        by_name_data = {}
    else:
        stats_data = data["stats"]
        total_data = data['total']
        incompleted = data["stats"]["incompleted"]
        daily_data = data['daily']
        by_env_data = data['by_env']
        by_name_data = data['by_name']

    # 集計情報の更新
    update_info_label(data=stats_data, count_label=count_label, rate_label=rate_label, detail=True)
    update_bar_chart(data=total_data, incompleted_count=incompleted, ax=ax, canvas=canvas, show_label=False)

    # ツールチップを更新
    if not hasattr(canvas.get_tk_widget(), '_tooltip'):
        # 初回のみツールチップを作成
        canvas.get_tk_widget()._tooltip = ToolTip(canvas.get_tk_widget(), msg="", delay=0.3, follow=False)
    
    # 既存のツールチップのメッセージを更新
    graph_tooltip = f'{make_results_text(total_data, incompleted)}'
    canvas.get_tk_widget()._tooltip.msg = graph_tooltip

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

def create_click_handler(filepath):
    return lambda event: open_file(file_path=filepath)

def make_results_text(results, incompleted):
    items = [f'{key}:{value}' for key, value in results.items() if value > 0]
    if incompleted: items.append(f'Not Run:{incompleted}')
    if len(items):
        return ', '.join(items)
    else:
        return "有効なデータがありません。エラーのないファイルのみが集計されます。"

def set_state_color(label, state_name):
    # キー名を取得
    state_key = Utility.find_key_by_name(settings["app"]["state"], state_name)
    # 見つからない場合は何もしない
    if state_key is None: return
    # 色を設定
    state_info = settings["app"]["state"][state_key]
    label.config(foreground=state_info["foreground"], background=state_info["background"])

def _extract_file_data(file_data: dict) -> dict:
    """ファイルデータから表示用の情報を抽出する"""
    base_info = {
        "on_warning": False,
        "on_error": False,
        "error_type": "",
        "error_message": "",
    }

    # ワーニングの確認
    if "warning" in file_data:
        base_info.update({
            "on_warning": True,
            "error_type": file_data["warning"]["type"],
            "error_message": file_data["warning"]["message"]
        })

    # エラー時はダミーデータを返却
    if "error" in file_data:
        return {
            **base_info,
            "on_error": True,
            "total_data": {
                "error": 0,
                "all": 0,        # エラー時のall追加
                "excluded": 0    # エラー時のexcluded追加
            },
            "state": "???",
            "completed": "",
            "available": "",
            "incompleted": 0,
            "comp_rate_text": "",
            "start_date": "",
            "last_update": "",
            "error_type": file_data["error"]["type"],
            "error_message": file_data["error"]["message"]
        }

    # 正常時のデータ抽出
    stats = file_data["stats"]
    run_data = file_data["run"]
    
    return {
        **base_info,
        "total_data": {
            **file_data["total"]
        },
        "state": run_data["status"],
        "all": stats["all"],
        "excluded": stats["excluded"],
        "completed": stats["completed"],
        "available": stats["available"],
        "incompleted": stats["incompleted"],
        "comp_rate_text": Utility.meke_rate_text(stats["completed"], stats["available"]),
        "start_date": run_data["start_date"],
        "last_update": run_data["last_update"]
    }

def update_filelist_table(table_frame):
    # テーブルの余白
    padx = 1
    pady = 3

    # ヘッダ
    headers = ["No.", "ファイル名", "State", "開始日", "更新日", "完了率", "テスト結果"]
    for col, text in enumerate(headers):
        ttk.Label(table_frame, text=text, foreground="#444444", background="#e0e0e0", relief="solid").grid(
            row=0, column=col, sticky=tk.W+tk.E, padx=padx, pady=pady
        )

    # 列のリサイズ設定
    table_frame.grid_columnconfigure(1, weight=3)

    # クリップボード出力用のヘッダ
    export_headers = ["No.", "ファイル名", "State", "開始日", "更新日", "項目数", "完了数", "完了率"]
    export_data = [export_headers + settings["test_status"]["results"] + [settings["test_status"]["labels"]["not_run"]]]

    # 各ファイルのデータ表示
    for index, file_data in enumerate(input_data, 1):
        display_data = _extract_file_data(file_data)
        # インデックス
        ttk.Label(table_frame, text=index).grid(row=index, column=0, padx=padx, pady=pady)
        export_row = [index] # エクスポート用データ
        
        # ファイル名
        filename = file_data['file']
        filename_label = ttk.Label(table_frame, text=filename)
        filename_label.grid(row=index, column=1, sticky=tk.W, padx=padx, pady=pady)
        tooltip_text = [filename]
        export_row.append(filename)

        # ファイル名ダブルクリック時
        filepath = file_data['filepath']
        filename_label.bind("<Double-Button-1>", create_click_handler(filepath))

        # State
        state_label = ttk.Label(table_frame, text=display_data["state"], anchor="center")
        state_label.grid(row=index, column=2, padx=padx, pady=pady, sticky=tk.W + tk.E)
        # set_state_color(state_label, display_data["state"])
        export_row.append(display_data["state"])

        # 開始日
        start_label = ttk.Label(table_frame, text=Utility.simplify_date(display_data["start_date"]))
        start_label.grid(row=index, column=3, padx=padx, pady=pady)
        export_row.append(display_data["start_date"] or "")

        # 最終更新日
        last_update_label = ttk.Label(table_frame, text=Utility.simplify_date(display_data["last_update"]))
        last_update_label.grid(row=index, column=4, padx=padx, pady=pady)
        export_row.append(display_data["last_update"] or "")

        # 完了率
        if display_data["on_error"]:
            comp_rate_export = ""
            comp_rate_display = "-"
        else:
            comp_rate_export = display_data["comp_rate_text"]
            comp_rate_display = f'{display_data["comp_rate_text"]} ({display_data["completed"]}/{display_data["available"]})'

        comp_rate_label = ttk.Label(table_frame, text=comp_rate_display)
        comp_rate_label.grid(row=index, column=5, padx=padx, pady=pady)
        # エクスポート用データ
        export_row.append(display_data["available"]) # 項目数
        export_row.append(display_data["completed"]) # 完了数
        export_row.append(comp_rate_export) # 完了率

        # エラー時赤色・ワーニング時オレンジ色
        if display_data["on_error"] or display_data["on_warning"]:
            color = "red" if display_data["on_error"] else "darkorange2"
            filename_label.config(foreground=color)
            state_label.config(foreground=color)
            start_label.config(foreground=color)
            comp_rate_label.config(foreground=color)

        # エラー時以外は進捗グラフを表示
        if not display_data["on_error"]:
            # 進捗グラフ
            fig, ax = plt.subplots(figsize=(2, 0.1))
            canvas = FigureCanvasTkAgg(fig, master=table_frame)
            plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
            canvas.get_tk_widget().grid(row=index, column=6, padx=padx, pady=pady)
            # グラフを更新
            update_bar_chart(data=display_data["total_data"], incompleted_count=display_data["incompleted"], ax=ax, canvas=canvas, show_label=False)
            # グラフのツールチップ
            graph_tooltop = f"項目数: {display_data['available']} (Total: {display_data['all']} / 対象外: {display_data['excluded']})\nState: {display_data['state']}\n{make_results_text(display_data['total_data'], display_data['incompleted'])}"
            ToolTip(canvas.get_tk_widget(), msg=graph_tooltop, delay=0.3, follow=False)
            # エクスポート用データ
            export_row += list(display_data["total_data"].values())
            export_row.append(display_data["incompleted"])

        # エクスポート用の配列に格納
        export_data.append(export_row)

        # エラー・ワーニング時はツールチップにメッセージを追加
        if display_data["on_error"] or display_data["on_warning"]:
            tooltip_text.append(f'{display_data["error_message"]}[{display_data["error_type"]}]')

        # ファイル名のツールチップ
        tooltip_text.append("<ダブルクリックで開きます>")
        ToolTip(filename_label, msg="\n".join(tooltip_text), delay=0.3, follow=False)

    return export_data

def sort_input_data(order: str, type: str = "asc") -> None:
    """入力データを指定された順序でソートする

    Args:
        order (str): ソート基準
            - "start_date": 開始日時でソート
            - "last_update": 最終更新日時でソート
            - "file_name": ファイル名でソート
            - "completed_rate": 完了率でソート
        type (str): ソート順
            - "asc": 昇順（デフォルト）
            - "desc": 降順
    """
    # ソートキーの定義
    def safe_get(x, keys, default=None):
        """ネストされたディクショナリから安全に値を取得する"""
        try:
            value = x
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            return default

    sort_keys = {
        "start_date": lambda x: (
            0 if "run" not in x else 1,
            0 if safe_get(x, ["run", "start_date"]) is None else 1,
            safe_get(x, ["run", "start_date"], "")
        ),
        "last_update": lambda x: (
            0 if "run" not in x else 1,
            0 if safe_get(x, ["run", "last_update"]) is None else 1,
            safe_get(x, ["run", "last_update"], "")
        ),
        "file_name": lambda x: (
            0 if "file" not in x else 1,
            0 if safe_get(x, ["file"]) is None else 1,
            safe_get(x, ["file"], "").lower()
        ),
        "completed_rate": lambda x: (
            0 if "stats" not in x else 1,
            0 if safe_get(x, ["stats", "available"]) is None else 1,
            safe_get(x, ["stats", "completed"], 0) / safe_get(x, ["stats", "available"], 1) 
            if safe_get(x, ["stats", "available"], 0) > 0 else 0
        )
    }

    # 存在するソートキーかチェック
    if order not in sort_keys:
        print(f"Warning: Unknown sort order '{order}'. Using 'start_date' as default.")
        order = "start_date"

    # ソート順の検証
    if type not in ["asc", "desc"]:
        print(f"Warning: Invalid sort type '{type}'. Using 'asc' as default.")
        type = "asc"

    # ソート実行
    global input_data
    input_data = sorted(
        input_data, 
        key=sort_keys[order], 
        reverse=(type == "desc")  # type が "desc" の場合は reverse=True
    )

def clear_frame(frame):
    for widget in frame.winfo_children():
        widget.destroy()

def change_sort_order(table_frame, order, sort_menu_button, on_change=False):
    global export_data
    sort_input_data(order, type=settings["app"]["sort"]["orders"][order]["type"])
    clear_frame(table_frame) # 表示をクリア
    if on_change:
        plt.close('all') # 表示していたグラフを開放する
    export_data = update_filelist_table(table_frame)
    sort_menu_button.config(text=f'ソート: {settings["app"]["sort"]["orders"][order]["label"]}')
    if on_change:
        # デフォルト設定に保存
        settings["app"]["sort"]["default"] = order
        AppConfig.save_settings(settings)

def create_summary_filelist_area(parent):
    global export_data

    # ファイル別テーブル
    table_frame = ttk.Frame(parent)
    table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    # ファイル別テーブルの更新
    export_data = update_filelist_table(table_frame)

    # メニュー
    menu_frame = ttk.Frame(parent)
    menu_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    # 並び替えメニュー
    sort_menu_button = ttk.Menubutton(menu_frame, text="ソート", direction="below")
    sort_menu = tk.Menu(sort_menu_button, tearoff=0)
    for key, order_info in settings["app"]["sort"]["orders"].items():
        sort_menu.add_command(label=order_info["label"], command=lambda key=key: change_sort_order(table_frame, key, sort_menu_button, on_change=True))
    sort_menu_button.config(menu=sort_menu)
    sort_menu_button.pack(anchor=tk.SW, side=tk.LEFT, padx=2, pady=5)
    change_sort_order(table_frame, settings["app"]["sort"]["default"], sort_menu_button, on_change=False)

    # エクスポートメニュー
    exp_menu_button = ttk.Menubutton(menu_frame, text="エクスポート", direction="below")
    expmenu = tk.Menu(exp_menu_button, tearoff=0)
    expmenu.add_command(label="CSVで保存", command=lambda: save_to_csv(export_data, f'進捗集計_{Utility.get_today_str()}'))
    expmenu.add_command(label="クリップボードにコピー", command=lambda: copy_to_clipboard(export_data))
    exp_menu_button.config(menu=expmenu)
    exp_menu_button.pack(anchor=tk.SW, side=tk.LEFT, padx=2, pady=5)

def write_to_excel(file_path, table_name):
    # ファイルパス未入力
    if not file_path:
        Dialog.show_messagebox(root, type="warning", title="Error", message=f"書込先のファイルを選択してください。")
        return

    # 確認
    response = Dialog.ask(title="保存確認", message=f'{len(input_data)}件のファイルから取得したデータをすべて書き込みます。よろしいですか？')
    if response == "no":
        return

    # フィールドの設定値をグローバルに反映
    settings["app"]["write"]["filepath"] = file_path
    settings["app"]["write"]["table_name"] = table_name

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
        except Exception as e:
            Dialog.show_messagebox(root, type="error", title="Error", message=f"ファイルを開けませんでした。\n{e}")
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
    file_path_entry = ttk.Entry(input_frame, width=80)
    file_path_entry.insert(0, settings["app"]["write"]["filepath"])
    file_path_entry.grid(row=0, column=1)
    ttk.Button(input_frame, text="...", width=3, command=lambda: select_write_file(file_path_entry)).grid(row=0, column=2, padx=2, pady=3)

    ttk.Label(input_frame, text="データシート名:").grid(row=0, column=3, padx=(4,2))
    table_name_entry = ttk.Entry(input_frame, width=20)
    table_name_entry.insert(0, settings["app"]["write"]["table_name"])
    table_name_entry.grid(row=0, column=4, padx=2)

    submit_frame = ttk.Frame(input_frame)
    submit_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=2, sticky=tk.W)
    ttk.Button(submit_frame, text="Excelへ書込", command=lambda: write_to_excel(file_path_entry.get(), table_name_entry.get())).pack(side=tk.LEFT, padx=2, pady=(0,2))
    # ttk.Button(submit_frame, text="書込先を開く", command=lambda: open_file(field_data["filepath"].get())).pack(side=tk.LEFT, padx=2, pady=5)

    ttk.Button(submit_frame, text="CSV保存", command=lambda: save_to_csv(WriteData.convert_to_2d_array(data=input_data, settings=settings), f'進捗集計_{Utility.get_today_str()}')).pack(side=tk.LEFT, padx=2, pady=(0,2))
    ttk.Button(submit_frame, text="クリップボードにコピー", command=lambda: copy_to_clipboard(WriteData.convert_to_2d_array(data=input_data, settings=settings))).pack(side=tk.LEFT, padx=2, pady=(0,2))

def update_info_label(data, count_label, rate_label, detail=True):
    if len(data) == 0 or"error" in data:
        # データなしまたはエラー時
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
    count_text = f'項目数: {count}'
    if detail:
        count_text += f' (Total: {all or "-"} / 対象外: {excluded or "-"})'
    # 完了率テキスト
    completed_rate_text = f'完了率: {Utility.meke_rate_text(completed, available)} [{completed or "-"}/{available or "-"}]'
    # 消化率テキスト
    filled_rate_text = f'消化率: {Utility.meke_rate_text(filled, available)} [{filled or "-"}/{available or "-"}]'

    # 表示を更新
    count_label.config(text=count_text)
    rate_label.config(text=f'{completed_rate_text}  |  {filled_rate_text}')

def update_bar_chart(data, incompleted_count, ax, canvas, show_label=True):
    # 表示順を固定
    sorted_labels = settings["test_status"]["results"]

    # 各結果(ラベル)とカウント(サイズ)
    labels = [label for label in sorted_labels if label in data]
    sizes = [data[label] for label in labels]

    # 未実施数を追加
    if incompleted_count > 0:
        labels += [settings["test_status"]["labels"]["not_run"]]
        sizes += [incompleted_count]

    # 有効データなし時
    if len(labels) == 0:
        labels = ["No Data"]
        sizes = [1]

    # 合計値
    total = sum(sizes)

    # 積み上げ横棒グラフ
    ax.clear()
    left = np.zeros(1)  # 初期の左端位置
    bars = []  # 各バーのオブジェクトを保存

    # ラベルごとの色設定
    bar_color_map = settings["app"]["bar"]["colors"]
    label_color_map = settings["app"]["bar"]["font_color_map"]

    for size, label in zip(sizes, labels):
        color = bar_color_map.get(label, "gainsboro")  # ラベルの色を取得（デフォルトは薄い灰色）
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

            # データなし時は%表示なし
            if label == "No Data": label_text = label

            # ラベルの色
            if color in label_color_map["black"]:
                label_color = 'black'
            elif color in label_color_map["gray"]:
                label_color = 'dimgrey'
            elif color == "gainsboro":
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

def create_summary_tab(parent):
    # 全体集計にはエラーとワーニングのあるデータを含めない
    filtered_data = Utility.filter_objects(input_data, exclude_keys=["error", "warning"])

    # 集計結果タブ
    total_frame = ttk.Frame(parent)
    total_frame.pack(fill=tk.X, padx=5, pady=5)

    # テストケース数、完了率(全体)
    total_count_label = ttk.Label(total_frame, anchor="w")
    total_count_label.pack(fill=tk.X, padx=5)
    total_rate_label = ttk.Label(total_frame, anchor="w")
    total_rate_label.pack(fill=tk.X, padx=5)

    # テストケース数、完了率を更新
    update_info_label(data=Utility.sum_values(filtered_data, "stats"), count_label=total_count_label, rate_label=total_rate_label, detail=True)

    # グラフ表示(全体)
    total_fig, total_ax = plt.subplots(figsize=(8, 0.25))
    total_canvas = FigureCanvasTkAgg(total_fig, master=total_frame)
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    total_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # グラフを更新
    incompleted = Utility.sum_values(filtered_data, "stats")["incompleted"] if len(filtered_data) > 0 else 0
    update_bar_chart(data=Utility.sum_values(filtered_data, "total"), incompleted_count=incompleted, ax=total_ax, canvas=total_canvas, show_label=True)

    # グラフのツールチップ(全体)
    filtered_total_data = Utility.sum_values(filtered_data, "total")
    graph_tooltip = f'{make_results_text(filtered_total_data, incompleted)}'
    ToolTip(total_canvas.get_tk_widget(), msg=graph_tooltip, delay=0.3, follow=False)

    # ファイル別データ表示部
    create_summary_filelist_area(parent=parent)

    # ファイル書き込みエリア
    create_input_area(parent=parent, settings=settings)

def create_byfile_tab(parent):
    by_file_frame = ttk.Frame(parent)
    by_file_frame.pack(fill=tk.X, padx=5, pady=5)

    # ファイル選択プルダウン
    file_selector = ttk.Combobox(by_file_frame, values=[file["selector_label"] for file in input_data], state="readonly")
    file_selector.pack(fill=tk.X, padx=20, pady=5)
    file_selector.bind("<<ComboboxSelected>>", lambda event: update_byfile_tab(selected_file=file_selector.get(), count_label=file_count_label, rate_label=file_rate_label, ax=file_ax, canvas=file_canvas, notebook=notebook))

    # テストケース数(ファイル別)
    file_count_label = ttk.Label(by_file_frame, anchor="w")
    file_count_label.pack(fill=tk.X, padx=20)

    # 完了率(ファイル別)
    file_rate_label = ttk.Label(by_file_frame, anchor="w")
    file_rate_label.pack(fill=tk.X, padx=20)

    # グラフ表示(ファイル別)
    file_fig, file_ax = plt.subplots(figsize=(6, 0.1))
    file_canvas = FigureCanvasTkAgg(file_fig, master=by_file_frame)
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    file_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

    # タブ表示
    notebook_height = 300 if len(input_data) > 1 else 355
    notebook = ttk.Notebook(by_file_frame, height=notebook_height)
    notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # グリッド表示
    if input_data:
        file_selector.current(0)
        update_byfile_tab(selected_file=input_data[0]['selector_label'], count_label=file_count_label, rate_label=file_rate_label, ax=file_ax, canvas=file_canvas, notebook=notebook)

def launch(data, args):
    global root, input_data, settings, input_args
    
    # 読込データ
    input_data = data

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
    create_summary_tab(tab1)
    # タブ2：ファイル別集計タブ
    create_byfile_tab(tab2)

    # データ抽出に失敗したファイルのリスト
    errors = [r for r in data if "error" in r]
    ers = "\n".join(["  "+ err["file"] for err in errors])

    # 1件もデータがなかった場合は終了
    if not len(data):
        Dialog.show_messagebox(root, type="error", title="抽出エラー", message=f"1件もデータが抽出できませんでした。終了します。\n\nFile(s):\n{ers}")
        sys.exit()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
    root.destroy()
