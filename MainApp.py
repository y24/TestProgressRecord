import tkinter as tk
from tkinter import ttk, filedialog
from tktooltip import ToolTip
import subprocess
import csv, os, sys, pprint
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib_fontja
import numpy as np
import json

from libs import Utility
from libs import AppConfig
from libs import Dialog
from libs import Project
from libs import DataConversion
from libs import Labels
from libs import Clipboard
from libs import CsvFile
from libs import FileOperation

project_data = None
project_path = None
show_byfile_graph = None
show_env_data = None

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
    executed = settings["test_status"]["labels"]["executed"]

    if structure == 'daily':
        return [values[0]] + [values[1].get(k, 0) for k in Utility.sort_by_master(
            master_list=settings["test_status"]["results"] + [executed, completed],
            input_list=all_keys
        )]
    elif structure == 'by_env':
        env, date, data = values
        return [env, date] + [data.get(k, 0) for k in Utility.sort_by_master(
            master_list=settings["test_status"]["results"] + [executed, completed],
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
            master_list=settings["test_status"]["results"] + [settings["test_status"]["labels"]["executed"]],
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
    executed = settings["test_status"]["labels"]["executed"]
    
    column_definitions = {
        'by_env': ["環境名", "日付"],
        'by_name': ["日付", "担当者", "消化数"],
        'daily': ["日付"]
    }
    
    base_columns = column_definitions.get(structure, [])
    
    # by_nameの場合は基本列のみ返す
    if structure == 'by_name':
        return base_columns
        
    # その他の場合は結果列を追加
    result_columns = Utility.sort_by_master(
        master_list=settings["test_status"]["results"] + [executed, completed],
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
                    command=lambda: CsvFile.save_2d_array(
                        treeview_to_array(tree), 
                        f'{os.path.splitext(file_name)[0]}_{settings["app"]["structures"][structure]}'
                    ))
    menu.add_command(label="クリップボードにコピー", 
                    command=lambda: Clipboard.copy_2d_array(treeview_to_array(tree), root))
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

def update_byfile_tab(selected_file, count_label, last_load_time_label, ax, canvas, notebook):
    # タブ切り替え
    current_tab = notebook.index(notebook.select()) if notebook.tabs() else 0
    for widget in notebook.winfo_children():
        widget.destroy()
    
    data = next(item for item in input_data if item['selector_label'] == selected_file)

    # ファイル別結果
    if "error" in data:
        stats_data = {'all': 0, 'excluded': 0, 'available': 0, 'executed': 0, 'completed': 0, 'incompleted': 0}
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
    update_info_label(data=stats_data, count_label=count_label, last_load_time_label=last_load_time_label, detail=True)
    update_bar_chart(data=total_data, incompleted_count=incompleted, ax=ax, canvas=canvas, show_label=False)

    # ツールチップを更新
    if not hasattr(canvas.get_tk_widget(), '_tooltip'):
        # 初回のみツールチップを作成
        canvas.get_tk_widget()._tooltip = ToolTip(canvas.get_tk_widget(), msg="", delay=0.3, follow=False)
    
    # 既存のツールチップのメッセージを更新
    graph_tooltip = f'{Labels.make_results_text(total_data, incompleted)}'
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
    return lambda event: FileOperation.run(file_path=filepath)

def set_state_color(label, state_name):
    # キー名を取得
    state_key = Utility.find_key_by_name(settings["app"]["state"], state_name)
    # 見つからない場合は何もしない
    if state_key is None: return
    # 色を設定
    state_info = settings["app"]["state"][state_key]
    label.config(foreground=state_info["foreground"], background=state_info["background"])

def create_file_row(frame, row, index, file_data, display_data):
    """ファイルデータ行を作成する"""
    # セルの余白
    padx = 1
    pady = 3
    # 列番号の初期化
    col = 0
    
    # インデックス
    ttk.Label(frame, text=index).grid(
        row=row, column=col, sticky=tk.W+tk.E, padx=padx, pady=pady
    )
    col += 1

    # ファイル名
    filename = file_data['file']
    filename_label = ttk.Label(frame, text=filename)
    filename_label.grid(row=row, column=col, sticky=tk.W, padx=padx, pady=pady)
    tooltip_text = [filename]
    col += 1

    # ファイル名ダブルクリック時
    filepath = file_data['filepath']
    filename_label.bind("<Double-Button-1>", create_click_handler(filepath))

    #項目数
    case_count_label = ttk.Label(frame, text=display_data["available"])
    case_count_label.grid(row=row, column=col, sticky=tk.W+tk.E, padx=padx, pady=pady)
    col += 1

    # 消化率・完了率ラベル
    if display_data["on_error"]:
        executed_rate_display = "Error"
        executed_rate_tooltip = "消化率: -"
        comp_rate_display = "Error"
        comp_rate_tooltip = "完了率: -"
    else:
        executed_rate_display = display_data["executed_rate_text"]
        executed_rate_tooltip = f'消化率: {display_data["executed_rate_text"]} ({display_data["executed"]}/{display_data["available"]})'
        comp_rate_display = display_data["comp_rate_text"]
        comp_rate_tooltip = f'完了率: {display_data["comp_rate_text"]} ({display_data["completed"]}/{display_data["available"]})'

    # 消化率
    executed_rate_label = ttk.Label(frame, text=executed_rate_display)
    executed_rate_label.grid(row=row, column=col, sticky=tk.W+tk.E, padx=padx, pady=pady)
    ToolTip(executed_rate_label, msg=executed_rate_tooltip, delay=0.3, follow=False)
    col += 1

    # 完了率
    comp_rate_label = ttk.Label(frame, text=comp_rate_display)
    comp_rate_label.grid(row=row, column=col, sticky=tk.W+tk.E, padx=padx, pady=pady)
    ToolTip(comp_rate_label, msg=comp_rate_tooltip, delay=0.3, follow=False)
    col += 1

    # エラー時赤色・ワーニング時オレンジ色
    if display_data["on_error"] or display_data["on_warning"]:
        color = "red" if display_data["on_error"] else "darkorange2"
        filename_label.config(foreground=color)
        executed_rate_label.config(foreground=color)
        comp_rate_label.config(foreground=color)

    # グラフ表示ON、かつエラーではない場合は進捗グラフを表示
    if show_byfile_graph.get() and not display_data["on_error"]:
        # 進捗グラフ
        fig, ax = plt.subplots(figsize=(2, 0.1))
        canvas = FigureCanvasTkAgg(fig, master=frame)
        plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
        canvas.get_tk_widget().grid(row=row, column=col, sticky=tk.W+tk.E, padx=padx, pady=pady)

        # グラフを更新
        update_bar_chart(
            data=display_data["total_data"],
            incompleted_count=display_data["incompleted"],
            ax=ax,
            canvas=canvas,
            show_label=False
        )

        # グラフのツールチップ
        graph_tooltop = Labels.make_graph_tooltip(display_data)
        ToolTip(canvas.get_tk_widget(), msg=graph_tooltop, delay=0.3, follow=False)

    # エラー・ワーニング時はツールチップにメッセージを追加
    if display_data["on_error"] or display_data["on_warning"]:
        tooltip_text.append(f'{display_data["error_message"]}')

    # ファイル名のツールチップ
    tooltip_text.append("<ダブルクリックで開きます>")
    ToolTip(filename_label, msg="\n".join(tooltip_text), delay=0.3, follow=False)

def create_export_data(input_data: list, settings: dict) -> list:
    """エクスポート用のデータを生成する

    Args:
        input_data: 集計データ
        settings: 設定情報

    Returns:
        list: エクスポート用の2次元配列データ
    """
    # クリップボード出力用のヘッダ
    export_headers = ["No.", "ファイル名", "項目数", "完了数", "消化率", "完了率"]
    export_data = [export_headers + settings["test_status"]["results"] + [settings["test_status"]["labels"]["not_run"]]]

    # 各ファイルのデータを追加
    for index, file_data in enumerate(input_data, 1):
        display_data = DataConversion._extract_file_data(file_data)
        export_row = [
            index,  # No.
            file_data['file'],  # ファイル名
            display_data["available"],  # 項目数
            display_data["completed"],  # 完了数
            display_data["executed_rate_text"],  # 消化率
            display_data["comp_rate_text"]  # 完了率
        ]

        # エラー時は空文字を設定
        if display_data["on_error"]:
            export_row[4:] = ["", "", ""]  # 完了数、消化率、完了率を空に

        # テスト結果データを追加
        if not display_data["on_error"]:
            export_row += list(display_data["total_data"].values())
            export_row.append(display_data["incompleted"])

        export_data.append(export_row)

    return export_data

def update_filelist_table(table_frame):
    # テーブルの余白
    padx = 1
    pady = 3

    # テーブルコンテンツを配置するフレーム
    global content_frame
    content_frame = ttk.Frame(table_frame)
    content_frame.pack(fill=tk.BOTH, expand=True)

    # 制御ボタンを配置するフレーム
    control_frame = ttk.Frame(table_frame)
    control_frame.pack(fill=tk.X, padx=padx, pady=pady)

    def create_table(show_env=False):
        """テーブルを作成する
        
        Args:
            show_env: 環境別データを表示するかどうか
        """
        # ヘッダ
        headers = ["No.", "ファイル名", "項目数", "消化率", "完了率"]
        # 表示設定ONの場合はグラフ表示
        if show_byfile_graph.get():
            headers.append("テスト結果")

        # テーブル全体を格納するフレーム
        table_content_frame = ttk.Frame(content_frame)
        table_content_frame.pack(fill=tk.BOTH, expand=True, padx=padx, pady=padx)

        # 列の設定
        for i in range(len(headers)):
            table_content_frame.grid_columnconfigure(i, weight=1 if i == 1 else 0)

        # ヘッダー行
        for col, text in enumerate(headers):
            ttk.Label(
                table_content_frame,
                text=text,
                foreground="#444444",
                background="#e0e0e0",
                relief="solid"
            ).grid(row=0, column=col, sticky=tk.W+tk.E, padx=padx, pady=padx)

        # 各ファイルのデータ表示
        row = 1
        for index, file_data in enumerate(input_data, 1):
            display_data = DataConversion._extract_file_data(file_data)
            
            # ファイルデータ行
            create_file_row(table_content_frame, row, index, file_data, display_data)
            row += 1

            # 環境別データ表示がONの場合
            if show_env and "by_env" in file_data and file_data["by_env"]:
                for env_name, env_data in file_data["by_env"].items():
                    # 環境別データの集計
                    total_count = 0
                    executed_count = 0
                    completed_count = 0
                    last_update = None
                    
                    for date, results in env_data.items():
                        # 項目数の合計
                        total_count += sum(results.values())
                        # 消化済みの数
                        executed_count += sum(1 for v in results.values() if v > 0)
                        # 完了済みの数
                        completed_count += sum(1 for v in results.values() if v in settings["test_status"]["completed_results"])
                        # 最終更新日
                        if last_update is None or date > last_update:
                            last_update = date

                    # 環境別データ行
                    create_env_row(table_content_frame, row, env_name, total_count, executed_count, completed_count)
                    row += 1

    def create_env_row(frame, row, env_name, total_count,
                      executed_count, completed_count):
        """環境別データ行を作成する"""
        col = 0

        # インデックス（空）
        ttk.Label(frame, text="").grid(
            row=row, column=col, sticky=tk.W+tk.E, padx=padx, pady=padx
        )
        col += 1

        # 環境名
        ttk.Label(
            frame,
            text=env_name,
            foreground="#666666"
        ).grid(row=row, column=col, sticky=tk.W, padx=(20, padx), pady=padx)
        col += 1

        # 項目数
        ttk.Label(
            frame,
            text=total_count
        ).grid(row=row, column=col, sticky=tk.W+tk.E, padx=padx, pady=padx)
        col += 1

        # 消化率
        ttk.Label(
            frame,
            text=Labels.make_count_and_rate_text(executed_count, total_count)
        ).grid(row=row, column=col, sticky=tk.W+tk.E, padx=padx, pady=padx)
        col += 1

        # 完了率
        ttk.Label(
            frame,
            text=Labels.make_count_and_rate_text(completed_count, total_count)
        ).grid(row=row, column=col, sticky=tk.W+tk.E, padx=padx, pady=padx)

    # テーブルを作成
    create_table(show_env=show_env_data.get())

    # create_table関数をグローバルに公開
    global _create_table
    _create_table = create_table

def toggle_env_data():
    """環境別データの表示/非表示を切り替える"""
    global show_env_data

    # 設定ファイルを保存
    settings["app"]["show_env_data"] = show_env_data.get()
    AppConfig.save_settings(settings)
    
    # テーブルコンテンツのみをクリア
    for widget in content_frame.winfo_children():
        widget.destroy()
    # テーブルを再描画
    _create_table(show_env=show_env_data.get())

def create_summary_filelist_area(parent):
    # メニュー
    menu_frame = ttk.Frame(parent)
    menu_frame.pack(fill=tk.BOTH, padx=5, pady=2)

    # 並び替えメニュー
    sort_menu_button = ttk.Menubutton(menu_frame, text="ソート", direction="below")
    sort_menu = tk.Menu(sort_menu_button, tearoff=0)
    for key, order_info in settings["app"]["sort"]["orders"].items():
        sort_menu.add_command(label=order_info["label"], command=lambda key=key: change_sort_order(table_frame, key, sort_menu_button, on_change_order=True))
    sort_menu_button.config(menu=sort_menu)
    sort_menu_button.pack(anchor=tk.SW, side=tk.LEFT, padx=(0, 4))

    # エクスポートメニュー
    exp_menu_button = ttk.Menubutton(menu_frame, text="エクスポート", direction="below")
    expmenu = tk.Menu(exp_menu_button, tearoff=0)
    expmenu.add_command(label="CSVで保存", command=lambda: CsvFile.save_2d_array(_get_filelist_data(), f'{Utility.get_today_str()}'))
    expmenu.add_command(label="クリップボードにコピー", command=lambda: Clipboard.copy_2d_array(_get_filelist_data(), root))
    expmenu.add_command(label="ファイル名一覧をコピー", command=lambda: Clipboard.copy_2d_array(_get_filelist_data(filename_only=True), root))
    exp_menu_button.config(menu=expmenu)
    exp_menu_button.pack(anchor=tk.SW, side=tk.LEFT, padx=(0, 4))

    # ファイル別テーブル
    table_frame = ttk.Frame(parent)
    table_frame.pack(fill=tk.BOTH, expand=True, padx=10)

    # ファイル別テーブルの更新
    update_filelist_table(table_frame)

    # 初期ソート順の反映
    change_sort_order(table_frame, settings["app"]["sort"]["default"], sort_menu_button, on_change_order=False)

def select_write_file(entry):
    filepath = filedialog.askopenfilename(title="書込先ファイルを選択", defaultextension=".xlsx", filetypes=[("Excel file", "*.xlsx")])
    if filepath:  # キャンセルで空文字が返ってきたときは変更しない
        entry.delete(0, tk.END)  # 既存の内容をクリア
        entry.insert(0, filepath)  # 新しいファイルパスをセット

def edit_settings():
    response = Dialog.ask_question(root=root, title="環境設定", message=f"環境設定ファイルを開きますか？\n設定を反映するにはアプリを再起動するか、データの再集計を実行してください。")
    if response == "yes":
        FileOperation.run(file_path="UserConfig.json", exit=False)

def load_files():
    files = Dialog.select_files(("Excel/Zipファイル", "*.xlsx;*.zip"))
    if files:
        # project_dataがない場合は初期化
        global project_data
        if not project_data:
            project_data = {}
        
        # filesキーがない場合は初期化
        if "files" not in project_data:
            project_data["files"] = []
            
        # ローカルファイルをproject_data["files"]から除去
        project_data["files"] = [
            file for file in project_data["files"] 
            if file["type"] != "local"
        ]
        # 読込ファイルをプロジェクトデータに追加
        for file in files:
            project_data["files"].append({
                "type": "local",
                "identifier": "",
                "path": file
            })
        # プロジェクトデータを保存
        save_project()
        # 新しいプロセスでプロジェクト読込
        new_process(inputs=list(files), project_path=project_path, on_reload=False)

def open_project():
    project_path = Dialog.select_file(("JSONファイル", "*.json"))
    if project_path: new_process(inputs=[project_path], on_reload=False)

def update_labels():
    """プロジェクト情報に関連するラベルを更新する"""
    # プロジェクト名の更新
    project_name = project_data.get("project_name", "<<新規プロジェクト>>")
    project_name_label.config(text=project_name)
    # ウィンドウタイトルも更新
    update_window_title(root)

def edit_project(after_save_callback=None):
    """プロジェクト設定を編集する"""
    def on_project_updated(new_project_data):
        global project_data, project_path, change_flg
        try:
            project_path = new_project_data.get("project_path", None)  # パスを取り出して削除
            project_data = new_project_data  # グローバル変数を更新

            # ラベルを更新
            update_labels()

            # コールバックがあれば実行
            if after_save_callback:
                after_save_callback()

            # 編集中フラグOFF（編集画面を閉じるときに保存している）
            change_flg = False

            update_window_title(root)

        except Exception as e:
            Dialog.show_messagebox(
                root=root,
                type="error",
                title="更新エラー",
                message=f"プロジェクト情報の更新に失敗しました: {str(e)}"
            )

    # ProjectEditorAppを子ウインドウとして起動
    from ProjectEditorApp import ProjectEditorApp
    editor = ProjectEditorApp(
        parent=root,
        callback=on_project_updated,
        project_path=project_path,
        gathered_data=input_data
    )

def save_project():
    global project_data, project_path, change_flg
    """現在のinput_dataをJSONファイルに保存する"""
    
    # プロジェクトファイルが開かれていない場合は新規作成
    if not project_path:
        def after_create():
            save_project()
        edit_project(after_save_callback=after_create)
        return

    try:
        # JSONファイルに保存
        Project.save_to_json(file_path=project_path, input_data=input_data, project_data=project_data)

        # 編集中フラグOFF
        change_flg = False

        # ウィンドウタイトルを更新
        update_window_title(root)
       
    except Exception as e:
        # エラーメッセージを表示
        Dialog.show_messagebox(root, type="error", title="保存エラー", message=str(e))

def toggle_byfile_graph():
    """ファイル毎のグラフ表示を切り替える"""
    global show_byfile_graph
    settings["app"]["show_byfile_graph"] = show_byfile_graph.get()
    AppConfig.save_settings(settings)

    # ファイル別グラフの表示/非表示を切り替え
    for widget in content_frame.winfo_children():
        widget.destroy()
    # テーブルを再描画
    _create_table(show_env=show_env_data.get())

def _get_write_data():
    return DataConversion.convert_to_2d_array(data=input_data, settings=settings)

def new_project():
    """新規プロジェクトを作成する"""
    if change_flg:
        # プロジェクトデータが変更されている場合、確認ダイアログを表示
        response = confirm_save()
        if response == None:
            return  # キャンセルでアプリに戻る
        elif response == True:
            save_project()  # 保存
    # 空プロジェクトを開く
    new_process(inputs=[], project_path=None, on_reload=False)

def create_menubar(parent, has_data=False):
    global show_byfile_graph, show_env_data

    menubar = tk.Menu(parent)
    parent.config(menu=menubar)
    # File
    file_menu = tk.Menu(menubar, tearoff=0)
    file_menu.add_command(label="新規作成", command=new_project, accelerator="Ctrl+N")
    file_menu.add_command(label="開く", command=open_project, accelerator="Ctrl+O")
    file_menu.add_command(label="保存", command=save_project, accelerator="Ctrl+S")
    file_menu.add_separator()
    file_menu.add_command(label="ローカルファイルを読込", command=load_files, accelerator="Ctrl+L")
    file_menu.add_separator()
    file_menu.add_command(label="プロジェクト設定", command=edit_project, accelerator="Ctrl+E")
    file_menu.add_command(label="環境設定", command=edit_settings)
    file_menu.add_separator()
    file_menu.add_command(label="終了", command=parent.quit, accelerator="Ctrl+Q")
    menubar.add_cascade(label="ファイル", menu=file_menu)
    # Data
    data_menu = tk.Menu(menubar, tearoff=0)
    data_menu.add_command(label="再集計を実行", command=reload_files, accelerator="Ctrl+R")
    menubar.add_cascade(label="集計", menu=data_menu)
    # View
    view_menu = tk.Menu(menubar, tearoff=0)
    view_menu.add_checkbutton(label="ファイル別のグラフを表示", variable=show_byfile_graph, command=toggle_byfile_graph)
    view_menu.add_checkbutton(label="環境別の集計データを表示", variable=show_env_data, command=toggle_env_data)
    menubar.add_cascade(label="表示", menu=view_menu)

    # キーバインドの追加
    parent.bind('<Control-s>', lambda e: save_project())
    parent.bind('<Control-o>', lambda e: open_project())
    parent.bind('<Control-l>', lambda e: load_files())
    parent.bind('<Control-r>', lambda e: reload_files())
    parent.bind('<Control-e>', lambda e: edit_project())
    parent.bind('<Control-q>', lambda e: parent.quit())

    # ファイルデータがない場合はメニュー無効化
    if not has_data:
        data_menu.entryconfig("再集計を実行", state=tk.DISABLED)

def create_input_area(parent, settings):
    input_frame = ttk.LabelFrame(parent, text="集計データ出力")
    input_frame.pack(fill=tk.X, padx=5, pady=3)

    submit_frame = ttk.Frame(input_frame)
    submit_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=2, sticky=tk.W)
    ttk.Button(submit_frame, text="クリップボードにコピー", command=lambda: Clipboard.copy_2d_array(_get_write_data(), root), width=22).pack(side=tk.LEFT, padx=2, pady=(0,2))

def update_info_label(data, count_label, last_load_time_label=None, detail=True):
    if len(data) == 0 or"error" in data:
        # データなしまたはエラー時
        all = None
        available  = None
        completed = None
        executed = None
        excluded = None
    else:
        all = data["all"]
        available = data["available"]
        completed = data["completed"]
        executed = data["executed"]
        excluded = data["excluded"]

    # ケース数テキスト
    count = available if available else "--"
    count_text = f'項目数: {count}'
    if detail:
        count_text += f' (Total: {all or "-"} / 対象外: {excluded or "-"})'
    # 完了率テキスト
    completed_rate_text = f'完了率: {Utility.meke_rate_text(completed, available)} [{completed or "-"}/{available or "-"}]'
    # 消化率テキスト
    executed_rate_text = f'消化率: {Utility.meke_rate_text(executed, available)} [{executed or "-"}/{available or "-"}]'

    # 集計表示を更新
    count_label.config(text=f'{count_text}  |  {completed_rate_text}  |  {executed_rate_text}')

    # 読込日時
    if last_load_time_label:
        last_load_time_text = f'読込日時: {Utility.get_latest_time(input_data)}'
        # 読込日時を更新
        last_load_time_label.config(text=last_load_time_text)

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
        for bar, size, result_label, color in bars:
            # 表示用の値
            if size:
                if size > total * 0.12:
                    # 割合12%以上は(件数)も表示
                    display_label = f"{result_label} ({size})"
                elif size > total * 0.08:
                    # 割合8%以上はラベルのみ表示
                    display_label = result_label
                else:
                    # 割合8%未満はラベルなし
                    display_label = ""
            else:
                # データ0件はラベルなし
                display_label = ""

            # データがない場合はNO DATAを表示
            if result_label == "No Data": display_label = result_label

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
                display_label,
                ha='center', va='center', fontsize=8, 
                color=label_color
            )

    if canvas:
        canvas.draw()

def save_window_position():
    # ウインドウの位置情報とサイズを保存
    geometry = root.geometry()
    # geometryからサイズと位置情報を取得 (例: "400x300+200+100" -> "400x300" と "+200+100")
    size, position = geometry.split("+", 1)
    position = "+" + position
    settings["app"]["window_size"] = size
    settings["app"]["window_position"] = position
    AppConfig.save_settings(settings)

def confirm_save():
    """保存確認ダイアログを表示"""
    response = Dialog.ask_yes_no_cancel(
        root=root,
        title="確認",
        message="プロジェクトデータが変更されています。保存して終了しますか？"
    )
    return response

def on_closing():
    # ウインドウ終了時
    if change_flg:
        # プロジェクトデータが変更されている場合、確認ダイアログを表示
        response = confirm_save()
        if response == None:
            return  # キャンセルでアプリに戻る
        elif response == True:
            save_project()  # 保存
    
    # ウインドウの位置情報とサイズを保存
    save_window_position()
    # 終了
    root.quit()

def close_all_dialogs():
    """開いている全てのToplevelウィンドウ（ダイアログ含む）を閉じる"""
    for widget in root.winfo_children():
        if isinstance(widget, tk.Toplevel):
            widget.destroy()

def new_process(inputs, on_reload=False, project_path=None):
    """新しいプロセスを起動"""
    # ダイアログを閉じる
    close_all_dialogs()
    # コマンドライン引数の設定
    python = sys.executable
    command = [python, "StartProcess.py"] + inputs
    if on_reload: command += ["--on_reload"]
    if project_path: command += ["--project", project_path]
    # 新しいプロセスを起動
    print(f"command: {command}")
    subprocess.Popen(command)
    sys.exit()

def reload_files(pre_message:str="", show_time:bool=True):
    # 更新日時を取得
    last_loaded = Utility.get_latest_time(input_data)
    last_updated = Utility.get_latest_time(input_data, key="last_updated")

    # 更新日時表示
    if show_time:
        time = f"\n\n最終読込日時: {last_loaded}\n最終更新日時: {last_updated}"
    else:
        time = ""

    # ダイアログ表示
    response = Dialog.ask_question(root=root, title="確認", message=f"{pre_message}最新のデータを集計しますか？{time}")

    if response == "yes":
        # プロジェクトファイルを開いている場合
        if project_path:
            try:
                # プロジェクトファイルを読み込み
                with open(project_path, "r", encoding="utf-8") as f:
                    project_json = json.load(f)
                
                # 古い集計データを削除
                if "gathered_data" in project_json:
                    del project_json["gathered_data"]
                
                # プロジェクトファイル保存
                with open(project_path, "w", encoding="utf-8") as f:
                    json.dump(project_json, f, ensure_ascii=False, indent=2)
            except Exception as e:
                # エラー発生時
                Dialog.show_messagebox(root=root, type="error", title="保存エラー", message=f"プロジェクトファイルの更新に失敗しました。\n{str(e)}")
                return

        # 再集計用の新プロセス起動
        if not project_data.get("files"):
            # 取得元の設定がない場合、集計データからパスを取得して起動
            file_paths = [item["filepath"] for item in input_data if "filepath" in item]
            if len(file_paths) > 0:
                new_process(inputs=file_paths, project_path=project_path, on_reload=True)
            else:
                Dialog.show_messagebox(root=root, type="warning", title="読込ファイルなし", message=f"再読み込みするファイルが設定されていません。")
        else:
            # 取得元の設定がある場合、通常通りプロジェクトファイルを読み込む
            new_process(inputs=list(input_args), project_path=project_path, on_reload=True)

def create_global_tab(parent, has_data=False):
    nb = ttk.Notebook(parent)
    # 集計結果タブ
    tab1 = tk.Frame(nb)
    nb.add(tab1, text=' 集計結果 ')
    # ファイル別タブ
    tab2 = tk.Frame(nb)
    if has_data: nb.add(tab2, text=' ファイル別 ')
    nb.pack(fill=tk.BOTH, expand=True)
    return tab1, tab2

def create_initial_screen(parent):
    # 初期画面
    initial_frame = ttk.Frame(parent)
    initial_frame.pack(fill=tk.BOTH, expand=True)
    # 初期画面のテキスト
    initial_text = ttk.Label(initial_frame, text="集計対象のファイルを読み込んでください。", anchor="center")
    initial_text.pack(pady=(30, 0))
    # ファイル読込ボタン
    load_files_button = ttk.Button(initial_frame, text="ローカルファイル読込", command=load_files)
    load_files_button.pack(pady=15)

def create_summary_tab(parent, has_data=False):
    # 全体集計にはエラーとワーニングのあるデータを含めない
    filtered_data = Utility.filter_objects(input_data, exclude_keys=["error", "warning"])

    # 総合集計の表示エリア
    total_frame = ttk.Frame(parent)
    total_frame.pack(fill=tk.X, padx=5, pady=8)

    # プロジェクト名称
    global project_name_label
    project_name = project_data.get("project_name", "<<新規プロジェクト>>")
    project_name_label = ttk.Label(total_frame, text=project_name, anchor="w")
    project_name_label.pack(fill=tk.X, padx=20)
    project_name_label.config(cursor="hand2", font=("Meiryo UI", 9, "underline"))
    project_name_label.bind("<Button-1>", lambda e: edit_project())

    # グラフ表示(総合)
    total_fig, total_ax = plt.subplots(figsize=(8, 0.25))
    total_canvas = FigureCanvasTkAgg(total_fig, master=total_frame)
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    total_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

    # グラフ(総合)を更新
    incompleted = Utility.sum_values(filtered_data, "stats")["incompleted"] if len(filtered_data) > 0 else 0
    update_bar_chart(data=Utility.sum_values(filtered_data, "total"), incompleted_count=incompleted, ax=total_ax, canvas=total_canvas, show_label=True)

    # グラフ(総合)のツールチップを設定
    filtered_total_data = Utility.sum_values(filtered_data, "total")
    graph_tooltip = f'{Labels.make_results_text(filtered_total_data, incompleted)}'
    ToolTip(total_canvas.get_tk_widget(), msg=graph_tooltip, delay=0.3, follow=False)

    # テストケース数/完了率/消化率(総合)
    total_count_label = ttk.Label(total_frame, anchor="w")
    total_count_label.pack(fill=tk.X, padx=20)
    update_info_label(data=Utility.sum_values(filtered_data, "stats"), count_label=total_count_label, detail=True)

    # 区切り線
    separator = ttk.Separator(parent, orient="horizontal")
    separator.pack(fill=tk.X, padx=0, pady=5)

    if has_data:
        # ファイル別集計データ表示部を作成
        create_summary_filelist_area(parent=parent)

        # ファイル書き込みエリアを作成
        create_input_area(parent=parent, settings=settings)
    else:
        # 初期画面を作成
        create_initial_screen(parent=parent)

def create_byfile_tab(parent):
    by_file_frame = ttk.Frame(parent)
    by_file_frame.pack(fill=tk.X, padx=5, pady=5)

    # ファイル選択プルダウン
    file_selector = ttk.Combobox(by_file_frame, values=[file["selector_label"] for file in input_data], state="readonly")
    file_selector.pack(fill=tk.X, padx=20, pady=5)
    file_selector.bind("<<ComboboxSelected>>", lambda event: update_byfile_tab(selected_file=file_selector.get(), count_label=file_count_label, last_load_time_label=file_last_load_time_label, ax=file_ax, canvas=file_canvas, notebook=notebook))

    # グラフ表示(ファイル別)
    file_fig, file_ax = plt.subplots(figsize=(6, 0.1))
    file_canvas = FigureCanvasTkAgg(file_fig, master=by_file_frame)
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    file_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

    # テストケース数/完了率/消化率(ファイル別)
    file_count_label = ttk.Label(by_file_frame, anchor="w")
    file_count_label.pack(fill=tk.X, padx=20)

    # 最終読込日時(ファイル別)
    file_last_load_time_label = ttk.Label(by_file_frame, anchor="w")
    file_last_load_time_label.pack(fill=tk.X, padx=20)

    # タブ表示
    notebook_height = 300 if len(input_data) > 1 else 355
    notebook = ttk.Notebook(by_file_frame, height=notebook_height)
    notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # グリッド表示
    if input_data:
        file_selector.current(0)
        update_byfile_tab(selected_file=input_data[0]['selector_label'], count_label=file_count_label, last_load_time_label=file_last_load_time_label, ax=file_ax, canvas=file_canvas, notebook=notebook)

def update_window_title(root):
    global project_data, change_flg
    # プロジェクト名（未設定の場合は"名称未設定"）
    title = project_data.get('project_name', '名称未設定')
    # 編集中フラグがONの場合は"*"を表示
    if change_flg:
        title = f"* {title}"
    # ウィンドウタイトルを更新
    root.title(f"TestTraQ - {title}")

def _show_startup_messages(has_data, pjpath, on_reload, input_data):
    """起動時のメッセージを表示する

    Args:
        has_data: データの有無
        pjpath: プロジェクトファイルパス
        on_reload: 再読込フラグ
        input_data: 入力データ
    """
    global root

    # データ抽出に失敗したファイルのリスト
    errors = [r for r in input_data if "error" in r]
    ers = "\n".join(["  "+ err["file"] for err in errors])

    # エラーあり、かつ初回集計時にメッセージ（リロードして2回目以降はメッセージなし）
    if len(errors) and not on_reload:
        Dialog.show_messagebox(root, type="error", title="抽出エラー", message=f"以下のファイルはデータが抽出できませんでした。\n\nFile(s):\n{ers}")

    if has_data:
        # プロジェクトファイルを読込時、ファイルが更新されている場合は再集計を促す
        if pjpath and not on_reload and _needs_reload(input_data):
            reload_files(pre_message="ファイルが更新されています。", show_time=True)
        # 再集計後の起動時にはプロジェクトを保存
        if pjpath and on_reload:
            save_project()

def _needs_reload(input_data):
    """再集計が必要かどうかを判定する

    Args:
        input_data: 集計データ

    Returns:
        bool: 再集計が必要な場合はTrue
    """
    for data in input_data:
        if "last_loaded" in data and "last_updated" in data:
            last_loaded = datetime.strptime(data["last_loaded"], "%Y-%m-%d %H:%M:%S")
            last_updated = datetime.strptime(data["last_updated"], "%Y-%m-%d %H:%M:%S")
            if last_updated > last_loaded:
                return True
    return False

def _update_file_timestamps(input_data):
    """input_dataの各ファイルの更新日時をチェックして更新する
    sourceがlocalのファイルのみ更新日時をチェックする

    Args:
        input_data: 集計データ

    Returns:
        list: 更新された集計データ
    """
    updated_data = []
    for data in input_data:
        if "filepath" in data and data.get("source") == "local":
            filepath = data["filepath"]
            if os.path.exists(filepath):
                # ファイルの最終更新日時を取得
                last_modified = datetime.fromtimestamp(os.path.getmtime(filepath))
                # 文字列形式に変換
                last_modified_str = last_modified.strftime("%Y-%m-%d %H:%M:%S")
                # 更新日時を更新
                data["last_updated"] = last_modified_str
        updated_data.append(data)
    return updated_data

def run(pjdata=None, pjpath=None, indata=None, args=None, on_reload=False):
    global root, input_data, settings, input_args, project_data, project_path, change_flg
    global show_byfile_graph, show_env_data

    # データの初期化
    project_data = pjdata or {}
    project_path = pjpath
    input_data = indata or []

    # 空プロジェクトかどうかのフラグ
    has_data = len(input_data)

    # 編集中フラグの初期化
    change_flg = False
    # プロジェクトファイル未保存かつファイルがある場合は編集中フラグON
    if not pjpath and has_data: change_flg = True

    # プロジェクトファイルを開いた場合、各ファイルの更新日時を一時的に最新化（編集中フラグは変更しない）
    if pjpath:
        input_data = _update_file_timestamps(input_data)

    # 起動時に指定したファイルパス（再読込用）
    input_args = args or []

    # 設定のロード
    settings = AppConfig.load_settings()

    # 起動
    # 親ウインドウ生成
    root = tk.Tk()
    # ウィンドウタイトルを更新
    update_window_title(root)
    # ウインドウ表示位置を復元
    geometry = f"{settings['app']['window_size']}{settings['app']['window_position']}"
    root.geometry(geometry)

    # BooleanVarを作成し、設定値を反映
    show_byfile_graph = tk.BooleanVar(value=settings["app"]["show_byfile_graph"])
    show_env_data = tk.BooleanVar(value=settings["app"].get("show_env_data", False))

    # メニューバー生成
    create_menubar(parent=root, has_data=has_data)

    # 起動時メッセージ
    _show_startup_messages(has_data, project_path, on_reload, input_data)

    # グローバルタブ生成
    tab1, tab2 = create_global_tab(parent=root, has_data=has_data)
    # タブ1：全体集計タブ
    create_summary_tab(tab1, has_data=has_data)
    # タブ2：ファイル別集計タブ
    if has_data: create_byfile_tab(tab2)

    # ウインドウ終了時の処理
    root.protocol("WM_DELETE_WINDOW", on_closing)
    # メインループ
    root.mainloop()
    # ウインドウ破棄
    root.destroy()

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

def change_sort_order(table_frame, order, sort_menu_button, on_change_order=False):
    sort_input_data(order, type=settings["app"]["sort"]["orders"][order]["type"])
    clear_frame(table_frame) # 表示をクリア
    if on_change_order:
        plt.close('all') # 表示していたグラフを開放する
    update_filelist_table(table_frame)
    sort_menu_button.config(text=f'ソート: {settings["app"]["sort"]["orders"][order]["label"]}')
    if on_change_order:
        # デフォルト設定に保存
        settings["app"]["sort"]["default"] = order
        AppConfig.save_settings(settings)

def _get_filelist_data(filename_only: bool = False):
    if filename_only:
        return [[row[1]] for row in DataConversion.create_export_data(input_data, settings)[1:]]
    else:
        return DataConversion.create_export_data(input_data, settings)

if __name__ == "__main__":
    run()
