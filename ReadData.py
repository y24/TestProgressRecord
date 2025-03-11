
import os, sys, pprint
from tqdm import tqdm
from collections import defaultdict

import App
from libs import OpenpyxlWrapper as Excel
from libs import Dialog
from libs import Logger
from libs import Utility
from libs import Zip
from libs import AppConfig
logger = Logger.get_logger(__name__, console=True, file=False, trace_line=False)

# データ集計
def get_daily(data, results: list[str], completed_results: list[str]):
    total_label = "Completed"
    result_count = defaultdict(lambda: defaultdict(int))
    
    for row in data:
        result, name, date = row
        
        # 日付が空のデータはスキップ
        if not date:
            continue

        # 各結果を0で初期化
        for keyword in results:
            result_count[date][keyword] = result_count[date].get(keyword, 0)
        result_count[date][total_label] = result_count[date].get(total_label, 0)

        # 結果列がフィルタ文字列に合致するものだけ抽出
        if result in results:
            result_count[date][result] += 1
        if result in completed_results:
            result_count[date][total_label] += 1
    
    # 出力
    out_data = {}
    for date, counts in sorted(result_count.items()):
        counts = {**counts}
        out_data[date] = counts
    
    return out_data

# データ集計（名前別）
def get_daily_by_name(data):
    date_name_count = defaultdict(lambda: defaultdict(int))

    # 日付が空の行は削除
    data = [row for row in data if len(row) > 2 and row[2] not in ("", None)]

    # 結果が空ではない行を日付および名前ごとにカウント
    for row in data:
        result, name, date = row
        if result:  # 結果が空ではない場合
            date_name_count[date][name] += 1

    # 集計結果を返却
    out_data = {}
    for date, name_counts in sorted(date_name_count.items()):
        daily_count = {}
        for name, count in sorted(name_counts.items()):
            daily_count[name] = count
        out_data[date] = daily_count
    return out_data

# 処理
def aggregate_results(filepath:str):

    # 対象シート名を取得
    workbook = Excel.load(filepath)
    sheet_names = Excel.get_sheetnames_by_keywords(workbook, keywords=settings["read"]["sheet_search_keys"], ignores=settings["read"]["sheet_search_ignores"])

    # シートがない場合はスキップ
    if len(sheet_names) == 0:
        # logger.error(f"Sheet not found. ({filepath})")
        return

    # 各シート処理
    all_data = []
    data_by_env = {}
    for sheet_name in sheet_names:
        # シート取得
        sheet = Excel.get_sheet_by_name(workbook=workbook, sheet_name=sheet_name)

        # ヘッダ行を探して取得
        header_rownum = Excel.find_row(sheet, search_col=settings["read"]["header"]["search_col"], search_str=settings["read"]["header"]["search_key"])
        header = Excel.get_row_values(sheet=sheet, row_num=header_rownum)

        # 列セット(例:環境別)を取得
        # 結果
        result_rows = Utility.find_rownum_by_keyword(list=header, keyword=settings["read"]["result_row"]["key"], ignore_words=settings["read"]["result_row"]["ignores"])
        # 担当者
        person_rows = Utility.find_rownum_by_keyword(list=header, keyword=settings["read"]["person_row"]["key"])
        # 日付
        date_rows = Utility.find_rownum_by_keyword(list=header, keyword=settings["read"]["date_row"]["key"])

        # セットが正しく取得できない場合はスキップ
        if Utility.check_lists_equal_length(result_rows, person_rows, date_rows) == False:
            # logger.error(f"Failed to get result column. (Sheet: {sheet_name})")
            continue

        # 行番号のセット(結果、担当者、日付)を作成
        sets = Utility.transpose_lists(result_rows, person_rows, date_rows)
        
        # 各セット処理
        for set in sets:
            # セットのデータ取得
            set_data = Excel.get_columns_data(sheet=sheet, col_nums=set, header_row=header_rownum, ignore_header=True)

            # 全セット合計のデータにも追加
            all_data = all_data + set_data

            # そのセットの1行目からセット名を取得
            set_name = Excel.get_cell_value(sheet=sheet, col=set[0], row=1)

            # セット名がない場合はスキップ
            if not set_name:
                # logger.error(f"Failed to get environment name. (Sheet: {sheet_name})")
                continue

            # 環境名
            env_name = f"{sheet_name}_{set_name}"

            # 環境ごとのデータ集計
            data_by_env[env_name] = get_daily(data=set_data, results=settings["common"]["results"], completed_results=settings["common"]["completed_results"])

    # 全セット集計(日付別)
    data_total = get_daily(data=all_data, results=settings["common"]["results"], completed_results=settings["common"]["completed_results"])

    # 全セット集計(担当者別)
    data_by_name = get_daily_by_name(data=all_data)

    # テストケース数カウント
    # 期待結果列の番号
    tobe_rows = Utility.find_rownum_by_keyword(list=header, keyword=settings["read"]["tobe_row"]["key"])
    # 期待結果列を取得
    tobe_data = Excel.get_column_values(sheet=sheet, col_nums=tobe_rows,header_row=header_rownum, ignore_header=True)
    # テストケース数
    case_count = sum(1 for item in tobe_data if any(x is not None for x in item))

    # 結果返却
    return {
        "case_count": case_count,
        "total": data_total,
        "by_name": data_by_name,
        "by_env": data_by_env
    }

def make_selector_label(file):
    file_name = file["file"]
    relative_path = f'[{file["relative_path"]}] ' if file["relative_path"] else ""
    return f"{relative_path}{file_name}"

# コンソール出力
def console_out(data):
    filename = data["file"]
    logger.info(f"FILE: {filename}")
    logger.info(f"CASES COUNT: {data['case_count']}")
    logger.debug(" ")

    # インデント
    dep = "  "

    # 全セット
    if any(data["total"]):
        logger.info("[Total]")
        for key, d in data["total"].items():
            logger.debug(dep + f"{key}: {d}")
        logger.debug(" ")

        logger.info("[By name]")
        for key, d in data["by_name"].items():
            logger.debug(dep + f"{key}: {d}")
        logger.debug(" ")

    # 出力(セット別)
    if any(data["by_env"]):
        logger.info("[By environment]")
        for key, d in data["by_env"].items():
            logger.debug(dep + f"{key}:")
            for key, d in d.items():
                logger.debug(dep * 2 + f"{key}: {d}")
        logger.debug(" ")

    logger.debug(" ")
    logger.info("~" * 50)


if __name__ == "__main__":
    # 設定読み込み
    global settings
    settings = AppConfig.load_settings()

    inputs = []
    files = []
    temp_dirs = []
    out_data = []
    errors = []

    # 引数でファイルパスを受け取る(複数可)
    inputs = sys.argv.copy()

    # 引数がない場合はファイル選択ダイアログ(複数可)
    if len(inputs) <= 1:
        inputs = Dialog.select_files(("Excel/Zipファイル", "*.xlsx;*.zip"))
        if not inputs:
            # キャンセル時は終了
            sys.exit()
    else:
        del inputs[0]

    # 拡張子判定
    for input in inputs:
        ext = Utility.get_ext_from_path(input)
        if ext == "xlsx":
            files.append({"fullpath": input, "temp_dir": ""})
        elif ext == "zip":
            # zipファイル展開
            extracted_files, temp_dir = Zip.extract_files_from_zip(input, extensions=['.xlsx'])
            for f in extracted_files:
                files.append({"fullpath": f, "temp_dir": temp_dir})
            temp_dirs.append(temp_dir)

    # ファイルを処理
    for file in tqdm(files):
        # 集計
        result = aggregate_results(filepath=file["fullpath"])
        # 出力
        if result and not Utility.is_empty(result):
            # ファイルパス
            result["file"] = Utility.get_filename_from_path(filepath=file["fullpath"])
            # zipファイル内の相対パスを取得
            if file["temp_dir"]:
                result["relative_path"] = Utility.get_relative_directory_path(full_path=file["fullpath"], base_dir=file["temp_dir"])
                result["selector_label"] = make_selector_label(result)
            else:
                # zipファイルではない場合
                result["relative_path"] = ""
                result["selector_label"] = result["file"]
            # コンソール出力
            # console_out(result)
            # ビューアに渡す配列に格納
            out_data.append(result)
        else:
            # 出力データがない場合
            errors.append(Utility.get_filename_from_path(filepath=file["fullpath"]))

    # pprint.pprint(out_data)

    # ビューア起動
    App.load_data(out_data, errors)

    # zipファイルを展開していた場合は一時フォルダを掃除
    if len(temp_dirs): Zip.cleanup_old_temp_dirs()
