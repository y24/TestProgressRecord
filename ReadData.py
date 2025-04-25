from collections import defaultdict
import pprint

from libs import OpenpyxlWrapper as Excel
from libs import Logger
from libs import Utility

logger = Logger.get_logger(__name__, console=True, file=False, trace_line=False)

# 日付ごとのデータ集計
def get_daily(data, results: list[str], completed_label:str, completed_results: list[str], filled_label:str, filled_results: list[str]):
    # 辞書を初期化：{日付: {結果タイプ: カウント}}
    result_count = defaultdict(lambda: defaultdict(int))
    
    for row in data:
        result, name, date = row
        
        # 日付が未設定の場合は特別な識別子「no_date」として扱う
        if not date: date = "no_date"

        # 各結果タイプのカウントを0で初期化し、結果が存在しない場合は0として表示されるようにする
        for keyword in results:
            result_count[date][keyword] = result_count[date].get(keyword, 0)
        result_count[date][completed_label] = result_count[date].get(completed_label, 0)
        result_count[date][filled_label] = result_count[date].get(filled_label, 0)

        # 結果の集計処理
        # 1. 個別の結果タイプをカウント
        if result in results:
            result_count[date][result] += 1
        # 2. 完了としてカウントすべき結果の場合、Completedとしてもカウント
        if result in completed_results:
            result_count[date][completed_label] += 1
        # 3. 消化としてカウントすべき結果の場合、Filledとしてもカウント
        if result in filled_results:
            result_count[date][filled_label] += 1

    # 集計結果を日付ありデータと日付なしデータに分離
    out_data = {}      # 日付ありデータ
    no_date_data = {}  # 日付なしデータ
    for date, counts in sorted(result_count.items()):
        counts = {**counts}  # 辞書のディープコピー
        if date == "no_date":
            no_date_data = counts
        else:
            out_data[date] = counts
    return out_data, no_date_data

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

# 対象外の数を取得
def get_excluded_count(data, targets:list[str]) -> int:
    return sum(1 for row in data if row and row[0] in targets)

# 全日付データ合計
def get_total_all_date(data, data_no_date, excludes:list[str]):
    result = {}
    # 全日付データ
    for values in data.values():
        for key, count in values.items():
            result[key] = result.get(key, 0) + count
    # 日付なしデータ
    for key, count in data_no_date.items():
        result[key] = result.get(key, 0) + count
    # Completedは除く
    for exclude in excludes:
        result.pop(exclude, None)
    return result

# 完了数を合計
def sum_completed_results(data: dict, completed_results: list) -> int:
    return sum(data.get(key, 0) for key in completed_results)

# 実施状況を判別
def make_run_status(count_stats: dict, settings: dict) -> str:
    if count_stats["filled"] == 0:
        # 未着手
        return settings["app"]["state"]["not_started"]["name"]
    elif count_stats["completed"] == count_stats["available"] and count_stats["incompleted"] == 0:
        # 完了
        return settings["app"]["state"]["completed"]["name"]
    elif count_stats["filled"] > 0:
        # 進行中
        return settings["app"]["state"]["in_progress"]["name"]
    else:
        return "???"

# Excelファイルからテスト結果データを読み取り、集計する関数
def aggregate_results(filepath:str, settings):
    # 設定された検索キーワードに基づいて対象シートを特定
    workbook = Excel.load(filepath)
    sheet_names = Excel.get_sheetnames_by_keywords(
        workbook, 
        keywords=settings["read_definition"]["sheet_search_keys"], 
        ignores=settings["read_definition"]["sheet_search_ignores"]
    )

    # 対象シートが見つからない場合はエラーを返却
    if len(sheet_names) == 0:
        return {
            "error": {
                "type": "sheet_not_found",
                "message": "シートが見つかりませんでした。"
            }
        }

    # 集計用の変数を初期化
    all_data = []          # 全シートの生データを格納
    data_by_env = {}       # 環境別の集計データを格納
    counts_by_sheet = []   # シート別の件数情報を格納

    # 各シートのデータを処理
    for sheet_name in sheet_names:
        # シートごとのデータを処理して取得
        sheet_data = _process_sheet(workbook=workbook, sheet_name=sheet_name, settings=settings)
        
        # エラーが発生した場合は即時返却
        if "error" in sheet_data:
            return sheet_data
        # 正常にデータが取得できた場合は集計用変数に追加
        elif sheet_data:
            all_data.extend(sheet_data["data"])           # 生データを追加
            data_by_env.update(sheet_data["env_data"])    # 環境別データを追加
            counts_by_sheet.append(sheet_data["counts"])   # 件数情報を追加

    # 全シートの集計データを生成して返却
    return _aggregate_final_results(
            all_data=all_data,           # 全シートの生データ
            data_by_env=data_by_env,     # 環境別の集計データ
            counts_by_sheet=counts_by_sheet,  # シート別の件数情報
            settings=settings            # 設定情報
        )

def _process_sheet(workbook, sheet_name: str, settings: dict):
    sheet = Excel.get_sheet_by_name(workbook=workbook, sheet_name=sheet_name)
    header_rownum = Excel.find_row(sheet, search_col=settings["read_definition"]["header"]["search_col"], search_str=settings["read_definition"]["header"]["search_key"])

    if not header_rownum:
        return {
            "error": {
                "type": "header_not_found",
                "message": "ヘッダー行が見つかりませんでした。"
            }
        }

    # ヘッダ行を取得
    header = Excel.get_row_values(sheet=sheet, row_num=header_rownum)

    # 列番号(例:環境別)を取得
    # 結果
    result_rows = Utility.find_colnum_by_keywords(lst=header, keywords=settings["read_definition"]["result_row"]["keys"], ignore_words=settings["read_definition"]["result_row"]["ignores"])
    # 担当者
    person_rows = Utility.find_colnum_by_keywords(lst=header, keywords=settings["read_definition"]["person_row"]["keys"])
    # 日付
    date_rows = Utility.find_colnum_by_keywords(lst=header, keywords=settings["read_definition"]["date_row"]["keys"])

    # 結果,担当者,日付の列セットが見つからないor同数でない場合はエラー
    if Utility.check_lists_equal_length(result_rows, person_rows, date_rows) == False:
        return {
            "error": {
                "type": "inconsistent_result_set",
                "message": "結果列のセットが正しく取得できませんでした。"
            }
        }

    # 列番号のセット(結果、担当者、日付)を作成
    sets = Utility.transpose_lists(result_rows, person_rows, date_rows)
    
    # 各セット処理
    data = []
    env_data = {}  # 環境データを格納する辞書を初期化
    
    for set in sets:
        # セットのデータ取得
        set_data = Excel.get_columns_data(sheet=sheet, col_nums=set, header_row=header_rownum, ignore_header=True)

        # 担当者名がNoneで結果と日付が存在する場合、"NO_NAME"に置き換え
        processed_data = []
        for row in set_data:
            if row[0] is not None and row[2] is not None and row[1] is None:
                row = [row[0], "NO_NAME", row[2]]
            processed_data.append(row)

        # 全セット合計のデータにも追加
        data.extend(processed_data)

        # そのセットの1行目からセット名を取得(セル内改行は_に置換)
        set_name = Excel.get_cell_value(sheet=sheet, col=set[0], row=1, replace_newline=True)

        # セット名がない場合はスキップ
        if not set_name:
            continue

        # 環境名
        env_name = f"[{sheet_name}]{set_name}"

        # 環境ごとのデータ集計
        env_data[env_name], _ = get_daily(
            data=processed_data, 
            results=settings["test_status"]["results"], 
            completed_label=settings["test_status"]["labels"]["completed"], 
            completed_results=settings["test_status"]["completed_results"],
            filled_label=settings["test_status"]["labels"]["filled"],
            filled_results=settings["test_status"]["filled_results"]
        )

    # 環境数
    env_count = len(sets)

    # テストケース数を計算
    tobe_rownunms = Utility.find_colnum_by_keywords(lst=header, keywords=settings["read_definition"]["tobe_row"]["keys"])
    tobe_data = Excel.get_column_values(sheet=sheet, col_nums=tobe_rownunms,header_row=header_rownum, ignore_header=True)
    case_count = sum(1 for item in tobe_data if any(x is not None for x in item))

    # 結果を返却
    return {
        "data": data,
        "env_data": env_data,
        "counts": {
            "sheet_name": sheet_name,
            "env_count": env_count,
            "all": case_count
        }
    }

def _aggregate_final_results(all_data, data_by_env, counts_by_sheet, settings):
    # 全セット集計(日付別)
    data_daily_total, data_no_date = get_daily(
        data=all_data,
        results=settings["test_status"]["results"],
        completed_label=settings["test_status"]["labels"]["completed"],
        completed_results=settings["test_status"]["completed_results"],
        filled_label=settings["test_status"]["labels"]["filled"],
        filled_results=settings["test_status"]["filled_results"]
    )

    # 全セット集計(担当者別)
    data_by_name = get_daily_by_name(all_data)
    
    # 全セット集計(全日付＋日付なし)
    data_total = get_total_all_date(data_daily_total, data_no_date, excludes=[settings["test_status"]["labels"]["completed"], settings["test_status"]["labels"]["filled"]])

    # 総テストケース数
    case_count_all = sum(item['env_count'] * item['all'] for item in counts_by_sheet)
    # 対象外テストケース数
    excluded_count = get_excluded_count(data=all_data, targets=settings["read_definition"]["excluded"])
    # 有効テストケース数
    available_count = case_count_all - excluded_count
    # 消化テストケース数
    filled_count = sum(data_total.values())
    # 完了テストケース数
    completed_count = sum_completed_results(data_total, settings["test_status"]["completed_results"])
    # 未実施テストケース数(マイナスは0)
    incompleted_count = max(0, available_count - filled_count)

    # 集計データ
    count_stats = {
        "all": case_count_all,
        "excluded": excluded_count,
        "available": available_count,
        "filled": filled_count,
        "completed": completed_count,
        "incompleted": incompleted_count
    }

    # 実施状況
    run_status = make_run_status(count_stats, settings)

    # 開始日・最終更新日
    start_date = None
    last_update = None
    # 日付別データがある場合
    if not Utility.is_empty(data_daily_total):
        # 開始日を取得
        start_date = min(data_daily_total.keys())
        # かつステータスが完了または進行中の場合
        if run_status == settings["app"]["state"]["completed"]["name"] or run_status == settings["app"]["state"]["in_progress"]["name"]:
            # 最終更新日を取得
            last_update = max(data_daily_total.keys())

    # 実施状況データ
    run_data = {
        "status": run_status,
        "start_date": start_date,
        "last_update": last_update
    }

    # 最終出力データ
    out_data = {
        "stats": count_stats,
        "run": run_data,
        "count_by_sheet": counts_by_sheet,
        "daily": data_daily_total,
        "total": data_total,
        "by_name": data_by_name,
        "by_env": data_by_env
    }

    # データチェック
    if case_count_all == 0:
        out_data["warning"] = {"type": "no_data", "message": "項目数を取得できませんでした。"}
    elif filled_count > available_count:
        out_data["warning"] = {"type": "inconsistent_count", "message": "テストケースの完了数が項目数を上回っています。"}

    return out_data

# コンソール出力
def console_out(data):
    filename = data["file"]
    logger.info(f"FILE: {filename}")
    logger.info(f"CASES COUNT: {data['count']}")
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
