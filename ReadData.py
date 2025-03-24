
from collections import defaultdict
import pprint

from libs import OpenpyxlWrapper as Excel
from libs import Logger
from libs import Utility

logger = Logger.get_logger(__name__, console=True, file=False, trace_line=False)

# データ集計
def get_daily(data, results: list[str], completed_label:str, completed_results: list[str]):
    result_count = defaultdict(lambda: defaultdict(int))
    
    for row in data:
        result, name, date = row
        
        # 日付が空のデータはスキップ
        if not date:
            continue

        # 各結果を0で初期化
        for keyword in results:
            result_count[date][keyword] = result_count[date].get(keyword, 0)
        result_count[date][completed_label] = result_count[date].get(completed_label, 0)

        # 結果列がフィルタ文字列に合致するものだけ抽出
        if result in results:
            result_count[date][result] += 1
        if result in completed_results:
            result_count[date][completed_label] += 1
    
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

# 対象外の数を取得
def get_excluded_count(data, targets:list[str]) -> int:
    return sum(1 for row in data if row and row[0] in targets)

# 全日付データ合計
def get_total_all_date(data, exclude:str):
    result = {}
    for values in data.values():
        for key, count in values.items():
            result[key] = result.get(key, 0) + count
    # Completedは除く
    result.pop(exclude, None)
    return result

def sum_completed_results(data: dict, completed_results: list) -> int:
    return sum(data.get(key, 0) for key in completed_results)

# 処理
def aggregate_results(filepath:str, settings):

    # 対象シート名を取得
    workbook = Excel.load(filepath)
    sheet_names = Excel.get_sheetnames_by_keywords(workbook, keywords=settings["read"]["sheet_search_keys"], ignores=settings["read"]["sheet_search_ignores"])

    # シートがない場合はエラー
    if len(sheet_names) == 0:
        return {
            "error": {
                "type": "sheet_not_found",
                "message": "シートが見つかりませんでした。"
            }
        }

    # 各シート処理
    all_data = []
    data_by_env = {}
    counts_by_sheet = []
    for sheet_name in sheet_names:
        sheet_data = _process_sheet(workbook=workbook, sheet_name=sheet_name, settings=settings)
        if "error" in sheet_data:
            return sheet_data
        elif sheet_data:
            all_data.extend(sheet_data["data"])
            data_by_env.update(sheet_data["env_data"])
            counts_by_sheet.append(sheet_data["counts"])

    return _aggregate_final_results(
            all_data=all_data,
            data_by_env=data_by_env,
            counts_by_sheet=counts_by_sheet,
            settings=settings
        )

def _process_sheet(workbook, sheet_name: str, settings: dict):
    sheet = Excel.get_sheet_by_name(workbook=workbook, sheet_name=sheet_name)
    header_rownum = Excel.find_row(sheet, search_col=settings["read"]["header"]["search_col"], search_str=settings["read"]["header"]["search_key"])

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
    result_rows = Utility.find_colnum_by_keywords(lst=header, keywords=settings["read"]["result_row"]["keys"], ignore_words=settings["read"]["result_row"]["ignores"])
    # 担当者
    person_rows = Utility.find_colnum_by_keywords(lst=header, keywords=settings["read"]["person_row"]["keys"])
    # 日付
    date_rows = Utility.find_colnum_by_keywords(lst=header, keywords=settings["read"]["date_row"]["keys"])

    # 結果,担当者,日付の列セットが見つからないor同数でない場合はエラー
    if Utility.check_lists_equal_length(result_rows, person_rows, date_rows) == False:
        return {
            "error": {
            "type": "set_is_incorrect",
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

        # 全セット合計のデータにも追加
        data.extend(set_data)

        # そのセットの1行目からセット名を取得(セル内改行は_に置換)
        set_name = Excel.get_cell_value(sheet=sheet, col=set[0], row=1, replace_newline=True)

        # セット名がない場合はスキップ
        if not set_name:
            continue

        # 環境名
        env_name = f"[{sheet_name}]{set_name}"

        # 環境ごとのデータ集計
        env_data[env_name] = get_daily(
            data=set_data, 
            results=settings["common"]["results"], 
            completed_label=settings["common"]["completed"], 
            completed_results=settings["common"]["completed_results"]
        )

    # 環境数
    env_count = len(sets)

    # テストケース数を計算
    tobe_rownunms = Utility.find_colnum_by_keywords(lst=header, keywords=settings["read"]["tobe_row"]["keys"])
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
    data_daily_total = get_daily(
        data=all_data,
        results=settings["common"]["results"],
        completed_label=settings["common"]["completed"],
        completed_results=settings["common"]["completed_results"]
    )

    # 全セット集計(担当者別)
    data_by_name = get_daily_by_name(all_data)
    
    # 全セット集計(全日付)
    data_total = get_total_all_date(data_daily_total, exclude=settings["common"]["completed"])

    # 総テストケース数
    case_count_all = sum(item['env_count'] * item['all'] for item in counts_by_sheet)
    # 対象外テストケース数
    excluded_count = get_excluded_count(data=all_data, targets=settings["read"]["excluded"])
    # 有効テストケース数
    available_count = case_count_all - excluded_count
    # 消化テストケース数
    filled_count = sum(data_total.values())
    # 完了テストケース数
    completed_count = sum_completed_results(data_total, settings["common"]["completed_results"])
    # 未実施テストケース数(マイナスは0)
    incompleted_count = max(0, available_count - filled_count)

    # 最終出力データ
    return {
        "count_total": {
            "all": case_count_all,
            "excluded": excluded_count,
            "available": available_count,
            "filled": filled_count,
            "completed": completed_count,
            "incompleted": incompleted_count
        },
        "count_by_sheet": counts_by_sheet,
        "daily": data_daily_total,
        "total": data_total,
        "by_name": data_by_name,
        "by_env": data_by_env
    }

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
