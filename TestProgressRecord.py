from libs import OpenpyxlWrapper as Excel
from libs import DataSummarize

import os, sys, pprint
from libs import Dialog
from libs import Logger
from libs import Utility
logger = Logger.get_logger(__name__, console=True, file=False, trace_line=False)

# 入力
# batファイルからファイルパスを受け取る(複数可)
inputs = sys.argv.copy()
print(inputs)

# 引数がない場合はファイル選択ダイアログ(複数可)
if len(inputs) <= 1:
    inputs = Dialog.select_files(("Excelファイル", "*.xlsx"))
else:
    del inputs[0]


# 処理
for filepath in inputs:
    filename = os.path.basename(filepath)
    logger.info(f"FILE: {filename}")
    logger.info(" ")

    # 対象シート名を取得
    workbook = Excel.load(filepath)
    sheet_names = Excel.get_sheetnames_by_keyword(workbook, "テスト項目")

    # シートがない場合はスキップ
    if len(sheet_names) == 0:
        logger.error(f"Sheet not found. ({filepath})")
        continue

    # 各シート処理
    all_data = []
    all_set_data = {}
    for sheet_name in sheet_names:
        # シート取得
        sheet = Excel.get_sheet_by_name(workbook=workbook, sheet_name=sheet_name)

        # ヘッダ行を探して取得
        row_num = Excel.find_row(sheet, search_row="A", search_str="#")
        header = Excel.get_row_values(sheet=sheet, row_num=row_num)

        # 列セット(例:環境別)を取得
        # 結果
        result_rows = Utility.find_rownum_by_keyword(list=header, keyword="結果")
        # 担当者
        person_rows = Utility.find_rownum_by_keyword(list=header, keyword="担当者")
        # 日付
        day_rows = Utility.find_rownum_by_keyword(list=header, keyword="日付")

        # セットが正しく取得できない場合はスキップ
        if Utility.check_lists_equal_length(result_rows, person_rows, day_rows) == False:
            logger.error(f"Failed to get result column. (Sheet: {sheet_name})")
            continue

        # 行番号のセット(結果、担当者、日付)を作成
        sets = Utility.transpose_lists(result_rows, person_rows, day_rows)
        
        # 各セット処理
        for set in sets:
            # セットのデータ取得
            set_data = Excel.get_columns_data(sheet=sheet, col_nums=set, header_row=row_num, ignore_header=True)

            # 全セット合計のデータにも追加
            all_data = all_data + set_data

            # シートからセット名を取得
            set_name = Excel.get_cell_value(sheet=sheet, col=set[0], row=1)
            # セット名がない場合はスキップ
            if not set_name:
                logger.error(f"Failed to get environment name. (Sheet: {sheet_name})")
                continue

            # セットごとのデータ集計
            all_set_data[set_name] = DataSummarize.get_daily(data=set_data, filter=["Pass", "Fixed"])

    # 全セット集計(日付別)
    summarized_data = DataSummarize.get_daily(data=all_data, filter=["Pass", "Fixed"])

    # 全セット集計(担当者別)
    summarized_data_by_name = DataSummarize.get_daily_by_name(data=all_data)

    # 出力
    dep = "  "
    # 全セット
    if any(summarized_data):
        logger.info("[ALL]")
        for key, d in summarized_data.items():
            logger.info(dep + f"{key}: {d}")
        logger.info(" ")

        logger.info("[By name]")
        for key, d in summarized_data_by_name.items():
            logger.info(dep + f"{key}: {d}")
        logger.info(" ")

    # 出力(セット別)
    if any(all_set_data):
        logger.info("[By environment]")
        for key, d in all_set_data.items():
            logger.info(dep + f"{key}:")
            for key, d in d.items():
                logger.info(dep * 2 + f"{key}: {d}")
        logger.info(" ")

    logger.info(" ")
    logger.info("~" * 50)

logger.info("Done.")