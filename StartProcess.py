import sys, pprint
from tqdm import tqdm

import ReadData
import App
from libs import Utility
from libs import Dialog
from libs import Zip
from libs import AppConfig

def get_xlsx_paths(inputs):
    files = []
    temp_dirs = []
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
    return files, temp_dirs

def file_processor(file, settings, id):
    # 集計
    result = ReadData.aggregate_results(filepath=file["fullpath"], settings=settings)
    # ファイル名
    filename = Utility.get_filename_from_path(filepath=file["fullpath"])
    # 出力
    if result and not Utility.is_empty(result):
        # ファイルパス
        result["file"] = filename
        # zipファイル内の相対パスを取得
        if file["temp_dir"]:
            result["relative_path"] = Utility.get_relative_directory_path(full_path=file["fullpath"], base_dir=file["temp_dir"])
        else:
            # zipファイルではない場合
            result["relative_path"] = ""
        # プルダウンメニュー表示用
        result["selector_label"] = ReadData.make_selector_label(result, id)
        # コンソール出力
        # console_out(result)
        # ビューアに渡す配列に格納
        return result
    else:
        # 出力データがない場合
        return { "error": filename }


def start():
    # 引数でファイルパスを受け取る(複数可)
    inputs = sys.argv.copy()

    # 引数がない場合はファイル選択ダイアログ(複数可)
    if len(inputs) <= 1:
        inputs = Dialog.select_files(("Excel/Zipファイル", "*.xlsx;*.zip"))
        # キャンセル時は終了
        if not inputs: sys.exit()
    else:
        del inputs[0]

    # xlsxファイルのパスを抽出
    files, temp_dirs = get_xlsx_paths(inputs=inputs)

    # 設定読み込み
    settings = AppConfig.load_settings()

    # ファイルを処理
    out_data = []
    errors = []
    for index, file in enumerate(tqdm(files)):
        id = index + 1
        res = file_processor(file=file, settings=settings, id=id)
        if not 'error' in res.keys():
            out_data.append(res)
        else:
            errors.append(res)

    # pprint.pprint(out_data)

    # ビューア起動
    App.launch(out_data, errors)

    # zipファイルを展開していた場合は一時フォルダを掃除
    if len(temp_dirs): Zip.cleanup_old_temp_dirs()


if __name__ == "__main__":
    start()