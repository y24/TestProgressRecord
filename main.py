import os, sys, pprint
from tqdm import tqdm

import ReadData
import App
from libs import Utility
from libs import Dialog
from libs import Zip
from libs import AppConfig

def main():
    # 設定読み込み
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
        result = ReadData.aggregate_results(filepath=file["fullpath"], settings=settings)
        # 出力
        if result and not Utility.is_empty(result):
            # ファイルパス
            result["file"] = Utility.get_filename_from_path(filepath=file["fullpath"])
            # zipファイル内の相対パスを取得
            if file["temp_dir"]:
                result["relative_path"] = Utility.get_relative_directory_path(full_path=file["fullpath"], base_dir=file["temp_dir"])
                result["selector_label"] = ReadData.make_selector_label(result)
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


if __name__ == "__main__":
    main()