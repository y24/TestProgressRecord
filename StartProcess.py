import sys
import argparse
from tqdm import tqdm
import os

import ReadData
import App
from libs import Utility, Dialog, Zip, AppConfig

def get_xlsx_paths(inputs):
    """
    入力パスからxlsxファイルのパスを取得する
    - xlsxファイル: そのまま処理
    - zipファイル: 展開してxlsxファイルを抽出
    - ディレクトリ: 再帰的にxlsxとzipファイルを検索
    """
    files = []
    temp_dirs = []
    
    def process_directory(dir_path):
        """ディレクトリ内のxlsxとzipファイルを再帰的に処理"""
        for entry in os.scandir(dir_path):
            if entry.is_file():
                ext = Utility.get_ext_from_path(entry.path)
                if ext == "xlsx":
                    files.append({"fullpath": entry.path, "temp_dir": ""})
                elif ext == "zip":
                    extracted_files, temp_dir = Zip.extract_files_from_zip(entry.path, extensions=['.xlsx'])
                    files.extend([{"fullpath": f, "temp_dir": temp_dir} for f in extracted_files])
                    temp_dirs.append(temp_dir)
            elif entry.is_dir():
                process_directory(entry.path)
    
    for input_path in inputs:
        if os.path.isdir(input_path):
            # ディレクトリの場合
            process_directory(input_path)
        else:
            # ファイルの場合
            ext = Utility.get_ext_from_path(input_path)
            if ext == "xlsx":
                files.append({"fullpath": input_path, "temp_dir": ""})
            elif ext == "zip":
                extracted_files, temp_dir = Zip.extract_files_from_zip(input_path, extensions=['.xlsx'])
                files.extend([{"fullpath": f, "temp_dir": temp_dir} for f in extracted_files])
                temp_dirs.append(temp_dir)
            
    return files, temp_dirs


def make_selector_label(file, id):
    """ファイル選択用のラベルを生成する"""
    file_name = file["file"]
    relative_path = f'[{file["relative_path"]}] ' if file["relative_path"] else ""
    return f"{id}: {relative_path}{file_name}"


def file_processor(file, settings, id):
    """
    個別のファイル処理
    """
    filename = Utility.get_filename_from_path(filepath=file["fullpath"])
    
    # データ集計
    result = ReadData.aggregate_results(filepath=file["fullpath"], settings=settings)
    
    # ファイル情報を付与
    result["file"] = filename
    result["filepath"] = file["fullpath"]
    result["relative_path"] = (
        Utility.get_relative_directory_path(full_path=file["fullpath"], base_dir=file["temp_dir"])
        if file["temp_dir"] else ""
    )
    result["selector_label"] = make_selector_label(result, id)
    
    return result


def start():
    # コマンドライン引数の設定
    parser = argparse.ArgumentParser(description="zipファイル/xlsxファイルを引数として起動します。(複数可)")
    parser.add_argument("--debug", action="store_true", help="デバッグモードを有効化")
    parser.add_argument("data_files", nargs="*", help="zipファイル/xlsxファイルのパス")
    args = parser.parse_args()

    # コマンドライン引数がない場合はファイル選択ダイアログを表示
    inputs = args.data_files or Dialog.select_files(("Excel/Zipファイル", "*.xlsx;*.zip"))
    if not inputs:
        sys.exit()

    # xlsxファイルのパスを取得
    files, temp_dirs = get_xlsx_paths(inputs)
    settings = AppConfig.load_settings()

    # 全ファイルの処理
    out_data = [file_processor(file, settings, i+1) for i, file in enumerate(tqdm(files))]

    # デバッグモード時は処理結果を表示
    if args.debug:
        from pprint import pprint
        pprint(out_data)

    # アプリケーションの起動
    App.launch(out_data, inputs)

    # 一時ディレクトリの掃除
    if temp_dirs: Zip.cleanup_old_temp_dirs()


if __name__ == "__main__":
    start()