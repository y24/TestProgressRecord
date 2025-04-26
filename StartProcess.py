import sys, argparse, os, re
from tqdm import tqdm

import ReadData
import MainApp
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

def _remove_duplicate_number(filename: str) -> str:
    """
    ファイル名から末尾の ' (数字)' パターンを削除する
    
    Args:
        filename (str): 元のファイル名
        
    Returns:
        str: ' (数字)' パターンを削除したファイル名
    """
    # 末尾の ' (数字)' パターンを検索して削除
    return re.sub(r' \(\d+\)(?=\.[^.]+$)', '', filename)

def file_processor(file, settings, id):
    """
    個別のファイル処理
    """
    filename = Utility.get_filename_from_path(filepath=file["fullpath"])
    
    # データ集計
    result = ReadData.aggregate_results(filepath=file["fullpath"], settings=settings)
    
    # ファイル情報を付与
    result["file"] = _remove_duplicate_number(filename)
    result["filepath"] = file["fullpath"]
    result["relative_path"] = (
        Utility.get_relative_directory_path(full_path=file["fullpath"], base_dir=file["temp_dir"])
        if file["temp_dir"] else ""
    )
    result["selector_label"] = make_selector_label(result, id)
    
    return result

def validate_input_files(inputs):
    """
    入力ファイルの種類を検証し、適切な処理を行う
    
    Args:
        inputs (list): 入力ファイルのパスリスト
        
    Returns:
        list: 処理対象のファイルパスリスト
    """
    json_files = []
    excel_zip_files = []
    
    # ファイルを種類ごとに分類
    for file_path in inputs:
        ext = Utility.get_ext_from_path(file_path)
        if ext == "json":
            json_files.append(file_path)
        elif ext in ["xlsx", "zip"]:
            excel_zip_files.append(file_path)
    
    # 混在している場合の処理
    if json_files and excel_zip_files:
        Dialog.show_messagebox(
            root=None,
            type="info",
            title="ファイルが混在",
            message="プロジェクトファイル(*.json)とテスト仕様書ファイル(*.xlsx/*.zip)が混在しています。\nテスト仕様書ファイルが優先されます。"
        )
        return excel_zip_files
    
    # 複数のJSONファイルがある場合の処理
    if len(json_files) > 1:
        Dialog.show_messagebox(
            root=None,
            type="info",
            title="複数のプロジェクト",
            message="複数のプロジェクトファイルが選択されました。\n最初のファイルが優先されます。"
        )
        return [json_files[0]]
    
    return inputs

def start():
    # コマンドライン引数の設定
    parser = argparse.ArgumentParser(description="zipファイル/xlsxファイルを引数として起動します。(複数可)")
    parser.add_argument("--debug", action="store_true", help="デバッグモードを有効化")
    parser.add_argument("data_files", nargs="*", help="zipファイル/xlsxファイルのパス")
    args = parser.parse_args()

    # コマンドライン引数がない場合はファイル選択ダイアログを表示
    inputs = args.data_files or Dialog.select_files(("JSON/Excel/Zipファイル", "*.json;*.xlsx;*.zip"))
    if not inputs:
        sys.exit()

    # 入力ファイルの検証
    inputs = validate_input_files(inputs)

    # ファイルの拡張子を取得
    ext = Utility.get_ext_from_path(inputs[0])
    if ext == "json":
        # JSONファイルの場合はプロジェクトを開く
        print(inputs[0])
    else:
        # xlsx/zipファイルの場合はデータ集計
        files, temp_dirs = get_xlsx_paths(inputs)

    settings = AppConfig.load_settings()

    # 全ファイルの集計処理
    aggregate_data = [file_processor(file, settings, i+1) for i, file in enumerate(tqdm(files))]

    # デバッグモード時は処理結果を表示
    if args.debug:
        from pprint import pprint
        pprint(aggregate_data)

    # アプリケーションの起動
    MainApp.run(aggregate_data, inputs)

    # 一時ディレクトリの掃除
    if temp_dirs: Zip.cleanup_old_temp_dirs()

if __name__ == "__main__":
    start()