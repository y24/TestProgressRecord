import sys, os, re, json
from tqdm import tqdm
from datetime import datetime

import ReadData
import MainApp
from libs import Utility, Dialog, Zip, AppConfig, TempDir, FileDownload

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
    # 最終読込日時を記録
    result["last_loaded"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # ファイルの更新日時を記録
    result["last_updated"] = datetime.fromtimestamp(os.path.getmtime(file["fullpath"])).strftime("%Y-%m-%d %H:%M:%S")
    # ファイルのソースを記録
    result["source"] = "sharepoint" if file.get("type") == "sharepoint" else "local"
    
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

def filter_xlsx_files(inputs):
    """
    xlsxファイル以外のパスを除去する
    
    Args:
        inputs (list): 入力ファイルのパスリスト
        
    Returns:
        list: xlsxファイルのみのパスリスト
    """
    return [file_path for file_path in inputs if Utility.get_ext_from_path(file_path) == "xlsx"]

def process_files(inputs, project_path=""):
    """
    ファイル処理のメイン関数
    
    Args:
        inputs (list): 入力ファイルのパスリスト
        project_path (str): プロジェクトファイルのパス（オプション）
    """
    # 設定ファイルの読み込み
    settings = AppConfig.load_settings()

    # 入力ファイルの検証
    inputs = validate_input_files(inputs)

    # プロジェクトデータを初期化
    project_data = {}

    # プロジェクトファイルのパスが指定されている場合はそのファイルを読み込む
    if project_path:
        with open(project_path, "r", encoding="utf-8") as f:
            json_data = json.load(f)
        project_data = json_data["project"]

    # ファイルの拡張子を取得
    ext = Utility.get_ext_from_path(inputs[0])
    if ext == "json":
        # プロジェクトファイル(jsonファイル)の場合
        try:
            with open(inputs[0], "r", encoding="utf-8") as f:
                json_data = json.load(f)

            # プロジェクトデータを取得
            if "project" in json_data:
                project_data = json_data["project"]
                project_path = inputs[0]

            # 記録済みのデータ（aggregate_data）がある場合はそのデータを使用
            if "aggregate_data" in json_data:
                aggregate_data = json_data["aggregate_data"]
            else:
                # ファイルの種類に応じて処理を分岐
                local_files = []
                sharepoint_files = []
                
                # プロジェクトデータからファイル情報を取得
                if "project" in json_data and "files" in json_data["project"]:
                    for file in json_data["project"]["files"]:
                        if file.get("type") == "local":
                            local_files.append(file.get("path"))
                        elif file.get("type") == "sharepoint":
                            sharepoint_files.append(file.get("path"))
                
                # localファイルの処理
                if local_files:
                    files, temp_dirs = get_xlsx_paths(local_files)
                    aggregate_data = [file_processor(file, settings, i+1) for i, file in enumerate(tqdm(files))]
                
                # sharepointファイルの処理
                if sharepoint_files:
                    files, temp_dirs = FileDownload.process_json_file(inputs[0])
                    # xlsxファイルのみフィルタ
                    files = filter_xlsx_files(files)
                    # 全ファイルの集計処理
                    if files:
                        aggregate_data.extend([file_processor(file, settings, i+1) for i, file in enumerate(tqdm(files))])

        except Exception as e:
            Dialog.show_messagebox(root=None, type="error", title="ファイル読込エラー", message=f"{str(e)}")
            # プロジェクトファイルの読込に失敗した場合はデータ0件とする
            aggregate_data = []
    else:
        # xlsx/zipファイルを指定した場合
        files, temp_dirs = get_xlsx_paths(inputs)
        # 全ファイルの集計処理
        aggregate_data = [file_processor(file, settings, i+1) for i, file in enumerate(tqdm(files))]

    # アプリケーションの起動
    MainApp.run(pjdata=project_data, pjpath=project_path, indata=aggregate_data, args=inputs)

    # 一時ディレクトリの掃除
    try:
        if temp_dirs:
            TempDir.cleanup_old_temp_dirs("_TEMP_")
    except NameError:
        pass

if __name__ == "__main__":
    # コマンドライン引数の設定
    import argparse
    parser = argparse.ArgumentParser(description="TestTraQ - ファイル処理")
    parser.add_argument("--project", help="プロジェクトファイルのパス")
    parser.add_argument("data_files", nargs="*", help="処理するファイルのパス")
    args = parser.parse_args()

    process_files(args.data_files, args.project)
