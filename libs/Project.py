import json
import os
from libs import Utility

def save_to_json(file_path: str, aggregated_data: list, input_data: list, project_data: dict = None) -> None:
    """プロジェクトデータをJSONファイルに保存する

    Args:
        file_path (str): 保存先のJSONファイルパス
        input_data (list): 集計データ
        project_data (dict, optional): プロジェクト設定データ。
            指定されていない場合は、file_pathのJSONファイルから読み込む。
            JSONファイルが存在しない場合は空の辞書を使用。

    Raises:
        Exception: ファイルの保存に失敗した場合
    """
    try:
        # 既存のJSONファイルがある場合は読み込む
        existing_data = {}
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="utf-8") as f:
                existing_data = json.load(f)
                # project_dataが指定されていない場合は、既存のデータを使用
                if project_data is None:
                    project_data = existing_data.get("project", {})
        elif project_data is None:
            # ファイルが存在せず、project_dataも指定されていない場合は空の辞書を使用
            project_data = {}
        
        # projectキーにproject_dataを保存
        existing_data["project"] = project_data
        # 最終読込日時を保存（最も遅い日時を使用）
        existing_data["project"]["last_loaded"] = Utility.get_latest_time(input_data)
        # gathered_dataキーに現在のinput_dataを保存
        existing_data["gathered_data"] = input_data

        # JSONファイルに保存
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(existing_data, f, ensure_ascii=False, indent=2)
            
    except Exception as e:
        raise Exception(f"データの保存に失敗しました。\n{str(e)}") 