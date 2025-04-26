import json
import os
import tempfile
import urllib.parse
import requests
from typing import List
from pathlib import Path

def is_sharepoint_url(url: str) -> bool:
    """URLがSharePointのURLかどうかを判定する"""
    parsed_url = urllib.parse.urlparse(url)
    return "sharepoint.com" in parsed_url.netloc

def download_sharepoint_file(url: str, temp_dir: str) -> str:
    """SharePointのファイルをダウンロードする"""
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()
        
        # URLからファイル名を取得
        file_name = os.path.basename(urllib.parse.urlparse(url).path)
        if not file_name:
            file_name = "downloaded_file"
            
        # 一時フォルダにファイルを保存
        file_path = os.path.join(temp_dir, file_name)
        with open(file_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    
        return file_path
    except Exception as e:
        raise Exception(f"ファイルのダウンロードに失敗しました: {str(e)}")

def process_json_file(json_path: str) -> List[str]:
    """
    JSONファイルからURLを取得し、ファイルをダウンロードする
    
    Args:
        json_path (str): JSONファイルのパス
        
    Returns:
        List[str]: ダウンロードされたファイルのパスのリスト
    """
    try:
        # JSONファイルを読み込む
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            
        # projectオブジェクトのfilesリストからURLを取得
        urls = []
        if 'project' in data and 'files' in data['project']:
            urls = [file.get('url') for file in data['project']['files'] if file.get('url')]
            
        # 一時フォルダを作成
        with tempfile.TemporaryDirectory(prefix="_TEMP_") as temp_dir:
            downloaded_files = []
            
            # 各URLに対して処理を実行
            for url in urls:
                if is_sharepoint_url(url):
                    file_path = download_sharepoint_file(url, temp_dir)
                    downloaded_files.append(file_path)
                    
            return downloaded_files, temp_dir
            
    except Exception as e:
        raise Exception(f"JSONファイルの処理中にエラーが発生しました: {str(e)}")
