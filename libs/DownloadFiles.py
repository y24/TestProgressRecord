import os
import tempfile
import requests
from urllib.parse import urlparse

TEMP_DIR_PREFIX = "_TEMP_"  # 一時フォルダの識別用プレフィックス

def download_files(sharepoint_files: list) -> tuple[list, str]:
    """
    SharePointのダウンロードURLからファイルをダウンロードし、一時フォルダに保存する

    Args:
        sharepoint_files (list): ダウンロードURLのリスト

    Returns:
        tuple[list, str]: (ダウンロードしたファイルのパスのリスト, 一時フォルダのパス)
    """
    # 一時フォルダを作成（識別しやすい名前）
    temp_dir = tempfile.mkdtemp(prefix=TEMP_DIR_PREFIX)
    downloaded_files = []

    for url in sharepoint_files:
        try:
            # ファイル名を取得（URLの最後の部分）
            parsed_url = urlparse(url)
            filename = os.path.basename(parsed_url.path)
            if not filename:
                # ファイル名が取得できない場合はスキップ
                continue

            # ファイルをダウンロード
            response = requests.get(url)
            response.raise_for_status()  # エラーチェック

            # 一時フォルダにファイルを保存
            file_path = os.path.join(temp_dir, filename)
            with open(file_path, 'wb') as f:
                f.write(response.content)

            temp_dirs = [temp_dir]

            downloaded_files.append(file_path)

        except Exception as e:
            print(f"Error downloading {url}: {str(e)}")
            continue

    return downloaded_files, temp_dirs 