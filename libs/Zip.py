import zipfile
import tempfile
import os

TEMP_DIR_PREFIX = "_TEMP_"  # 一時フォルダの識別用プレフィックス

def extract_files_from_zip(zip_path, extensions=None):
     """
     指定したZIPファイルを解凍し、一時フォルダに配置する。
     指定した拡張子のファイルのフルパスをリストで返却する。
 
     :param zip_path: ZIPファイルのパス
     :param extensions: フィルタする拡張子のリスト (例: ['.xlsx', '.csv'])。None の場合はすべてのファイルを取得。
     :return: (ファイルのフルパスのリスト, 一時フォルダのパス)
     """
     if not zipfile.is_zipfile(zip_path):
         raise ValueError("指定されたファイルはZIPファイルではありません")
 
     # 一時フォルダを作成（識別しやすい名前）
     temp_dir = tempfile.mkdtemp(prefix=TEMP_DIR_PREFIX)
 
     with zipfile.ZipFile(zip_path, 'r') as zip_ref:
         zip_ref.extractall(temp_dir)  # ZIPを解凍
 
     # 解凍されたフォルダ内のファイルを取得
     all_files = [
         os.path.join(root, file)
         for root, _, files in os.walk(temp_dir)
         for file in files
     ]
 
     # 拡張子でフィルタリング（拡張子が指定されている場合）
     if extensions:
         extensions = {ext.lower() for ext in extensions}  # 大文字小文字を統一
         filtered_files = [file for file in all_files if os.path.splitext(file)[1].lower() in extensions]
     else:
         filtered_files = all_files  # すべてのファイルを取得
 
     return filtered_files, temp_dir

