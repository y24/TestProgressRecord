import zipfile
import tempfile
import os
import shutil
import glob

TEMP_DIR_PREFIX = "_TEMP_"  # 一時フォルダの識別用プレフィックス

def cleanup_old_temp_dirs():
    """
    スクリプトの起動時に、過去の一時フォルダを削除する。
    """
    temp_base_dir = tempfile.gettempdir()
    temp_dirs = glob.glob(os.path.join(temp_base_dir, TEMP_DIR_PREFIX + "*"))

    for temp_dir in temp_dirs:
        if os.path.isdir(temp_dir):
            shutil.rmtree(temp_dir)

def extract_files_from_zip(zip_path, extensions=None):
    """
    指定したZIPファイルを解凍し、日本語のファイル名の文字化けを修正。
    指定した拡張子のファイルのフルパスをリストで返却する。

    :param zip_path: ZIPファイルのパス
    :param extensions: フィルタする拡張子のリスト (例: ['.xlsx', '.csv'])。None の場合はすべてのファイルを取得。
    :return: (ファイルのフルパスのリスト, 一時フォルダのパス)
    """
    if not zipfile.is_zipfile(zip_path):
        raise ValueError("指定されたファイルはZIPファイルではありません")

    # 一時フォルダを作成
    temp_dir = tempfile.mkdtemp(prefix=TEMP_DIR_PREFIX)

    extracted_files = []
    
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        # 文字化け修正 (CP437 -> Shift_JIS)
        for file in zip_ref.namelist():
            try:
                correct_name = file.encode('cp437').decode('shift_jis')
            except UnicodeDecodeError:
                correct_name = file  # 失敗した場合はそのまま
                
            extracted_path = os.path.join(temp_dir, correct_name)

            # ディレクトリの場合
            if file.endswith('/'):
                os.makedirs(extracted_path, exist_ok=True)
            else:
                os.makedirs(os.path.dirname(extracted_path), exist_ok=True)
                with open(extracted_path, 'wb') as f:
                    f.write(zip_ref.read(file))
                extracted_files.append(extracted_path)

    # 拡張子でフィルタリング
    if extensions:
        extensions = {ext.lower() for ext in extensions}  # 大文字小文字を統一
        filtered_files = [file for file in extracted_files if os.path.splitext(file)[1].lower() in extensions]
    else:
        filtered_files = extracted_files  # すべてのファイルを取得

    return filtered_files, temp_dir

def cleanup_temp_dir(temp_dir):
    """
    指定された一時フォルダを削除する。

    :param temp_dir: 削除する一時フォルダのパス
    """
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
