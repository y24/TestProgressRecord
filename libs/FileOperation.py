import os
import subprocess
from libs import Dialog

def run(file_path: str, exit: bool = False) -> None:
    """ファイルを実行/開く

    Args:
        file_path: 実行/開くファイルのパス
        exit: 実行後にプログラムを終了するかどうか
    """
    if os.path.isfile(file_path):
        try:
            os.startfile(file_path)  # Windows
        except AttributeError:
            subprocess.run(["xdg-open", file_path])  # Linux/Mac
        except Exception as e:
            Dialog.show_messagebox(type="error", title="Error", message=f"ファイルを開けませんでした。\n{e}")
        if exit:
            import sys
            sys.exit()
    else:
        Dialog.show_messagebox(type="error", title="Error", message="指定されたファイルが見つかりません") 