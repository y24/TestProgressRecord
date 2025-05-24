import tkinter as tk

def copy_2d_array(data: list, root: tk.Tk) -> None:
    """クリップボードにデータをコピーする

    Args:
        data: コピーするデータ（2次元配列）
        root: Tkinterのルートウィンドウ
    """
    # タブ区切りのテキストに変換
    text = "\n".join(["\t".join(map(str, row)) for row in data])
    
    # クリップボードにコピー
    root.clipboard_clear()
    root.clipboard_append(text)
    root.update() 