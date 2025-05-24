import csv
from tkinter import filedialog
from libs import Dialog

def save_2d_array(data: list, filename: str) -> None:
    """2次元配列をCSVファイルに保存する

    Args:
        data: 保存するデータ（2次元配列）
        filename: 保存するファイル名（デフォルト）
    """
    # 保存先の選択
    file_path = filedialog.asksaveasfilename(
        initialfile=filename,
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv")]
    )
    if not file_path:
        return
    
    # 保存
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerows(data)

    # 保存完了
    response = Dialog.ask_question(
        title="保存完了",
        message=f"CSVデータを保存しました。\n{file_path}\n\nファイルを開きますか？"
    )
    if response == "yes":
        from libs.FileOperation import run_file  # 循環参照を避けるため、ここでインポート
        run_file(file_path=file_path) 