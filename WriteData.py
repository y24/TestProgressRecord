from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

def update_table(data, file_path, sheet_name):
    try:
        # 既存のブックを開く
        wb = load_workbook(file_path)
    except FileNotFoundError:
        # ファイルが存在しない場合、新規作成
        wb = Workbook()
    except PermissionError:
        print(f"Error: The file '{file_path}' is open in another program. Please close it and try again.")
        return
    
    # 既存のシートを削除
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    
    # 新しいシートを作成
    ws = wb.create_sheet(title=sheet_name)
    
    # データを書き込む
    for row_idx, row_data in enumerate(data, start=1):
        for col_idx, value in enumerate(row_data, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    # テーブル範囲を設定（A1から最終データセル）
    max_row = len(data)
    max_col = len(data[0]) if data else 0
    table_ref = f"A1:{chr(64 + max_col)}{max_row}"  # 例: A1:D10
    
    # テーブルの作成
    table = Table(displayName=sheet_name, ref=table_ref)
    style = TableStyleInfo(name="TableStyleLight9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    table.tableStyleInfo = style
    ws.add_table(table)
    
    try:
        # ブックを保存
        wb.save(file_path)
        print(f"Updated table in sheet '{sheet_name}' and saved to {file_path}")
    except PermissionError:
        raise
