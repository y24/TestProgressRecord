from libs import OpenpyxlWrapper as Excel

def update_table(data, file_path, sheet_name):
    # ブック読み込み
    wb = Excel.load(file_path=file_path)
    
    # 既存のデータシートを削除して新規作成
    ws = Excel.create_sheet(workbook=wb, sheet_name=sheet_name, overwrite=True)
    
    # データを書き込んでテーブルを作成
    Excel.create_datatable(ws, sheet_name, data)
    
    # 列幅の調整
    Excel.adjust_colwidth_by_headername(sheet=ws, target_headers=["ファイル名", "環境名", "日付"], header_row=1)

    # 保存
    try:
        wb.save(file_path)
        return True
    except PermissionError:
        raise
