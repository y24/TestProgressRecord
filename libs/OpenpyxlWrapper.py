from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import column_index_from_string
from unicodedata import east_asian_width
from datetime import datetime

width_dict = {
  'F': 2,   # Fullwidth
  'H': 1,   # Halfwidth
  'W': 2,   # Wide
  'Na': 1,  # Narrow
  'A': 2,   # Ambiguous
  'N': 1    # Neutral
}

Font_depend = 1.2

def load(file_path:str, auto_create:bool=False):
    # ブックを開く
    try:
        wb = load_workbook(file_path)
    except FileNotFoundError:
        # ファイルが存在しない場合、新規作成
        if auto_create:
            wb = Workbook()
        else:
            raise FileNotFoundError(f"Error: {file_path} が見つかりません。")
    except PermissionError:
        raise PermissionError(f"Error: '{file_path}' は他のプログラムによって開かれています。")
    return wb

def create_sheet(workbook, sheet_name:str, overwrite:bool=False):
    # 既存のデータシートがあれば削除
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]
    # 新しいデータシートを作成
    return workbook.create_sheet(title=sheet_name)


def get_sheet_by_name(workbook, sheet_name:str):
    return workbook[sheet_name]


def get_sheetnames_by_keyword(workbook, keyword:str):
    return [sheet for sheet in workbook.sheetnames if keyword in sheet]

def get_sheetnames_by_keywords(workbook, keywords: list, ignores: list = None):
    if ignores is None:
        ignores = []
    return [
        sheet for sheet in workbook.sheetnames 
        if any(keyword in sheet for keyword in keywords) and not any(ignore in sheet for ignore in ignores)
    ]


def find_row(sheet, search_col:str, search_str:str):
    try:
        # 列名
        col_num = column_index_from_string(search_col) - 1

        # 指定列をループして値を確認
        for row in sheet.iter_rows(min_col=1, max_col=1):
            cell = row[col_num]  # 指定列のセル
            if cell.value == search_str:  # 値が search_str のセル
                return cell.row
        return None

    except Exception as e:
        print(f"Error: {e}")


def get_row_values(sheet, row_num:int):
    return [cell.value for cell in sheet[row_num]]


def get_column_values(sheet, col_nums: list, header_row: int = 1, ignore_header=False):
    if ignore_header:
        header_row += 1
    return [[sheet.cell(row=i, column=col_num).value for col_num in col_nums] for i in range(header_row, sheet.max_row + 1)]


def get_columns_data(sheet, col_nums: list, header_row: int = 1, ignore_header=False):
    if ignore_header:
        header_row += 1
    data = [[cell.value.strftime('%Y-%m-%d') if isinstance(cell.value, datetime) else cell.value 
             for col_num in col_nums 
             for cell in [sheet.cell(row=i, column=col_num)]] 
            for i in range(header_row, sheet.max_row + 1)]
    return data

def get_cell_value(sheet, col:int, row:int):
    return sheet.cell(row=row, column=col).value

def create_datatable(ws, sheet_name, data):
    # シートのA1からデータ書き込み
    for row_idx, row_data in enumerate(data, start=1):
        for col_idx, value in enumerate(row_data, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    # テーブル範囲の設定
    max_row = len(data)
    max_col = len(data[0]) if data else 0
    table_ref = f"A1:{chr(64 + max_col)}{max_row}"  # 例: A1:D10

    # テーブル作成
    table = Table(displayName=sheet_name, ref=table_ref)
    style = TableStyleInfo(name="TableStyleLight9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    table.tableStyleInfo = style
    ws.add_table(table)

# ヘッダ名に合致する列の幅を調整
def adjust_colwidth_by_headername(sheet, target_headers:list[str], header_row:int=1):
    for col in sheet.columns:
        # ヘッダ行の値（header_row=1：1行目をヘッダとする）
        header = col[header_row-1].value
        # target_headersに合致するヘッダがあれば調整
        if header in target_headers:
            max_length= 1
            max_diameter = 1
            column= col[1].column_letter
            for cell in col:
                diameter = (cell.font.size*Font_depend)/10
                if diameter > max_diameter:
                    max_diameter = diameter
                try:
                    if(cell.value == None) : continue
                    chars = [char for char in str(cell.value)]
                    east_asian_width_list = [east_asian_width(char) for char in chars]
                    width_list = [width_dict[east_asian_width] for east_asian_width in east_asian_width_list]
                    if sum(width_list) > max_length:
                        max_length= sum(width_list)
                except:
                    pass
                sheet.column_dimensions[column].width= max_length*max_diameter + 1.2