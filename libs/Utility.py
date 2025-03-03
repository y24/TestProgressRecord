import os
from pathlib import Path
from collections import OrderedDict

def find_rownum_by_keyword(list, keyword, ignore_words=None):
    if ignore_words is None:
        ignore_words = []
    return [i + 1 for i, item in enumerate(list) if item and keyword in str(item) and str(item) not in ignore_words]


def transpose_lists(*lists):
    return [list(row) for row in zip(*lists)]


def check_lists_equal_length(*lists):
    """
    任意の数のリストを受け取り、それらの要素数がすべて同じかどうかを判定する。
    
    :param lists: 可変長引数としてリストを受け取る
    :return: すべてのリストの長さが同じ場合は True、異なる場合は False
    """
    if not lists:
        return True  # 空の入力の場合は True を返す
    
    first_length = len(lists[0])
    return all(len(lst) == first_length for lst in lists)


def get_ext_from_path(filepath:str):
    return os.path.splitext(filepath)[1][1:]


def get_filename_from_path(filepath:str):
    return os.path.basename(filepath)


def is_empty(obj):
    """
    再帰的にオブジェクトが空かどうかを判定する関数。
    
    - 空の辞書、リスト、ネストされた空の辞書・リストは "空" と判定
    - 文字列、数値、None などは "空でない" と判定
    
    :param obj: 判定対象のオブジェクト
    :return: 空なら True、そうでなければ False
    """
    if isinstance(obj, dict):
        # 辞書の場合、すべてのキーが空と判定されたら空
        return all(is_empty(value) for value in obj.values())
    
    elif isinstance(obj, list):
        # リストの場合、すべての要素が空と判定されたら空
        return all(is_empty(item) for item in obj)
    
    # 空の辞書やリストを除いた場合、その他のオブジェクトは空ではない
    return False


def get_relative_path(fullpath: str, base_dir: str) -> str:
    """
    指定した基準ディレクトリ以降のパスを取得する関数
    
    Args:
        full_path (str): フルパス
        base_dir (str): 基準ディレクトリ
    
    Returns:
        str: 基準ディレクトリ以降の相対パス
    """
    fullpath = Path(fullpath).resolve()
    base_dir = Path(base_dir).resolve()
    
    if base_dir not in fullpath.parents:
        raise ValueError(f"指定されたパス '{fullpath}' は、基準ディレクトリ '{base_dir}' に含まれていません。")
    
    return str(fullpath.relative_to(base_dir))

def sort_nested_dates_desc(data):
    """
    一番上の階層のキーはそのままで、その下のキー（日付）を降順でソートする関数
    """
    sorted_data = {}
    for env, dates in data.items():
        sorted_data[env] = OrderedDict(sorted(dates.items(), key=lambda x: x[0], reverse=True))
    return sorted_data