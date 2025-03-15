import os
from pathlib import Path
from collections import OrderedDict
from collections import defaultdict

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


def get_relative_path(full_path: str, base_dir: str) -> str:
    """
    指定した基準ディレクトリ以降のパスを取得する関数
    
    Args:
        full_path (str): フルパス
        base_dir (str): 基準ディレクトリ
    
    Returns:
        str: 基準ディレクトリ以降の相対パス
    """
    full_path = Path(full_path).resolve()
    base_dir = Path(base_dir).resolve()
    
    if base_dir not in full_path.parents:
        raise ValueError(f"指定されたパス '{full_path}' は、基準ディレクトリ '{base_dir}' に含まれていません。")
    return str(full_path.relative_to(base_dir))


def get_relative_directory_path(full_path: str, base_dir: str) -> str:
    """
    指定したディレクトリ(base_dir) 以降のパスを取得する関数
    """
    # 絶対パスに変換して正規化
    full_path = os.path.abspath(full_path)
    base_dir = os.path.abspath(base_dir)

    # base_dir が file_path の先頭部分と一致するか確認
    if not full_path.startswith(base_dir):
        raise ValueError("指定したディレクトリ(base_dir)が、ファイルのパス内に見つかりません。")

    # base_dir の長さを取得し、それ以降のパスを取得
    relative_path = full_path[len(base_dir):].lstrip(os.sep)

    # ディレクトリ部分のみ取得
    return os.path.dirname(relative_path)

def sort_nested_dates_desc(data):
    """
    一番上の階層のキーはそのままで、その下のキー（日付）を降順でソートする関数
    """
    sorted_data = {}
    for env, dates in data.items():
        sorted_data[env] = OrderedDict(sorted(dates.items(), key=lambda x: x[0], reverse=True))
    return sorted_data

def sort_by_master(master_list, input_list):
    """
    指定されたマスタリストの順番に基づいて、入力リストを並び替える。
    """
    return sorted(input_list, key=lambda x: master_list.index(x) if x in master_list else float('inf'))

def meke_rate_text(value1, value2):
    # パーセンテージ表示用の文字列を生成
    if value2:
        rate = (value1 / value2) * 100
        # rate = rate if rate < 100 else 100
        return f"{rate:.1f}%"
    else:
        return "--"

def sum_values(list, param):
    # オブジェクト配列の各キーごとに合計値を計算
    result = defaultdict(int)
    for entry in list:
        for key, value in entry[param].items():
            result[key] += value
    return result