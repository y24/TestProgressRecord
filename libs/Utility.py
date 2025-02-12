def find_rownum_by_keyword(list, keyword):
    return [i + 1 for i, item in enumerate(list) if item and keyword in str(item)]


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