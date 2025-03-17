import os, json, codecs

USER_JSON_NAME = "UserConfig.json"
DEFAULT_JSON_NAME = "DefaultConfig.json"

# 設定の保存
def save_settings(settings, json_name:str=USER_JSON_NAME):
    with codecs.open(json_name, "w", "utf-8") as f:
        json.dump(settings, f, indent=4, ensure_ascii=False)


def merge_missing_keys(default_data, user_data):
    if not isinstance(default_data, dict) or not isinstance(user_data, dict):
        return user_data

    for key, value in default_data.items():
        if key not in user_data:
            user_data[key] = value
        elif isinstance(value, dict):
            user_data[key] = merge_missing_keys(value, user_data.get(key, {}))
    return user_data

# 設定の読込
def load_settings(user_config_path=USER_JSON_NAME, default_config_path=DEFAULT_JSON_NAME):
    """
    DefaultConfig.jsonを元にUserConfig.jsonに不足しているキーを補完する。
    UserConfig.jsonがなければ作成する。

    :param user_config_path: UserConfig.jsonのパス
    :param default_config_path: DefaultConfig.jsonのパス
    """
    # DefaultConfig.json の読み込み
    if not os.path.exists(default_config_path):
        raise FileNotFoundError(f"{default_config_path} が見つかりません。")

    with open(default_config_path, "r", encoding="utf-8") as f:
        default_data = json.load(f)

    # UserConfig.json の読み込み（なければ空の辞書）
    if os.path.exists(user_config_path):
        with open(user_config_path, "r", encoding="utf-8") as f:
            user_data = json.load(f)
    else:
        user_data = {}

    # 不足キーを補完
    updated_data = merge_missing_keys(default_data, user_data)

    # 更新後のUserConfig.jsonを書き込み
    with open(user_config_path, "w", encoding="utf-8") as f:
        json.dump(updated_data, f, indent=4, ensure_ascii=False)

    return updated_data