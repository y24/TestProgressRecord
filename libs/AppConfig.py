import os, json, codecs

JSON_NAME = "AppConfig.json"

# 設定の保存
def save_settings(settings, json_name:str=JSON_NAME):
    with codecs.open(json_name, "w", "utf-8") as f:
        json.dump(settings, f, indent=4, ensure_ascii=False)

# 設定の読み込み
def load_settings(json_name: str = JSON_NAME):
    if os.path.exists(json_name):
        with open(json_name, "r", encoding="utf-8") as f:
            try:
                data = json.load(f)
                return data if data else None
            except json.JSONDecodeError:
                return None
    return None
