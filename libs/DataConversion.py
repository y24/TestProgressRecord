from libs import Utility

base_header = ["ファイル名", "フォルダ", "環境名", "日付"]

def convert_to_2d_array(data, settings):
    # ヘッダーの作成
    completed_label = settings["test_status"]["labels"]["completed"]
    executed_label = settings["test_status"]["labels"]["executed"]
    results = settings["test_status"]["results"]
    out_results = results + [executed_label, completed_label]
    header = base_header + out_results

    # 出力用の2次元配列の作成
    out_arr = [header]

    # データの書き込み
    for entry in data:
        file_name = entry.get("file", "")
        by_env_data = entry.get("by_env", {})
        daily_data = entry.get("daily", {})
        if not Utility.is_empty_recursive(by_env_data):
            # 環境別データがある場合
            for env, env_data in by_env_data.items():
                for date, values in env_data.items():
                    out_arr.append([file_name, entry["relative_path"], env, date] + [values.get(v, 0) for v in out_results])
        elif not Utility.is_empty_recursive(daily_data):
            # 環境別データがないが日付別データがある場合は、環境名は空で出力
            for date, values in entry.get("daily", {}).items():
                out_arr.append([file_name, entry["relative_path"], "", date] + [values.get(v, 0) for v in out_results])
        else:
            # 環境別データも日付別データもない場合は、環境名と日付を空で合計データを出力
            total_data = entry.get("total", {})
            stats_data = entry.get("stats", {})
            out_arr.append([file_name, entry["relative_path"], "", ""] + [total_data.get(v, 0) for v in results] + [stats_data.get("executed", 0), stats_data.get("completed", 0)])
    return out_arr