from collections import defaultdict

def get_daily_by_name(data):
    date_name_count = defaultdict(lambda: defaultdict(int))

    # 日付が空の行は削除
    data = [row for row in data if len(row) > 2 and row[2] not in ("", None)]

    # 結果が空ではない行を日付および名前ごとにカウント
    for row in data:
        result, name, date = row
        if result:  # 結果が空ではない場合
            date_name_count[date][name] += 1

    # 集計結果を返却
    result = {}
    for date, name_counts in sorted(date_name_count.items()):
        daily_count = {}
        for name, count in sorted(name_counts.items()):
            daily_count[name] = count
        result[date] = daily_count
    return result

def get_daily(data, filter: list[str]):
    total_label = "Completed"
    result_count = defaultdict(lambda: defaultdict(int))
    
    for row in data:
        result, name, date = row
        
        # 日付が空のデータはスキップ
        if not date:
            continue

        # 各結果を0で初期化
        for keyword in filter:
            result_count[date][keyword] = result_count[date].get(keyword, 0)
        result_count[date][total_label] = result_count[date].get(total_label, 0)

        # 結果列がフィルタ文字列に合致するものだけ抽出
        if result in filter:
            result_count[date][result] += 1
            result_count[date][total_label] += 1
    
    # 出力
    aggregated_results = {}
    for date, counts in sorted(result_count.items()):
        counts = {**counts}
        aggregated_results[date] = counts
    
    return aggregated_results