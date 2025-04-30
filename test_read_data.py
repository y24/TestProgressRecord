import pytest
from ReadData import (
    get_daily,
    get_daily_by_name,
    get_excluded_count,
    get_total_all_date,
    sum_completed_results,
    make_run_status
)

def test_get_daily():
    # テストデータの準備
    test_data = [
        ["OK", "User1", "2024-01-01"],
        ["NG", "User2", "2024-01-01"],
        ["OK", "User1", "2024-01-02"],
        ["Pending", "User2", "2024-01-02"],
        ["OK", "User1", None]  # 日付なしのケース
    ]
    
    results = ["OK", "NG", "Pending"]
    completed_label = "Completed"
    completed_results = ["OK"]
    executed_label = "Executed"
    executed_results = ["OK", "NG", "Pending"]

    # 関数を実行
    daily_data, no_date_data = get_daily(
        test_data,
        results,
        completed_label,
        completed_results,
        executed_label,
        executed_results
    )

    # アサーション
    assert "2024-01-01" in daily_data
    assert daily_data["2024-01-01"]["OK"] == 1
    assert daily_data["2024-01-01"]["NG"] == 1
    assert daily_data["2024-01-01"]["Completed"] == 1
    assert daily_data["2024-01-01"]["Executed"] == 2

    assert "2024-01-02" in daily_data
    assert daily_data["2024-01-02"]["OK"] == 1
    assert daily_data["2024-01-02"]["Pending"] == 1

    assert "no_date" in no_date_data
    assert no_date_data["no_date"]["OK"] == 1
    assert no_date_data["no_date"]["Completed"] == 1
    assert no_date_data["no_date"]["Executed"] == 1

def test_get_daily_by_name():
    # テストデータの準備
    test_data = [
        ["OK", "User1", "2024-01-01"],
        ["NG", "User2", "2024-01-01"],
        ["OK", "User1", "2024-01-02"],
        ["Pending", "User2", "2024-01-02"]
    ]

    # 関数を実行
    result = get_daily_by_name(test_data)

    # アサーション
    assert "2024-01-01" in result
    assert result["2024-01-01"]["User1"] == 1
    assert result["2024-01-01"]["User2"] == 1
    assert result["2024-01-02"]["User1"] == 1
    assert result["2024-01-02"]["User2"] == 1

def test_get_excluded_count():
    # テストデータの準備
    test_data = [
        ["Excluded1", "User1", "2024-01-01"],
        ["Excluded2", "User2", "2024-01-01"],
        ["OK", "User1", "2024-01-02"]
    ]
    targets = ["Excluded1", "Excluded2"]

    # 関数を実行
    result = get_excluded_count(test_data, targets)

    # アサーション
    assert result == 2

def test_sum_completed_results():
    # テストデータの準備
    test_data = {
        "OK": 5,
        "NG": 3,
        "Pending": 2
    }
    completed_results = ["OK"]

    # 関数を実行
    result = sum_completed_results(test_data, completed_results)

    # アサーション
    assert result == 5

def test_make_run_status():
    # テストデータの準備
    count_stats = {
        "executed": 0,
        "completed": 0,
        "available": 10,
        "incompleted": 10
    }
    settings = {
        "app": {
            "state": {
                "not_started": {"name": "未着手"},
                "completed": {"name": "完了"},
                "in_progress": {"name": "進行中"}
            }
        }
    }

    # 未着手のケース
    result = make_run_status(count_stats, settings)
    assert result == "未着手"

    # 完了のケース
    count_stats["executed"] = 10
    count_stats["completed"] = 10
    count_stats["incompleted"] = 0
    result = make_run_status(count_stats, settings)
    assert result == "完了"

    # 進行中のケース
    count_stats["executed"] = 5
    count_stats["completed"] = 3
    count_stats["incompleted"] = 5
    result = make_run_status(count_stats, settings)
    assert result == "進行中" 