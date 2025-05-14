# テスト進捗管理システム データモデル仕様書

## 1. データモデル概要

### 1.1 基本構造
テスト進捗管理システムは、以下の主要なデータモデルで構成されています：

1. プロジェクト設定データ
2. テスト結果データ
3. 集計データ

## 2. プロジェクト設定データ

### 2.1 基本設定（DefaultConfig.json）
```json
{
    "read_definition": {
        "sheet_search_keys": ["テスト項目"],
        "sheet_search_ignores": [],
        "header": {"search_col": "A", "search_key": "#"},
        "tobe_row": {"keys": ["期待", "実施対象"]},
        "result_row": {"keys": ["結果"], "ignores": ["期待結果"]},
        "person_row": {"keys": ["担当者"]},
        "date_row": {"keys": ["日付", "実施日"]},
        "excluded": ["対象外"]
    },
    "test_status": {
        "results": ["Pass", "Fixed", "Fail", "Blocked", "Suspend", "N/A"],
        "completed_results": ["Pass", "Fixed", "Suspend", "N/A"],
        "executed_results": ["Pass", "Fixed", "Fail", "Blocked", "Suspend", "N/A"],
        "labels": {
            "completed": "完了数",
            "executed": "消化数",
            "not_run": "未着手"
        }
    }
}
```

### 2.2 設定項目の説明

#### 2.2.1 読み込み定義（read_definition）
- `sheet_search_keys`: 対象シートを特定するための検索キーワード
- `sheet_search_ignores`: 除外するシート名のキーワード
- `header`: ヘッダー行の検索条件
- `tobe_row`: 期待結果行の定義
- `result_row`: 結果行の定義
- `person_row`: 担当者行の定義
- `date_row`: 日付行の定義
- `excluded`: 除外対象のキーワード

#### 2.2.2 テストステータス（test_status）
- `results`: 定義されている全ての結果タイプ
- `completed_results`: 完了として扱う結果タイプ
- `executed_results`: 実行済みとして扱う結果タイプ
- `labels`: 表示用ラベル

## 3. テスト結果データ

### 3.1 基本構造
テスト結果データは以下の3つの主要なビューで管理されます：

1. 日付別ビュー
2. 環境別ビュー
3. 担当者別ビュー

### 3.2 データ形式

#### 3.2.1 日付別データ
```python
{
    "YYYY-MM-DD": {
        "Pass": 10,
        "Fail": 2,
        "Blocked": 1,
        "Fixed": 3,
        "Suspend": 1,
        "N/A": 0,
        "完了数": 15,
        "消化数": 17
    }
}
```

#### 3.2.2 環境別データ
```python
{
    "[シート名]環境名": {
        "YYYY-MM-DD": {
            "Pass": 5,
            "Fail": 1,
            "Blocked": 0,
            "Fixed": 2,
            "Suspend": 0,
            "N/A": 0,
            "完了数": 7,
            "消化数": 8
        }
    }
}
```

#### 3.2.3 担当者別データ
```python
{
    "YYYY-MM-DD": {
        "担当者名1": 10,
        "担当者名2": 5,
        "担当者名3": 2
    }
}
```

## 4. 集計データ

### 4.1 基本統計情報
```python
{
    "stats": {
        "all": 100,        # 総テストケース数
        "excluded": 5,     # 除外テストケース数
        "available": 95,   # 有効テストケース数
        "executed": 80,    # 消化テストケース数
        "completed": 70,   # 完了テストケース数
        "incompleted": 15  # 未実施テストケース数
    }
}
```

### 4.2 実行状況データ
```python
{
    "run": {
        "status": "進行中",     # 実施状況
        "start_date": "2024-01-01",  # 開始日
        "last_update": "2024-03-15"  # 最終更新日
    }
}
```

## 5. データフロー

### 5.1 データ読み込みフロー
1. Excelファイルからデータを読み込み
2. シートごとにデータを解析
3. 環境別にデータを集計
4. 日付別・担当者別にデータを集計
5. 全体の統計情報を計算

### 5.2 データ更新フロー
1. 新しいテスト結果データの読み込み
2. 既存データとのマージ
3. 集計データの再計算
4. 表示データの更新

## 6. データ検証

### 6.1 入力データ検証
- シートの存在確認
- ヘッダー行の検証
- 必須列の存在確認
- データ形式の検証

### 6.2 集計データ検証
- テストケース数の整合性確認
- 完了数と消化数の整合性確認
- 日付データの妥当性確認
- 環境別データの整合性確認 