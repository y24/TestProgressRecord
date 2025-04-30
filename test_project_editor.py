import pytest
import tkinter as tk
from ProjectEditorApp import ProjectEditorApp
import json
import os
from pathlib import Path
from unittest.mock import patch

@pytest.fixture
def root_window():
    """テスト用のルートウィンドウを作成"""
    root = tk.Tk()
    root.withdraw()  # ウィンドウを非表示に
    yield root
    root.destroy()

@pytest.fixture
def project_editor(root_window):
    """テスト用のProjectEditorAppインスタンスを作成"""
    editor = ProjectEditorApp(parent=root_window)
    yield editor
    editor.root.destroy()

def test_initial_state(project_editor):
    """初期状態のテスト"""
    assert project_editor.project_name_entry.get() == ""
    assert project_editor.write_path_entry.get() == ""
    assert project_editor.data_sheet_entry.get() == "DATA"
    assert len(project_editor.files_frame.winfo_children()) == 1  # Addボタンのみ

def test_sanitize_filename(project_editor):
    """ファイル名のサニタイズテスト"""
    # 正常なファイル名
    assert project_editor.sanitize_filename("test_project") == "test_project"
    
    # 特殊文字を含むファイル名
    assert project_editor.sanitize_filename("test/project") == "test_project"
    assert project_editor.sanitize_filename("test:project") == "test_project"
    assert project_editor.sanitize_filename("test*project") == "test_project"

@patch('tkinter.messagebox.showinfo')
@patch('tkinter.messagebox.showerror')
def test_save_project(mock_error, mock_info, tmp_path, project_editor):
    """プロジェクトの保存テスト"""
    # テストデータの設定
    project_editor.project_name_entry.insert(0, "Test Project")
    project_editor.write_path_entry.insert(0, str(tmp_path / "output.xlsx"))
    project_editor.data_sheet_entry.delete(0, tk.END)
    project_editor.data_sheet_entry.insert(0, "TEST_DATA")
    
    # ファイル情報を追加
    project_editor.add_file_info({
        "type": "local",
        "identifier": "test_file",
        "path": str(tmp_path / "test.xlsx")
    })
    
    # プロジェクトを保存
    project_editor.project_path = str(tmp_path / "test_project.json")
    project_editor.save_project()
    
    # メッセージボックスが呼び出されたことを確認
    mock_info.assert_called_once()
    mock_error.assert_not_called()
    
    # 保存されたファイルの内容を確認
    assert os.path.exists(project_editor.project_path)
    with open(project_editor.project_path, "r", encoding="utf-8") as f:
        data = json.load(f)
        project_data = data["project"]
        assert project_data["project_name"] == "Test Project"
        assert project_data["write_path"] == str(tmp_path / "output.xlsx")
        assert project_data["data_sheet_name"] == "TEST_DATA"
        assert len(project_data["files"]) == 1
        assert project_data["files"][0]["identifier"] == "test_file"

@patch('tkinter.messagebox.showerror')
def test_load_project(mock_error, tmp_path, project_editor):
    """プロジェクトの読み込みテスト"""
    # テスト用のプロジェクトファイルを作成
    test_project = {
        "project": {
            "project_name": "Loaded Project",
            "files": [{
                "type": "local",
                "identifier": "loaded_file",
                "path": str(tmp_path / "loaded.xlsx")
            }],
            "write_path": str(tmp_path / "loaded_output.xlsx"),
            "data_sheet_name": "LOADED_DATA"
        }
    }
    
    project_path = tmp_path / "test_load.json"
    with open(project_path, "w", encoding="utf-8") as f:
        json.dump(test_project, f, ensure_ascii=False, indent=2)
    
    # プロジェクトを読み込み
    project_editor.load_project_from_path(str(project_path))
    
    # エラーメッセージが表示されていないことを確認
    mock_error.assert_not_called()
    
    # 読み込まれた内容を確認
    assert project_editor.project_name_entry.get() == "Loaded Project"
    assert project_editor.write_path_entry.get() == str(tmp_path / "loaded_output.xlsx")
    assert project_editor.data_sheet_entry.get() == "LOADED_DATA"

def test_get_file_info(project_editor):
    """ファイル情報の取得テスト"""
    # ファイル情報を追加
    project_editor.add_file_info({
        "type": "local",
        "identifier": "test_file",
        "path": "test.xlsx"
    })
    
    # ファイル情報を取得
    files = project_editor.get_file_info()
    
    # 取得した情報を確認
    assert len(files) == 1
    assert files[0]["type"] == "local"
    assert files[0]["identifier"] == "test_file"
    assert files[0]["path"] == "test.xlsx" 