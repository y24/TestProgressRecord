import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import json
import os
from datetime import datetime
import numpy as np
from pathlib import Path
import matplotlib.pyplot as plt
import io
import base64
import subprocess
import sys
import time
import re

# 設定の読み込み
def load_settings():
    try:
        with open("UserConfig.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        with open("DefaultConfig.json", "r", encoding="utf-8") as f:
            return json.load(f)

# プロジェクトデータの読み込み
def load_project_data(project_path):
    try:
        with open(project_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        st.error(f"プロジェクトファイルの読み込みに失敗しました: {str(e)}")
        return None

# 進捗状況のグラフを作成
def create_progress_chart(data, settings):
    results = data.get("total", {})
    incompleted = data.get("stats", {}).get("incompleted", 0)

    # 結果ごとに値と色を準備
    bar_data = []
    for result in settings["test_status"]["results"]:
        value = results.get(result, 0)
        if value > 0:
            bar_data.append({
                "name": result,
                "value": value,
                "color": settings["app"]["bar"]["colors"].get(result, "gainsboro")
            })
    # 未着手
    if incompleted > 0:
        bar_data.append({
            "name": settings["test_status"]["labels"]["not_run"],
            "value": incompleted,
            "color": settings["app"]["bar"]["colors"].get(settings["test_status"]["labels"]["not_run"], "gainsboro")
        })

    # 積み上げバー作成
    fig = go.Figure()
    for bar in bar_data:
        fig.add_trace(go.Bar(
            x=[bar["value"]],
            y=["進捗状況"],
            orientation='h',
            name=bar["name"],
            marker_color=bar["color"],
            text=[bar["value"]],
            textposition='inside',
        ))

    fig.update_layout(
        barmode='stack',
        showlegend=True,
        height=100,
        margin=dict(l=0, r=0, t=30, b=0),
    )
    return fig

# 日付別データのテーブルを作成
def create_daily_table(data, settings):
    if not data.get("daily"):
        return pd.DataFrame()
    
    # データの準備
    daily_data = []
    for date, values in data["daily"].items():
        row = {"日付": date}
        row.update(values)
        daily_data.append(row)
    
    # DataFrameの作成
    df = pd.DataFrame(daily_data)
    
    # 日付でソート
    if not df.empty:
        df = df.sort_values("日付", ascending=False)
    
    return df

# 環境別データのテーブルを作成
def create_env_table(data, settings):
    if not data.get("by_env"):
        return pd.DataFrame()
    
    # データの準備
    env_data = []
    for env, dates in data["by_env"].items():
        for date, values in dates.items():
            row = {"環境名": env, "日付": date}
            row.update(values)
            env_data.append(row)
    
    # DataFrameの作成
    df = pd.DataFrame(env_data)
    
    # 日付でソート
    if not df.empty:
        df = df.sort_values(["環境名", "日付"], ascending=[True, False])
    
    return df

# 担当者別データのテーブルを作成
def create_person_table(data, settings):
    if not data.get("by_name"):
        return pd.DataFrame()
    
    # データの準備
    person_data = []
    for date, names in data["by_name"].items():
        for name, count in names.items():
            person_data.append({
                "日付": date,
                "担当者": name,
                "消化数": count
            })
    
    # DataFrameの作成
    df = pd.DataFrame(person_data)
    
    # 日付でソート
    if not df.empty:
        df = df.sort_values(["日付", "担当者"], ascending=[False, True])
    
    return df

def make_progress_svg(data, settings, width=120, height=16):
    results = data.get("total", {})
    incompleted = data.get("stats", {}).get("incompleted", 0)
    bar_data = []
    total = sum(results.get(r, 0) for r in settings["test_status"]["results"]) + incompleted
    if total == 0:
        return ""
    for result in settings["test_status"]["results"]:
        value = results.get(result, 0)
        if value > 0:
            bar_data.append((settings["app"]["bar"]["colors"].get(result, "#ccc"), value))
    if incompleted > 0:
        bar_data.append((settings["app"]["bar"]["colors"].get(settings["test_status"]["labels"]["not_run"], "#ccc"), incompleted))
    # SVG生成
    svg = f'<svg width="{width}" height="{height}">' \
        + ''.join([
            f'<rect x="{sum(width * v / total for _, v in bar_data[:i])}" y="0" width="{width * value / total}" height="{height}" fill="{color}" />'
            for i, (color, value) in enumerate(bar_data)
        ]) + '</svg>'
    return svg

# エラー情報のテーブルを作成
def create_error_table(project_data):
    error_data = []
    for data in project_data.get("aggregate_data", []):
        if "error" in data:
            error_data.append({
                "ファイル名": data.get("file", ""),
                "種別": "エラー",
                "エラー種別": data["error"].get("type", "不明"),
                "メッセージ": data["error"].get("message", "")
            })
        elif "warning" in data:
            error_data.append({
                "ファイル名": data.get("file", ""),
                "種別": "ワーニング",
                "エラー種別": data["warning"].get("type", "不明"),
                "メッセージ": data["warning"].get("message", "")
            })
    
    return pd.DataFrame(error_data)

# デフォルトプロジェクトの設定を読み込む
def load_default_project():
    try:
        with open("default_project_config.json", "r", encoding="utf-8") as f:
            config = json.load(f)
            return config.get("default_project")
    except FileNotFoundError:
        return None

# デフォルトプロジェクトを設定
def save_default_project(project_name):
    config = {"default_project": project_name}
    with open("default_project_config.json", "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

# プロジェクトファイル作成用のポップアップ
def create_project_popup():
    # セッション状態の初期化
    if 'file_count' not in st.session_state:
        st.session_state.file_count = 1
    if 'current_tab' not in st.session_state:
        st.session_state.current_tab = 0
    if 'tabs' not in st.session_state:
        st.session_state.tabs = [0]  # タブのインデックスリスト
    if 'show_project_popup' not in st.session_state:
        st.session_state.show_project_popup = True
    
    with st.form("project_form"):
        st.subheader("プロジェクト設定")
        
        # プロジェクト名称
        project_name = st.text_input("プロジェクト名称 (*)", key="project_name")
        
        # ファイル情報
        st.subheader("取得元ファイルパス/URL")
        
        # タブの表示
        if st.session_state.tabs:
            tab_labels = [f"ファイル {i+1}" for i in range(len(st.session_state.tabs))]
            tabs = st.tabs(tab_labels)
            
            # 現在のタブを追跡
            current_tab_id = tabs[st.session_state.current_tab].id
            for i, tab in enumerate(tabs):
                if tab.id == current_tab_id:
                    st.session_state.current_tab = i
                    break
            
            file_info = []
            for i, (tab, tab_id) in enumerate(zip(tabs, st.session_state.tabs)):
                with tab:                    
                    identifier = st.text_input("名称", key=f"identifier_{tab_id}")
                    file_type = st.selectbox(
                        "タイプ *", 
                        ["local", "sharepoint"], 
                        key=f"type_{tab_id}"
                    )
                    path = st.text_input("パスまたはURL *", key=f"path_{tab_id}")
                    
                    # 必須項目が入力されている場合のみファイル情報を追加
                    if file_type and path.strip():
                        file_info.append({
                            "type": file_type,
                            "identifier": identifier.strip(),
                            "path": path.strip()
                        })

                    # 削除ボタン（タブが1つの場合は非活性）
                    disabled = len(st.session_state.tabs) <= 1
                    if st.form_submit_button(f"🗑 ファイル{i+1} を削除", disabled=disabled):
                        del st.session_state.tabs[i]
                        # 削除後のタブ数に応じて現在のタブインデックスを調整
                        st.session_state.current_tab = min(i, len(st.session_state.tabs) - 1)
                        st.rerun()

        # ファイル追加ボタン
        if st.form_submit_button("＋ 追加", help="ファイルを追加"):
            new_tab = max(st.session_state.tabs) + 1 if st.session_state.tabs else 0
            st.session_state.tabs.append(new_tab)
            st.session_state.current_tab = len(st.session_state.tabs) - 1
            st.rerun()

        # 区切り線
        st.markdown("---")

        # ボタンを横に並べる
        col1, col2, col3 = st.columns([1, 2, 7])
        with col1:
            submitted = st.form_submit_button("保存", type="primary")
        with col2:
            if st.form_submit_button("キャンセル"):
                st.session_state.show_project_popup = False
                st.rerun()
        
        if submitted:
            if not project_name:
                st.error("プロジェクト名称を入力してください")
                return None
                
            if not file_info:
                st.error("少なくとも1つのファイル情報を入力してください")
                return None
                
            # プロジェクトデータの作成
            project_data = {
                "project": {
                    "project_name": project_name,
                    "files": file_info
                }
            }
            
            # projectsディレクトリの作成
            projects_dir = Path("projects")
            projects_dir.mkdir(exist_ok=True)
            
            # ファイル名のサニタイズ
            safe_project_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', project_name)
            json_path = projects_dir / f"{safe_project_name}.json"
            
            # JSONファイルの保存
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(project_data, f, ensure_ascii=False, indent=2)
                
            st.success(f"プロジェクトファイルを保存しました。\n{json_path.name}")
            # セッション状態をクリア
            st.session_state.file_count = 1
            st.session_state.current_tab = 0
            st.session_state.tabs = [0]
            st.session_state.show_project_popup = False
            return str(json_path)
            
        return None

# メインアプリケーション
def main():
    st.set_page_config(
        page_title="TestTraQ",
        page_icon="📊",
        layout="wide"
    )
    
    # 設定の読み込み
    settings = load_settings()
    
    # サイドバー
    st.sidebar.title("TestTraQ")
    reload_clicked = st.sidebar.button("🔄 集計データ再読み込み")
    create_project_clicked = st.sidebar.button("📝 新規プロジェクト作成")
    
    if create_project_clicked:
        st.session_state.show_project_popup = True
        
    if st.session_state.get("show_project_popup", False):
        project_path = create_project_popup()
        if project_path:
            st.session_state.show_project_popup = False
            st.rerun()
        return  # プロジェクト作成中は以降の処理をスキップ

    # プロジェクトファイルの選択
    project_files = list(Path("projects").glob("*.json"))
    if not project_files:
        st.error("プロジェクトファイルが見つかりません。")
        return

    # URLクエリパラメータから前回選択したプロジェクトを取得
    params = st.query_params
    last_project = params.get("project")
    
    # プロジェクトファイルのオプションを準備
    project_options = {p.stem: p for p in project_files}
    
    # デフォルトプロジェクトの読み込み
    default_project = load_default_project()
    
    # プロジェクト名のリストを作成（デフォルトプロジェクトには☆を付ける）
    project_names = []
    project_name_to_stem = {}  # 表示名からstemへのマッピング
    for stem in project_options.keys():
        display_name = f"⭐ {stem}" if stem == default_project else stem
        project_names.append(display_name)
        project_name_to_stem[display_name] = stem

    # 前回選択したプロジェクトまたはデフォルトプロジェクトが存在する場合は、それをデフォルト値として設定
    default_index = 0
    if last_project in project_options:
        display_name = f"⭐ {last_project}" if last_project == default_project else last_project
        default_index = project_names.index(display_name)
    elif default_project in project_options:
        default_index = project_names.index(f"⭐ {default_project}")

    selected_display_name = st.sidebar.selectbox(
        "プロジェクトを選択",
        options=project_names,
        index=default_index
    )
    
    # 表示名からプロジェクト名（stem）を取得
    selected_project_name = project_name_to_stem[selected_display_name]
    
    # 選択されたプロジェクト名をURLクエリパラメータに保存
    st.query_params["project"] = selected_project_name
    
    # 選択されたプロジェクトのPathオブジェクトを取得
    selected_project = project_options[selected_project_name]

    # 再集計状態管理
    if 'reload_state' not in st.session_state:
        st.session_state['reload_state'] = 'idle'

    if reload_clicked:
        if selected_project:
            project_path = str(selected_project)
            python_exe = sys.executable
            flag_path = f"{project_path}.reloading"
            with open(flag_path, "w") as f:
                f.write("reloading")
            cmd = [python_exe, "StartProcess.py", project_path, "--project", project_path, "--on_reload"]
            subprocess.Popen(cmd)
            st.session_state['reload_state'] = 'waiting'
            st.rerun()
        else:
            st.warning("プロジェクトファイルを選択してください。")

    if st.session_state.get('reload_state') == 'waiting':
        flag_path = f"{str(selected_project)}.reloading"
        if not os.path.exists(flag_path):
            st.session_state['reload_state'] = 'idle'
            st.success("再集計が完了しました。")
            st.rerun()
        else:
            st.info("再集計中です。しばらくお待ちください。")
            time.sleep(2)
            st.rerun()
    
    # プロジェクトデータの読み込み
    project_data = load_project_data(selected_project)
    if not project_data:
        return
    
    # プロジェクト名とお気に入りボタンの表示
    col1, col2 = st.columns([10, 1])
    with col1:
        st.title(project_data["project"]["project_name"])
    with col2:
        is_default = selected_project_name == default_project
        star_icon = "⭐" if is_default else "☆"
        if st.button(star_icon, key="favorite_button"):
            if is_default:
                # デフォルト設定を解除
                save_default_project(None)
            else:
                # デフォルトプロジェクトとして設定
                save_default_project(selected_project_name)
            st.rerun()

    st.header("集計結果")
    
    # タブの作成
    tab1, tab2, tab3 = st.tabs(["全体集計", "ファイル別集計", "エラー情報"])
    
    with tab1:
        # 全体集計タブ
        if "aggregate_data" in project_data:
            # エラーとワーニングのあるデータを除外
            filtered_data = [d for d in project_data["aggregate_data"] 
                           if "error" not in d and "warning" not in d]
            
            if filtered_data:
                # 総合集計の表示
                total_stats = {
                    "all": sum(d["stats"]["all"] for d in filtered_data),
                    "excluded": sum(d["stats"]["excluded"] for d in filtered_data),
                    "available": sum(d["stats"]["available"] for d in filtered_data),
                    "executed": sum(d["stats"]["executed"] for d in filtered_data),
                    "completed": sum(d["stats"]["completed"] for d in filtered_data),
                    "incompleted": sum(d["stats"]["incompleted"] for d in filtered_data)
                }
                
                # 集計情報の表示
                col1, col2, col3 = st.columns(3)
                with col1:
                    main_text = f"{total_stats['available']}"
                    sub_text = f"(Total: {total_stats['all']} / 対象外: {total_stats['excluded']})"
                    st.markdown(
                        f"<div>項目数</div>"
                        f"<div style='font-size:2.5em; line-height:1.1;'>{main_text} "
                        f"<span style='font-size:0.5em; color:#aaa;'>{sub_text}</span></div>",
                        unsafe_allow_html=True
                    )
                with col2:
                    main_text = f"{total_stats['completed']}/{total_stats['available']}"
                    sub_text = f"({total_stats['completed']/total_stats['available']*100:.1f}%)"
                    st.markdown(
                        f"<div>完了率</div>"
                        f"<div style='font-size:2.5em; line-height:1.1;'>{main_text} "
                        f"<span style='font-size:0.5em; color:#aaa;'>{sub_text}</span></div>",
                        unsafe_allow_html=True
                    )
                with col3:
                    main_text = f"{total_stats['executed']}/{total_stats['available']}"
                    sub_text = f"({total_stats['executed']/total_stats['available']*100:.1f}%)"
                    st.markdown(
                        f"<div>消化率</div>"
                        f"<div style='font-size:2.5em; line-height:1.1;'>{main_text} "
                        f"<span style='font-size:0.5em; color:#aaa;'>{sub_text}</span></div>",
                        unsafe_allow_html=True
                    )

                # 全体集計グラフの表示
                total_results = {}
                for data in filtered_data:
                    for key, value in data["total"].items():
                        total_results[key] = total_results.get(key, 0) + value
                
                st.plotly_chart(create_progress_chart(
                    {"total": total_results,
                     "stats": total_stats},
                    settings
                ), use_container_width=True, key="summary_progress_chart")

                # 区切り線
                st.markdown("---")

                # ファイル一覧の表示
                file_data = []
                for data in project_data["aggregate_data"]:
                    if "error" in data:
                        status = "⚠️エラー"
                        status_color = "red"
                        file_data.append({
                            "ファイル名": data.get("file", ""),
                            "項目数": "-",
                            "進捗": "-",
                            "消化率": "-",
                            "完了率": "-",
                            "状態": status,
                            "更新日": data.get("last_updated", "")
                        })
                    elif "warning" in data:
                        status = "ワーニング"
                        status_color = "orange"
                        file_data.append({
                            "ファイル名": data.get("file", ""),
                            "項目数": "-",
                            "進捗": make_progress_svg(data, settings),
                            "消化率": "-",
                            "完了率": "-",
                            "状態": status,
                            "更新日": data.get("last_updated", "")
                        })
                    elif "stats" in data:
                        status = ""
                        status_color = "green"
                        available = data["stats"].get("available", 0)
                        executed = data["stats"].get("executed", 0)
                        completed = data["stats"].get("completed", 0)
                        file_data.append({
                            "ファイル名": data.get("file", ""),
                            "項目数": available,
                            "進捗": make_progress_svg(data, settings),
                            "消化率": f"{executed}/{available} ({(executed/available*100):.1f}%)" if available else "-",
                            "完了率": f"{completed}/{available} ({(completed/available*100):.1f}%)" if available else "-",
                            "状態": status,
                            "更新日": data.get("last_updated", "")
                        })
                    else:
                        file_data.append({
                            "ファイル名": data.get("file", ""),
                            "項目数": "-",
                            "進捗": "-",
                            "消化率": "-",
                            "完了率": "-",
                            "状態": "不明",
                            "更新日": data.get("last_updated", "")
                        })
                
                df = pd.DataFrame(file_data)
                for col in ["項目数", "消化率", "完了率"]:
                    if col in df.columns:
                        df[col] = df[col].astype(str)
                # HTMLテーブルでSVGを表示
                def df_to_html(df):
                    html = '<table style="width:100%; border-collapse:collapse;">'
                    html += '<tr>' + ''.join(f'<th style="padding:2px 4px;">{c}</th>' for c in df.columns) + '</tr>'
                    for _, row in df.iterrows():
                        html += '<tr>' + ''.join(
                            f'<td style="padding:2px 4px; vertical-align:middle;">{row[c] if c != "進捗" else row[c]}</td>'
                            for c in df.columns
                        ) + '</tr>'
                    html += '</table>'
                    return html
                st.markdown(df_to_html(df), unsafe_allow_html=True)
    
    with tab2:
        # ファイル別タブ
        if "aggregate_data" in project_data:
            # ファイル選択
            file_options = [d["selector_label"] for d in project_data["aggregate_data"]]
            selected_file = st.selectbox("ファイルを選択", options=file_options)
            
            # 選択されたファイルのデータを取得
            matching_data = [d for d in project_data["aggregate_data"] if d["selector_label"] == selected_file]
            if matching_data:
                file_data = matching_data[0]
                if "error" not in file_data:
                    # 集計情報の表示
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        main_text = f"{file_data['stats']['available']}"
                        sub_text = f"(Total: {file_data['stats']['all']} / 対象外: {file_data['stats']['excluded']})"
                        st.markdown(
                            f"<div>項目数</div>"
                            f"<div style='font-size:2.5em; line-height:1.1;'>{main_text} "
                            f"<span style='font-size:0.5em; color:#aaa;'>{sub_text}</span></div>",
                            unsafe_allow_html=True
                        )
                    with col2:
                        main_text = f"{file_data['stats']['completed']}/{file_data['stats']['available']}"
                        sub_text = f"({file_data['stats']['completed']/file_data['stats']['available']*100:.1f}%)"
                        st.markdown(
                            f"<div>完了率</div>"
                            f"<div style='font-size:2.5em; line-height:1.1;'>{main_text} "
                            f"<span style='font-size:0.5em; color:#aaa;'>{sub_text}</span></div>",
                            unsafe_allow_html=True
                        )
                    with col3:
                        main_text = f"{file_data['stats']['executed']}/{file_data['stats']['available']}"
                        sub_text = f"({file_data['stats']['executed']/file_data['stats']['available']*100:.1f}%)"
                        st.markdown(
                            f"<div>消化率</div>"
                            f"<div style='font-size:2.5em; line-height:1.1;'>{main_text} "
                            f"<span style='font-size:0.5em; color:#aaa;'>{sub_text}</span></div>",
                            unsafe_allow_html=True
                        )

                    # 進捗状況の表示
                    st.plotly_chart(create_progress_chart(file_data, settings), use_container_width=True, key=f"file_progress_chart_{selected_file}")
                    # 区切り線
                    st.markdown("---")

                    # サブタブの作成
                    subtab1, subtab2, subtab3 = st.tabs(["日付別", "環境別", "担当者別"])
                    with subtab1:
                        st.dataframe(create_daily_table(file_data, settings), hide_index=True, use_container_width=True)
                    with subtab2:
                        st.dataframe(create_env_table(file_data, settings), hide_index=True, use_container_width=True)
                    with subtab3:
                        st.dataframe(create_person_table(file_data, settings), hide_index=True, use_container_width=True)
                else:
                    st.error(f"エラー: {file_data['error']['message']}")
            else:
                st.error(f"選択されたファイル '{selected_file}' のデータが見つかりません。")

    with tab3:
        # エラー情報タブ
        error_df = create_error_table(project_data)
        if not error_df.empty:
            st.dataframe(error_df, hide_index=True, use_container_width=True)
        else:
            st.info("エラー・ワーニングはありません。")

if __name__ == "__main__":
    main() 