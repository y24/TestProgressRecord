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
    
    # 前回選択したプロジェクトが存在する場合は、それをデフォルト値として設定
    default_index = 0
    if last_project in project_options:
        default_index = list(project_options.keys()).index(last_project)

    selected_project_name = st.sidebar.selectbox(
        "プロジェクトを選択",
        options=list(project_options.keys()),
        index=default_index
    )
    
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
    
    # プロジェクト名の表示
    st.title(project_data["project"]["project_name"])
    
    # タブの作成
    tab1, tab2 = st.tabs(["全体集計", "ファイル別"])
    
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
                
                # 進捗状況の表示
                total_results = {}
                for data in filtered_data:
                    for key, value in data["total"].items():
                        total_results[key] = total_results.get(key, 0) + value
                
                st.plotly_chart(create_progress_chart(
                    {"total": total_results,
                     "stats": total_stats},
                    settings
                ), use_container_width=True, key="summary_progress_chart")
                
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
                
                # ファイル一覧の表示
                file_data = []
                for data in project_data["aggregate_data"]:
                    if "error" in data:
                        status = "エラー"
                        status_color = "red"
                        file_data.append({
                            "ファイル名": data.get("file", ""),
                            "項目数": "-",
                            "進捗": "⚠️",
                            "消化率": "-",
                            "完了率": "-",
                            "状態": status,
                            "更新日": data.get("last_updated", "")
                        })
                    elif "warning" in data:
                        status = "警告"
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
                        status = "正常"
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
                    # 進捗状況の表示
                    st.plotly_chart(create_progress_chart(file_data, settings), use_container_width=True, key=f"file_progress_chart_{selected_file}")
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

if __name__ == "__main__":
    main() 