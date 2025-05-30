import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import json, os, time, re, subprocess, sys
from datetime import datetime
import numpy as np
from pathlib import Path
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import pyperclip
import base64

from libs import AppConfig, Labels, DataConversion
from libs.webui_chart_manager import ChartManager

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
    return ChartManager.create_progress_chart(data, settings)

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
    return ChartManager.make_progress_svg(data, settings, width, height)

# エラー情報のテーブルを作成
def create_error_table(project_data):
    error_data = []
    for data in project_data.get("gathered_data", []):
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
    settings = AppConfig.load_settings()
    return settings.get("webui", {}).get("default_project")

# デフォルトプロジェクトを設定
def save_default_project(project_name):
    settings = AppConfig.load_settings()
    if "webui" not in settings:
        settings["webui"] = {}
    settings["webui"]["default_project"] = project_name
    AppConfig.save_settings(settings)

# 表示設定を読み込む
def load_display_settings():
    settings = AppConfig.load_settings()
    return settings.get("webui", {}).get("display_settings", {
        "axis_type": "時間軸で表示",  # デフォルト値
        "show_plan_line": True  # デフォルト値
    })

# 表示設定を保存
def save_display_settings(display_settings):
    settings = AppConfig.load_settings()
    if "webui" not in settings:
        settings["webui"] = {}
    settings["webui"]["display_settings"] = display_settings
    AppConfig.save_settings(settings)

# PB図を作成
def create_pb_chart(project_data, settings, axis_type="時間軸で表示", show_plan_line=True):
    return ChartManager.create_pb_chart(project_data, settings, axis_type, show_plan_line)

# メインアプリケーション
def main():
    st.set_page_config(
        page_title="TestTraQ",
        page_icon="📊",
        layout="wide"
    )
    
    # セッション状態の初期化
    if 'show_data' not in st.session_state:
        st.session_state.show_data = False
    if 'reload_state' not in st.session_state:
        st.session_state['reload_state'] = 'idle'
    if 'previous_project' not in st.session_state:
        st.session_state.previous_project = None

    # 設定の読み込み
    settings = AppConfig.load_settings()
    
    # サイドバー
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

    # st.sidebar.markdown("### プロジェクト")

    selected_display_name = st.sidebar.selectbox(
        "プロジェクトを選択",
        options=project_names,
        index=default_index
    )
    
    # 表示名からプロジェクト名（stem）を取得
    selected_project_name = project_name_to_stem[selected_display_name]
    
    # プロジェクトが変更された場合、URLパラメータを更新して画面を再読み込み
    if st.session_state.previous_project != selected_project_name:
        st.session_state.previous_project = selected_project_name
        st.query_params["project"] = selected_project_name
        st.rerun()
    
    # 選択されたプロジェクトのPathオブジェクトを取得
    selected_project = project_options[selected_project_name]

    # プロジェクトデータの読み込み
    project_data = load_project_data(selected_project)
    if not project_data:
        return

    # サイドバーにtsvデータボタンを追加
    if st.sidebar.button("📋 tsvデータを表示"):
        st.session_state.show_data = not st.session_state.show_data

    # tsvデータ表示モードの場合
    if st.session_state.show_data:
        # 集計データを2次元配列に変換
        array_data = DataConversion.convert_to_2d_array(project_data["gathered_data"], settings)
        # TSV形式に変換
        tsv_data = "\n".join(["\t".join(map(str, row)) for row in array_data])
        
        st.markdown("### tsvデータ")
        st.text("以下のデータをコピーしてください：")
        st.code(tsv_data, height=600)
        if st.button("戻る"):
            st.session_state.show_data = False
            st.rerun()
        return  # メイン画面の表示をスキップ

    st.sidebar.markdown("---")

    # グラフ表示設定
    st.sidebar.markdown("### グラフ表示設定")
    # 表示設定の読み込み
    display_settings = load_display_settings()
    
    # 計画線表示の設定
    show_plan_line = st.sidebar.toggle(
        "計画線を表示",
        value=display_settings.get("show_plan_line", True),
        help="PB図の計画線の表示/非表示を切り替えます"
    )

    # 進捗グラフの表示設定
    axis_type = st.sidebar.radio(
        "横軸の表示方法",
        ["時間軸", "等間隔"],
        index=0 if display_settings["axis_type"] == "時間軸で表示" else 1,
        captions=["実際の日付間隔で表示する", "データのない日付は詰めて表示する"]
    )

    # 設定が変更された場合は保存
    if (axis_type != display_settings["axis_type"] or
        show_plan_line != display_settings.get("show_plan_line", True)):
        save_display_settings({
            "axis_type": axis_type,
            "show_plan_line": show_plan_line
        })
        st.rerun()  # 設定を反映するために再読み込み

    # 再集計状態管理
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
    
    # プロジェクト名とお気に入りボタンの表示
    col1, col2, col3 = st.columns([14, 1, 1])
    with col1:
        # プロジェクト名
        st.header(project_data["project"]["project_name"])
    with col2:
        is_default = selected_project_name == default_project
        star_icon = "⭐" if is_default else "☆"
        if st.button(star_icon, key="favorite_button", help="デフォルトプロジェクトに設定"):
            if is_default:
                # デフォルト設定を解除
                save_default_project(None)
            else:
                # デフォルトプロジェクトとして設定
                save_default_project(selected_project_name)
            st.rerun()
    with col3:
        if st.button("🔄", key="reload_button", help="再集計を実行"):
            if selected_project:
                project_path = str(selected_project)
                python_exe = sys.executable
                flag_path = f"{project_path}.reloading"
                with open(flag_path, "w") as f:
                    f.write("reloading")
                cmd = [python_exe, "StartProcess.py", project_path, "--project", project_path, "--on_reload", "--webui"]
                subprocess.Popen(cmd)
                st.session_state['reload_state'] = 'waiting'
                st.rerun()

    # 最終更新日時の表示
    if "last_loaded" in project_data["project"]:
        last_loaded = datetime.fromisoformat(project_data["project"]["last_loaded"])
        st.caption(f"最終更新: {last_loaded.strftime('%Y/%m/%d %H:%M')}")

    # エラー情報の確認
    error_df = create_error_table(project_data)
    has_errors = not error_df.empty

    # タブの作成（エラーがある場合のみエラー情報タブを表示）
    tabs = ["全体集計", "ファイル別集計"]
    if has_errors:
        tabs.append("⚠️エラー情報")
    
    all_tabs = st.tabs(tabs)
    tab1, tab2 = all_tabs[0], all_tabs[1]
    tab3 = all_tabs[2] if has_errors else None
    
    with tab1:
        # 全体集計タブ
        if "gathered_data" in project_data:

            # PB図の表示
            pb_fig = create_pb_chart(project_data, settings, axis_type, display_settings.get("show_plan_line", True))
            if pb_fig:
                st.plotly_chart(pb_fig, use_container_width=True, key=f"pb_chart_{selected_project_name}", config={"displayModeBar": False, "scrollZoom": False})

            # エラーとワーニングのあるデータを除外
            filtered_data = [d for d in project_data["gathered_data"] 
                           if "error" not in d and "warning" not in d]
            
            if filtered_data:
                # 総合集計の表示
                total_stats = {
                    "all": sum(d["stats"]["all"] for d in filtered_data),
                    "excluded": sum(d["stats"]["excluded"] for d in filtered_data),
                    "available": sum(d["stats"]["available"] for d in filtered_data),
                    "executed": sum(d["stats"]["executed"] for d in filtered_data),
                    "completed": sum(d["stats"]["completed"] for d in filtered_data),
                    "incompleted": sum(d["stats"]["incompleted"] for d in filtered_data),
                    "planned": sum(d["stats"]["planned"] for d in filtered_data)
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
                ), use_container_width=True, key="summary_progress_chart", config={"displayModeBar": False, "scrollZoom": False})

                # 区切り線
                st.markdown("---")

                # ファイル一覧の表示
                file_data = []
                for data in project_data["gathered_data"]:
                    if "error" in data:
                        status = "❌エラー"
                        status_color = "red"
                        file_data.append({
                            "ファイル名": data.get("file", ""),
                            "DL": "📥",
                            "項目数": "-",
                            "計画数": "-",
                            "進捗": "-",
                            "消化率": "-",
                            "完了率": "-",
                            "状態": status,
                            "更新日時": data.get("last_updated", ""),
                            "filepath": data.get("filepath", "")
                        })
                    elif "warning" in data:
                        status = "⚠️警告"
                        status_color = "orange"
                        # 警告時の統計情報の取得
                        available = data.get("stats", {}).get("available", 0)
                        planned = data.get("stats", {}).get("planned", 0)
                        executed = data.get("stats", {}).get("executed", 0)
                        completed = data.get("stats", {}).get("completed", 0)
                        file_data.append({
                            "ファイル名": data.get("file", ""),
                            "DL": "📥",
                            "項目数": available,
                            "計画数": planned,
                            "進捗": make_progress_svg(data, settings),
                            "消化率": Labels.make_count_and_rate_text(executed, available),
                            "完了率": Labels.make_count_and_rate_text(completed, available),
                            "状態": status,
                            "更新日時": data.get("last_updated", ""),
                            "filepath": data.get("filepath", "")
                        })
                    elif "stats" in data:
                        status = "✅正常"
                        status_color = "green"
                        available = data["stats"].get("available", 0)
                        planned = data["stats"].get("planned", 0)
                        executed = data["stats"].get("executed", 0)
                        completed = data["stats"].get("completed", 0)
                        file_data.append({
                            "ファイル名": data.get("file", ""),
                            "DL": "📥",
                            "項目数": available,
                            "計画数": planned,
                            "進捗": make_progress_svg(data, settings),
                            "消化率": Labels.make_count_and_rate_text(executed, available),
                            "完了率": Labels.make_count_and_rate_text(completed, available),
                            "状態": status,
                            "更新日時": data.get("last_updated", ""),
                            "filepath": data.get("filepath", "")
                        })
                    else:
                        file_data.append({
                            "ファイル名": data.get("file", ""),
                            "DL": "-",
                            "項目数": "-",
                            "計画数": "-",
                            "進捗": "-",
                            "消化率": "-",
                            "完了率": "-",
                            "状態": "不明",
                            "更新日時": data.get("last_updated", ""),
                            "filepath": data.get("filepath", "")
                        })
                
                df = pd.DataFrame(file_data)
                for col in ["項目数", "計画数", "消化率", "完了率"]:
                    if col in df.columns:
                        df[col] = df[col].astype(str)

                # HTMLテーブルでSVGを表示し、ダウンロードボタンを追加
                def df_to_html(df):
                    html = '<table style="width:100%; border-collapse:collapse;">'
                    html += '<tr>' + ''.join(f'<th style="padding:2px 4px;">{c}</th>' for c in df.columns if c != "filepath") + '</tr>'
                    for _, row in df.iterrows():
                        html += '<tr>'
                        for c in df.columns:
                            if c == "filepath":
                                continue
                            elif c == "DL" and row["filepath"]:
                                # ダウンロードボタンを作成
                                file_path = row["filepath"]
                                if os.path.exists(file_path):
                                    try:
                                        # サブプロセスでファイルエンコードを実行
                                        python_exe = sys.executable
                                        encoder_script = os.path.join(os.path.dirname(__file__), "file_encoder.py")
                                        result = subprocess.run(
                                            [python_exe, encoder_script, file_path],
                                            capture_output=True,
                                            text=True,
                                            timeout=30  # タイムアウトを30秒に設定
                                        )
                                        
                                        if result.returncode == 0:
                                            b64 = result.stdout.strip()
                                            download_link = f'<span style="font-size:0.8em;"><a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}" style="text-decoration:none;">■</a></span>'
                                            html += f'<td style="padding:2px 4px; vertical-align:middle;">{download_link}</td>'
                                        else:
                                            st.error(f"ファイルのエンコードに失敗しました: {result.stderr}")
                                            html += '<td style="padding:2px 4px; vertical-align:middle;">-</td>'
                                    except subprocess.TimeoutExpired:
                                        st.warning(f"ファイルのエンコードがタイムアウトしました: {os.path.basename(file_path)}")
                                        html += '<td style="padding:2px 4px; vertical-align:middle;">-</td>'
                                    except Exception as e:
                                        st.error(f"エラーが発生しました: {str(e)}")
                                        html += '<td style="padding:2px 4px; vertical-align:middle;">-</td>'
                                else:
                                    html += '<td style="padding:2px 4px; vertical-align:middle;">-</td>'
                            else:
                                html += f'<td style="padding:2px 4px; vertical-align:middle;">{row[c] if c != "進捗" else row[c]}</td>'
                        html += '</tr>'
                    html += '</table>'
                    return html
                st.markdown(df_to_html(df), unsafe_allow_html=True)
    
    with tab2:
        # ファイル別タブ
        if "gathered_data" in project_data:
            # ファイル選択
            file_options = [d["selector_label"] for d in project_data["gathered_data"]]
            selected_file = st.selectbox("ファイルを選択", options=file_options)
            
            # 選択されたファイルのデータを取得
            matching_data = [d for d in project_data["gathered_data"] if d["selector_label"] == selected_file]
            if matching_data:
                file_data = matching_data[0]
                if "error" not in file_data:
                    # PB図の表示
                    pb_fig = create_pb_chart({"gathered_data": [file_data]}, settings, axis_type, display_settings.get("show_plan_line", True))
                    if pb_fig:
                        st.plotly_chart(pb_fig, use_container_width=True, key=f"file_pb_chart_{selected_file}", config={"displayModeBar": False, "scrollZoom": False})

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
                    st.plotly_chart(create_progress_chart(file_data, settings), use_container_width=True, key=f"file_progress_chart_{selected_file}", config={"displayModeBar": False, "scrollZoom": False})
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

    if has_errors:
        with tab3:
            # エラー情報タブ
            st.dataframe(error_df, hide_index=True, use_container_width=True)

if __name__ == "__main__":
    main() 

# launch command
# streamlit run WebUI.py
