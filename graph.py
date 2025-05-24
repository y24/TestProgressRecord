import streamlit as st
import json
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime

# JSONファイル読み込み
with open("projects/サンプルプロジェクト01.json", encoding="utf-8") as f:
    data = json.load(f)

# 1つ目のgathered_dataを利用
aggregate = data["gathered_data"][0]
daily = aggregate["daily"]

# 日付でソート
dates = sorted(daily.keys())
# 総テスト件数
total_tests = aggregate["stats"]["all"]

# データフレーム生成
df = pd.DataFrame([
    {
        "date": d,
        "未実施テスト項目数": total_tests - sum([daily[dt]["消化数"] for dt in dates if dt <= d]),
        "消化数": daily[d]["消化数"],
        "Fail": daily[d]["Fail"]
    }
    for d in dates
])
df["累積Fail数"] = df["Fail"].cumsum()

# Plotlyグラフ
fig = go.Figure()

# 未実施テスト項目数
fig.add_trace(go.Scatter(
    x=df["date"],
    y=df["未実施テスト項目数"],
    mode="lines+markers",
    name="未実施テスト項目数",
    line=dict(width=3, color="#636EFA"),
    marker=dict(size=8),
    fill='tozeroy',
    fillcolor="rgba(99,110,250,0.08)"
))

# 累積Fail数
fig.add_trace(go.Scatter(
    x=df["date"],
    y=df["累積Fail数"],
    mode="lines+markers",
    name="累積バグ検出数（Fail）",
    line=dict(width=3, color="#EF553B", dash="dot"),
    marker=dict(size=8)
))

# モダンなデザイン調整
fig.update_layout(
    title="テスト進捗と不具合検出状況（PB図）",
    xaxis_title="日付",
    yaxis_title="件数",
    xaxis=dict(
        tickmode='array',
        tickvals=df["date"],
        ticktext=[datetime.strptime(d, "%Y-%m-%d").strftime("%m/%d") for d in df["date"]],
        showgrid=False
    ),
    yaxis=dict(
        showgrid=True,
        gridcolor="rgba(200,200,200,0.2)"
    ),
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    ),
    plot_bgcolor="#FFF",
    font=dict(
        family="sans-serif",
        size=16
    )
)

st.title("テスト進捗と不具合検出状況（PB図）")
st.plotly_chart(fig, use_container_width=True)
