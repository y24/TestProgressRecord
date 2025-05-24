import streamlit as st
import json
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime

# JSONファイル読み込み
with open("projects/サンプルプロジェクト01.json", encoding="utf-8") as f:
    data = json.load(f)

# データ取得
all_data = data["all_data"]
daily = all_data["daily"]
stats = all_data["stats"]

# 日付ごとのデータ
dates = sorted(daily.keys())
# 総テスト件数
total_tests = stats["all"]

# 計画消化数の累積
cumulative_plan = []
plan_sum = 0
for d in dates:
    plan_sum += daily[d].get("計画数", 0)
    cumulative_plan.append(plan_sum)

df = pd.DataFrame([
    {
        "date": d,
        "未実施テスト項目数": total_tests - sum([daily[dt].get("消化数", 0) for dt in dates if dt <= d]),
        "消化数": daily[d].get("消化数", 0),
        "Fail": daily[d].get("Fail", 0),
        "計画累計消化数": cumulative_plan[i],
        "計画未実施数": total_tests - cumulative_plan[i]
    }
    for i, d in enumerate(dates)
])
df["累積Fail数"] = df["Fail"].cumsum()

fig = go.Figure()

# 未実施テスト項目数
fig.add_trace(go.Scatter(
    x=df["date"], y=df["未実施テスト項目数"],
    mode="lines",
    name="未実施テスト項目数",
    line=dict(width=3, color="#636EFA"),
    fill='tozeroy',
    fillcolor="rgba(99,110,250,0.08)"
))

# 計画線
fig.add_trace(go.Scatter(
    x=df["date"], y=df["計画未実施数"],
    mode="lines",
    name="計画線（未実施予定）",
    line=dict(width=2, color="#00CC96", dash="dash")
))

# 累積Fail数
fig.add_trace(go.Scatter(
    x=df["date"], y=df["累積Fail数"],
    mode="lines",
    name="累積バグ検出数（Fail）",
    line=dict(width=3, color="#EF553B", dash="dot")
))

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