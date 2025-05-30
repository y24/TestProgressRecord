import plotly.graph_objects as go
import pandas as pd
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional

class ChartManager:
    @staticmethod
    def create_progress_chart(data: Dict[str, Any], settings: Dict[str, Any]) -> go.Figure:
        """進捗状況のグラフを作成"""
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
                    "color": settings["webui"]["bar"]["colors"].get(result, "gainsboro")
                })
        # 未着手
        if incompleted > 0:
            bar_data.append({
                "name": settings["test_status"]["labels"]["not_run"],
                "value": incompleted,
                "color": settings["webui"]["bar"]["colors"].get(settings["test_status"]["labels"]["not_run"], "gainsboro")
            })

        # 積み上げバー作成
        fig = go.Figure()
        for bar in bar_data:
            fig.add_trace(go.Bar(
                x=[bar["value"]],
                y=[""],
                orientation='h',
                name=bar["name"],
                marker_color=bar["color"],
                text=[bar["value"]],
                textposition='inside',
                textfont=dict(
                    size=16
                ),
            ))

        fig.update_layout(
            barmode='stack',
            showlegend=False,
            dragmode=False,
            height=100,
            margin=dict(l=0, r=0, t=25, b=0),
            xaxis=dict(
                showticklabels=False,
                showgrid=False,
                zeroline=False,
                ticks=''
            ),
            yaxis=dict(
                showticklabels=False,
                showgrid=False,
                zeroline=False,
                ticks=''
            )
        )
        return fig

    @staticmethod
    def create_pb_chart(project_data: Dict[str, Any], settings: Dict[str, Any], axis_type: str = "時間軸で表示", show_plan_line: bool = True) -> Optional[go.Figure]:
        """PB図を作成"""
        if not project_data.get("gathered_data"):
            return None

        # 日付ごとのデータ
        from libs.DataConversion import aggregate_all_daily, aggregate_all_stats
        daily = aggregate_all_daily(project_data.get("gathered_data"))
        # 総テスト件数
        stats = aggregate_all_stats(project_data.get("gathered_data"))

        if not daily or not stats:
            return None

        # 日付ごとのデータ
        dates = sorted(daily.keys())
        if not dates:
            return None

        # 総テスト件数
        total_tests = stats.get("all", 0)

        # 計画消化数の累積
        cumulative_plan = []
        plan_sum = 0
        for d in dates:
            plan_sum += daily[d].get("計画数", 0)
            cumulative_plan.append(plan_sum)

        # 今日の日付を取得
        today = datetime.now().strftime("%Y-%m-%d")

        df = pd.DataFrame([
            {
                "date": d,
                "未実施テスト項目数": total_tests - sum([daily[dt].get("消化数", 0) for dt in dates if dt <= d]),
                "消化数": daily[d].get("消化数", 0),
                "Fail": daily[d].get("Fail", 0),
                "計画累計消化数": cumulative_plan[i],
                "計画未実施数": total_tests - cumulative_plan[i],
                "計画数": daily[d].get("計画数", 0),
                "完了数": daily[d].get("完了数", 0)
            }
            for i, d in enumerate(dates)
        ])
        df["累積Fail数"] = df["Fail"].cumsum()

        # 今日の日付以降のデータを除外した実績値のデータフレームを作成
        df_actual = df[df["date"] <= today].copy()

        fig = go.Figure()

        # --- 縦棒グラフ ---
        # 計画件数（灰色）
        if show_plan_line:
            fig.add_trace(go.Bar(
                x=df["date"], y=df["計画数"],
                name="計画数(日次)",
                marker_color=settings["webui"]["graph"]["colors"]["plan"],
                opacity=0.65,
                width=86400000 * 0.2 if axis_type == "時間軸で表示" else 0.2,
                yaxis="y"
            ))
        # 完了件数
        fig.add_trace(go.Bar(
            x=df["date"], y=df["消化数"],
            name="消化数(日次)",
            marker_color=settings["webui"]["graph"]["colors"]["daily_executed"],
            opacity=0.85,
            width=86400000 * 0.2 if axis_type == "時間軸で表示" else 0.2,
            yaxis="y"
        ))

        # --- 折れ線グラフ ---
        # 未実施テスト項目数（明日以降は除外）
        fig.add_trace(go.Scatter(
            x=df_actual["date"], y=df_actual["未実施テスト項目数"],
            mode="lines",
            name="残項目数",
            line=dict(width=3, color=settings["webui"]["graph"]["colors"]["untested"]),
            fill='tozeroy',
            fillcolor="rgba(99,110,250,0.08)"
        ))

        # 計画線（全期間表示）
        if show_plan_line:
            fig.add_trace(go.Scatter(
                x=df["date"], y=df["計画未実施数"],
                mode="lines",
                name="計画線",
                line=dict(width=2, color=settings["webui"]["graph"]["colors"]["plan"])
            ))

        date_objs = [datetime.strptime(d, "%Y-%m-%d") for d in df["date"]]
        min_date = date_objs[0]
        max_date = date_objs[-1]

        # 7日おきの日付ラベル
        tickvals = []
        ticktexts = []
        cur = min_date
        while cur <= max_date:
            d_str = cur.strftime("%Y-%m-%d")
            if d_str in df["date"].values:
                tickvals.append(d_str)
                ticktexts.append(cur.strftime("%m/%d"))
            cur += timedelta(days=1)

        # 必ず最終日も追加
        if df["date"].values[-1] not in tickvals:
            tickvals.append(df["date"].values[-1])
            ticktexts.append(date_objs[-1].strftime("%m/%d"))

        # 累積Fail数（明日以降は除外）
        fig.add_trace(go.Scatter(
            x=df_actual["date"], y=df_actual["累積Fail数"],
            mode="lines",
            name="不具合検出数(累積)",
            line=dict(width=3, color=settings["webui"]["graph"]["colors"]["fail"], dash="dot"),
            fill='tozeroy',
            fillcolor="rgba(229,103,10,0.08)"
        ))

        # --- グラフの表示設定 ---
        fig.update_layout(
            title="テスト進捗 / 不具合検出状況",
            barmode="group",
            bargap=0.4,
            bargroupgap=0.0,
            xaxis=dict(
                type="date" if axis_type == "時間軸" else "category",
                tickmode='array',
                tickvals=tickvals,
                ticktext=ticktexts,
                showgrid=True,
                gridcolor="rgba(200,200,200,0.2)",
                gridwidth=0.5,
                categoryorder="array",
                categoryarray=dates
            ),
            yaxis=dict(
                showgrid=True,
                gridcolor="rgba(200,200,200,0.2)",
                gridwidth=0.5,
                title="件数"
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
            ),
            dragmode=False
        )

        return fig

    @staticmethod
    def make_progress_svg(data: Dict[str, Any], settings: Dict[str, Any], width: int = 120, height: int = 16) -> str:
        """進捗状況のSVGを作成"""
        results = data.get("total", {})
        incompleted = data.get("stats", {}).get("incompleted", 0)
        bar_data = []
        total = sum(results.get(r, 0) for r in settings["test_status"]["results"]) + incompleted
        if total == 0:
            return ""
        for result in settings["test_status"]["results"]:
            value = results.get(result, 0)
            if value > 0:
                bar_data.append((settings["webui"]["bar"]["colors"].get(result, "#ccc"), value))
        if incompleted > 0:
            bar_data.append((settings["webui"]["bar"]["colors"].get(settings["test_status"]["labels"]["not_run"], "#ccc"), incompleted))
        # SVG生成
        svg = f'<svg width="{width}" height="{height}">' \
            + ''.join([
                f'<rect x="{sum(width * v / total for _, v in bar_data[:i])}" y="0" width="{width * value / total}" height="{height}" fill="{color}" />'
                for i, (color, value) in enumerate(bar_data)
            ]) + '</svg>'
        return svg

    @staticmethod
    def create_bug_curve_chart(project_data: Dict[str, Any], settings: Dict[str, Any]) -> Optional[go.Figure]:
        """バグ収束曲線を作成"""
        # エラーとワーニングのあるデータを除外
        filtered_data = [d for d in project_data["gathered_data"] 
                        if "error" not in d and "warning" not in d]
        
        if not filtered_data:
            return None

        # 日付順にデータを集計
        executed_counts = []
        bug_counts = []
        total_executed = 0
        total_bugs = 0
        
        for data in filtered_data:
            if "daily" in data:
                for date, values in sorted(data["daily"].items()):
                    # 完了数を累積
                    total_executed += values.get("完了数", 0)
                    # Failを累積
                    total_bugs += values.get("Fail", 0)
                    executed_counts.append(total_executed)
                    bug_counts.append(total_bugs)
        
        if not executed_counts:
            return None

        # グラフの作成
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=executed_counts,
            y=bug_counts,
            mode='lines+markers',
            name='バグ収束曲線',
            line=dict(
                color='red',
                shape='spline',  # スプライン曲線を使用
                smoothing=0.8    # スムージングの強さを指定（0-1.3）
            ),
            marker=dict(
                size=6,          # マーカーのサイズを小さめに
                color='red'
            ),
            hovertemplate='テスト完了数: %{x}<br>不具合検出数: %{y}<extra></extra>'
        ))

        # レイアウトの設定
        fig.update_layout(
            title='バグ収束曲線',
            xaxis_title='テスト完了数（累積）',
            yaxis_title='不具合検出数（累積）',
            showlegend=True,
            height=400,
            margin=dict(t=50, b=50),
            plot_bgcolor="#FFF",  # PB図と同じ背景色に
            font=dict(
                family="sans-serif",
                size=16
            ),
            dragmode=False
        )

        return fig 