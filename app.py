"""
2026年度 EC部 予実管理ダッシュボード — マルチプロジェクト対応
Stylish Modern Design with #f96cb4 accent
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import re
import requests

# ═════════════════════════════════════════════
# 0. Project Configuration
# ═════════════════════════════════════════════
# 各プロジェクトの表示名と Google スプレッドシートの gid（シート番号）を設定します。
# gid の調べ方: スプレッドシートで対象シートを開き、URL 末尾の #gid=XXXX の数字です。
# 例: https://docs.google.com/spreadsheets/d/.../edit#gid=123456789
#     → gid は 123456789

PROJECTS = {
    "AKIBABROADWAY": {
        "label": "AKIBA BROADWAY",
        "subtitle": "お屋敷公演",
        "gid": "0",
    },
    "VIRTUAL_CAFE": {
        "label": "バーチャルあっとほぉーむカフェ",
        "subtitle": "バーチャル配信",
        "gid": "",
    },
    "NICOLIVE": {
        "label": "ニコ生",
        "subtitle": "ニコニコ生放送",
        "gid": "",
    },
    "AT_LIVE": {
        "label": "あっとライブ",
        "subtitle": "ライブイベント",
        "gid": "",
    },
    "WEDDING": {
        "label": "ウェディング",
        "subtitle": "ウェディング事業",
        "gid": "",
    },
    "PHOTOSHOT": {
        "label": "撮影会",
        "subtitle": "撮影会イベント",
        "gid": "",
    },
}

# ─────────────────────────────────────────────
# Design Tokens
# ─────────────────────────────────────────────
PINK = "#f96cb4"
PINK_LIGHT = "#fdb8d8"
WHITE = "#FFFFFF"
LIGHT_GRAY = "#F8F9FA"
MID_GRAY = "#DEE2E6"
CHARCOAL = "#333333"
DARK = "#1a1a2e"
FONT = "'Noto Sans JP', 'Helvetica Neue', Arial, sans-serif"

MONTHS = ["4月", "5月", "6月", "7月", "8月", "9月",
          "10月", "11月", "12月", "1月", "2月", "3月"]

SALES_CATEGORIES = [
    "物販（会場内チェキ）",
    "物販（会場内グッズ）",
    "ガイドツアー",
    "グッズ（MD）",
    "グッズ通販（MD）",
    "チケット（前売り）",
    "チケット（当日）",
    "チケット（海外）",
    "外部",
]


# ═════════════════════════════════════════════
# 1. Global CSS
# ═════════════════════════════════════════════
def inject_css():
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;600;700&display=swap');

        html, body, p, span, div, td, th, label, input, select, textarea,
        .stMarkdown, .stDataFrame, .stSelectbox, .stTextInput,
        .stCaption, .stButton, [data-testid="stSidebar"] {{
            font-family: {FONT};
            color: {CHARCOAL};
        }}
        /* Streamlit の Material Icons を上書きしない */
        [data-testid="stIconMaterial"],
        .material-symbols-rounded,
        [class*="Icon"] {{
            font-family: 'Material Symbols Rounded' !important;
        }}
        .stApp {{ background: {LIGHT_GRAY}; }}

        /* ── Header ── */
        .dash-header {{
            background: linear-gradient(135deg, {DARK} 0%, #2d2d4e 100%);
            padding: 1.3rem 2rem 1.1rem;
            border-radius: 0 0 1rem 1rem;
            margin: -1rem -1rem 1.5rem -1rem;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}
        .header-left {{
            display: flex;
            align-items: center;
            gap: 1rem;
        }}
        .header-accent {{
            width: 3px; height: 34px;
            background: {PINK};
            border-radius: 2px;
        }}
        .header-left h1 {{
            color: {WHITE};
            font-size: 1.25rem;
            font-weight: 600;
            letter-spacing: .03em;
            margin: 0;
        }}
        .header-left p {{
            color: rgba(255,255,255,.5);
            font-size: .72rem;
            font-weight: 300;
            letter-spacing: .06em;
            margin: 0;
        }}

        /* ── Hamburger ── */
        .hamburger-btn {{
            background: rgba(255,255,255,.1);
            border: 1px solid rgba(255,255,255,.15);
            border-radius: .45rem;
            padding: .5rem .6rem;
            cursor: pointer;
            display: flex;
            flex-direction: column;
            gap: 4px;
            transition: background .2s;
        }}
        .hamburger-btn:hover {{ background: rgba(255,255,255,.2); }}
        .hamburger-line {{
            width: 18px; height: 2px;
            background: {WHITE};
            border-radius: 1px;
        }}

        /* ── KPI Card ── */
        .kpi-card {{
            background: {WHITE};
            border-radius: .7rem;
            padding: 1.3rem 1.4rem;
            box-shadow: 0 1px 4px rgba(0,0,0,.06);
            border: 1px solid rgba(0,0,0,.04);
            transition: box-shadow .2s;
        }}
        .kpi-card:hover {{ box-shadow: 0 4px 16px rgba(0,0,0,.08); }}
        .kpi-label {{
            font-size: .7rem; font-weight: 500;
            letter-spacing: .1em; text-transform: uppercase;
            color: #999; margin-bottom: .45rem;
        }}
        .kpi-value {{
            font-size: 1.55rem; font-weight: 700;
            color: {CHARCOAL}; line-height: 1.15;
        }}
        .kpi-sub {{ font-size: .72rem; color: #aaa; margin-top: .25rem; }}
        .kpi-accent {{ color: {PINK}; }}

        /* ── Section Card ── */
        .section-card {{
            background: {WHITE};
            border-radius: .7rem;
            padding: 1.6rem 1.8rem;
            box-shadow: 0 1px 4px rgba(0,0,0,.06);
            border: 1px solid rgba(0,0,0,.04);
            margin-bottom: 1.2rem;
        }}
        .section-title {{
            font-size: .9rem; font-weight: 600;
            color: {CHARCOAL};
            margin-bottom: 1rem; padding-bottom: .5rem;
            border-bottom: 2px solid {PINK};
            display: inline-block;
        }}

        /* ── Progress bar ── */
        .prog-track {{
            background: {LIGHT_GRAY}; border-radius: 6px;
            height: 7px; overflow: hidden; margin-top: .5rem;
        }}
        .prog-fill {{
            height: 100%; border-radius: 6px;
            background: linear-gradient(90deg, {PINK}, {PINK_LIGHT});
            transition: width .6s ease;
        }}

        /* ── Hide Streamlit chrome ── */
        #MainMenu, footer {{visibility: hidden;}}
        header {{background: transparent !important;}}
        .block-container {{
            padding-top: 0 !important;
            max-width: 1100px;
        }}

        /* ── Sidebar ── */
        [data-testid="stSidebar"] {{ background: {WHITE}; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# ═════════════════════════════════════════════
# 2. Data Helpers
# ═════════════════════════════════════════════
def yen_to_num(v):
    if pd.isna(v) or v in ("", "#DIV/0!"):
        return 0
    s = str(v).replace("¥", "").replace(",", "").replace(" ", "")
    try:
        return int(s)
    except ValueError:
        return 0


def extract_spreadsheet_id(raw_url: str) -> str:
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", raw_url)
    return m.group(1) if m else ""


def build_csv_url(spreadsheet_id: str, gid: str) -> str:
    return (f"https://docs.google.com/spreadsheets/d/"
            f"{spreadsheet_id}/export?format=csv&gid={gid}")


def load_from_gsheet(csv_url: str) -> pd.DataFrame:
    resp = requests.get(csv_url, timeout=15)
    resp.raise_for_status()
    return pd.read_csv(
        io.StringIO(resp.content.decode("utf-8-sig")),
        header=None, dtype=str,
    )


def load_csv(path: str) -> pd.DataFrame:
    return pd.read_csv(path, header=None, dtype=str, encoding="utf-8-sig")


def parse_data(raw: pd.DataFrame) -> dict:
    data = {}

    def row_values(label: str):
        for i in range(raw.shape[0]):
            cell = str(raw.iloc[i, 0]).strip()
            if cell == label:
                return [yen_to_num(raw.iloc[i, j])
                        for j in range(1, min(13, raw.shape[1]))]
        return [0] * 12

    data["budget"] = row_values("売上予算")
    data["sales_incl_tax"] = row_values("売上合計(税込)")
    data["sales_excl_tax"] = row_values("売上合計(税抜)")
    data["last_year"] = row_values("昨年売上")
    data["cost_budget"] = row_values("変動費予算")
    data["cost_incl_tax"] = row_values("原価合計(税込)")
    data["cost_excl_tax"] = row_values("原価合計(税抜)")
    data["sg_and_a"] = row_values("販管費")
    data["operating_profit"] = row_values("営業利益")

    cats = {}
    for cat in SALES_CATEGORIES:
        cats[cat] = row_values(cat)
    data["categories"] = cats
    return data


# ═════════════════════════════════════════════
# 3. Charts
# ═════════════════════════════════════════════
PLOTLY_BASE = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family=FONT, size=12, color=CHARCOAL),
    margin=dict(l=50, r=30, t=40, b=40),
    hoverlabel=dict(bgcolor=WHITE, font_size=12, bordercolor=PINK),
)


def chart_budget_vs_actual(data):
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=MONTHS, y=data["sales_incl_tax"], name="実績（税込）",
        marker_color=PINK, marker_line_width=0, opacity=0.85))
    fig.add_trace(go.Bar(
        x=MONTHS, y=data["last_year"], name="前年実績",
        marker_color=MID_GRAY, marker_line_width=0, opacity=0.55))
    fig.add_trace(go.Scatter(
        x=MONTHS, y=data["budget"], name="売上予算",
        mode="lines+markers",
        line=dict(color=CHARCOAL, width=2, dash="dot"),
        marker=dict(size=5, color=CHARCOAL)))
    fig.update_layout(
        **PLOTLY_BASE, barmode="group",
        legend=dict(orientation="h", yanchor="bottom", y=1.04,
                    xanchor="center", x=0.5, font=dict(size=11)),
        yaxis=dict(gridcolor="rgba(0,0,0,.06)", zeroline=False,
                   tickformat=",", tickprefix="¥"),
        xaxis=dict(showgrid=False), height=370)
    return fig


def chart_donut(data):
    labels = list(data["categories"].keys())
    values = [sum(v) for v in data["categories"].values()]
    n = len(labels)
    colors = [f"rgb({int(249*(1-i/max(n-1,1))+222*(i/max(n-1,1)))},"
              f"{int(108*(1-i/max(n-1,1))+226*(i/max(n-1,1)))},"
              f"{int(180*(1-i/max(n-1,1))+230*(i/max(n-1,1)))})"
              for i in range(n)]
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.58,
        marker=dict(colors=colors, line=dict(color=WHITE, width=2)),
        textinfo="percent", textfont=dict(size=11, color=CHARCOAL),
        hovertemplate="%{label}<br>¥%{value:,}<br>%{percent}<extra></extra>",
        sort=False))
    fig.update_layout(
        **{**PLOTLY_BASE, "margin": dict(l=20, r=160, t=30, b=20)},
        showlegend=True,
        legend=dict(orientation="v", yanchor="middle", y=0.5,
                    xanchor="left", x=1.02, font=dict(size=10.5)),
        height=370)
    return fig


def chart_monthly_profit(data):
    profits = data["operating_profit"]
    colors = [PINK if v >= 0 else "#e74c3c" for v in profits]
    fig = go.Figure(go.Bar(
        x=MONTHS, y=profits, marker_color=colors,
        marker_line_width=0, opacity=0.85))
    fig.update_layout(
        **PLOTLY_BASE, height=300,
        yaxis=dict(gridcolor="rgba(0,0,0,.06)", zeroline=True,
                   zerolinecolor="rgba(0,0,0,.12)",
                   tickformat=",", tickprefix="¥"),
        xaxis=dict(showgrid=False))
    return fig


# ═════════════════════════════════════════════
# 4. UI Components
# ═════════════════════════════════════════════
def fmt_yen(v):
    return f"¥{v:,.0f}"


def render_header(project_key):
    proj = PROJECTS[project_key]
    st.markdown(f"""
        <div class="dash-header">
            <div class="header-left">
                <div class="header-accent"></div>
                <div>
                    <h1>{proj["label"]}</h1>
                    <p>{proj["subtitle"]} &mdash; 2026年度 予実管理</p>
                </div>
            </div>
        </div>""", unsafe_allow_html=True)


def render_kpi_cards(data):
    total_actual = sum(data["sales_incl_tax"])
    total_budget = sum(data["budget"])
    total_lastyear = sum(data["last_year"])
    budget_rate = (total_actual / total_budget * 100) if total_budget else 0
    yoy_rate = (total_actual / total_lastyear * 100) if total_lastyear else 0
    total_profit = sum(data["operating_profit"])
    profit_rate = (total_profit / total_actual * 100) if total_actual else 0

    cols = st.columns(4, gap="medium")
    with cols[0]:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-label">通期累計実績（税込）</div>
            <div class="kpi-value">{fmt_yen(total_actual)}</div>
            <div class="kpi-sub">予算 {fmt_yen(total_budget)}</div>
        </div>""", unsafe_allow_html=True)
    with cols[1]:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-label">予算達成率</div>
            <div class="kpi-value kpi-accent">{budget_rate:.1f}%</div>
            <div class="prog-track"><div class="prog-fill"
                style="width:{min(budget_rate,100):.1f}%"></div></div>
        </div>""", unsafe_allow_html=True)
    with cols[2]:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-label">前年同期比</div>
            <div class="kpi-value">{yoy_rate:.1f}%</div>
            <div class="kpi-sub">前年 {fmt_yen(total_lastyear)}</div>
        </div>""", unsafe_allow_html=True)
    with cols[3]:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-label">営業利益 累計</div>
            <div class="kpi-value">{fmt_yen(total_profit)}</div>
            <div class="kpi-sub">利益率 {profit_rate:.1f}%</div>
        </div>""", unsafe_allow_html=True)


# ═════════════════════════════════════════════
# 5. Main
# ═════════════════════════════════════════════
def main():
    st.set_page_config(
        page_title="予実管理 2026",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    inject_css()

    # ── Sidebar ──
    with st.sidebar:
        st.markdown(
            f'<div style="display:flex;align-items:center;gap:.6rem;'
            f'margin-bottom:1.2rem;">'
            f'<div style="width:3px;height:26px;background:{PINK};'
            f'border-radius:2px;"></div>'
            f'<span style="font-size:.95rem;font-weight:600;'
            f'color:{CHARCOAL};">予実管理 2026</span></div>',
            unsafe_allow_html=True)

        # Project Selector
        st.markdown(
            '<p style="font-size:.68rem;font-weight:500;letter-spacing:.1em;'
            'text-transform:uppercase;color:#999;margin-bottom:.2rem;">'
            'プロジェクト選択</p>', unsafe_allow_html=True)

        project_keys = list(PROJECTS.keys())
        project_labels = {k: v["label"] for k, v in PROJECTS.items()}

        selected_key = st.selectbox(
            "プロジェクト",
            options=project_keys,
            format_func=lambda k: project_labels[k],
            index=0,
            label_visibility="collapsed",
        )

        st.markdown("---")

        # Spreadsheet URL
        st.markdown(
            '<p style="font-size:.68rem;font-weight:500;letter-spacing:.1em;'
            'text-transform:uppercase;color:#999;margin-bottom:.2rem;">'
            'スプレッドシート URL</p>', unsafe_allow_html=True)

        sheet_url = st.text_input(
            "URL",
            value=st.session_state.get("sheet_url", ""),
            placeholder="https://docs.google.com/spreadsheets/d/xxxxx/edit",
            label_visibility="collapsed",
        )
        if sheet_url:
            st.session_state["sheet_url"] = sheet_url

        st.markdown(
            '<p style="font-size:.68rem;color:#aaa;line-height:1.5;'
            'margin-top:.2rem;">'
            '共有 →「リンクを知っている全員：閲覧者」に設定<br>'
            '1つの URL で全プロジェクトを読み込みます</p>',
            unsafe_allow_html=True)

        st.markdown("---")

        # Auto refresh
        auto_sec = st.selectbox(
            "自動更新間隔",
            options=[0, 10, 30, 60],
            format_func=lambda x: "手動のみ" if x == 0 else f"{x}秒ごと",
            index=0,
        )

        st.markdown("---")

        # gid Settings
        with st.expander("各シートの gid 設定"):
            st.markdown(
                '<p style="font-size:.68rem;color:#aaa;line-height:1.5;">'
                '各シートを開いた時の URL 末尾<br>'
                '「#gid=数字」の数字を入力</p>',
                unsafe_allow_html=True)
            for key, proj in PROJECTS.items():
                new_gid = st.text_input(
                    proj["label"],
                    value=st.session_state.get(
                        f"gid_{key}", proj["gid"]),
                    placeholder="例: 0",
                    key=f"gid_input_{key}",
                )
                st.session_state[f"gid_{key}"] = new_gid

    # ── Auto-refresh ──
    if auto_sec > 0:
        st.markdown(
            f'<meta http-equiv="refresh" content="{auto_sec}">',
            unsafe_allow_html=True)

    # ── Header ──
    render_header(selected_key)

    # ── Load Data ──
    raw = None
    load_error = None
    current_gid = st.session_state.get(
        f"gid_{selected_key}", PROJECTS[selected_key]["gid"])

    if sheet_url.strip():
        ssid = extract_spreadsheet_id(sheet_url.strip())
        if ssid and current_gid:
            try:
                raw = load_from_gsheet(build_csv_url(ssid, current_gid))
            except Exception as e:
                load_error = str(e)
        elif ssid and not current_gid:
            load_error = (
                f"「{PROJECTS[selected_key]['label']}」の gid が未設定です。\n"
                f"サイドバーの「各シートの gid 設定」で入力してください。")
        else:
            load_error = "スプレッドシートの URL を認識できませんでした。"

    if raw is None:
        try:
            raw = load_csv("data.csv")
            if load_error:
                st.warning(f"スプレッドシート読み込み失敗。ローカル CSV を表示中。\n\n{load_error}")
        except FileNotFoundError:
            if load_error:
                st.error(load_error)
            else:
                st.error("データがありません。サイドバーで URL を入力してください。")
            st.stop()

    data = parse_data(raw)

    # ── Reload + Status ──
    col_r, col_s, _ = st.columns([1, 3, 2])
    with col_r:
        if st.button("データを再読み込み"):
            st.rerun()
    with col_s:
        if sheet_url.strip() and load_error is None and raw is not None:
            st.caption(
                f"Google Sheets (gid={current_gid}) から取得済み")
        elif not sheet_url.strip():
            st.caption("ローカル CSV を使用中")

    st.markdown("<div style='height:.3rem'></div>", unsafe_allow_html=True)

    # ── KPI ──
    render_kpi_cards(data)
    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

    # ── Charts ──
    left, right = st.columns([3, 2], gap="medium")
    with left:
        st.markdown(
            '<div class="section-card">'
            '<div class="section-title">月別 予算 vs 実績 vs 前年</div>',
            unsafe_allow_html=True)
        st.plotly_chart(chart_budget_vs_actual(data),
                        use_container_width=True,
                        config={"displayModeBar": False})
        st.markdown("</div>", unsafe_allow_html=True)
    with right:
        st.markdown(
            '<div class="section-card">'
            '<div class="section-title">売上内訳</div>',
            unsafe_allow_html=True)
        st.plotly_chart(chart_donut(data),
                        use_container_width=True,
                        config={"displayModeBar": False})
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown(
        '<div class="section-card">'
        '<div class="section-title">月別 営業利益</div>',
        unsafe_allow_html=True)
    st.plotly_chart(chart_monthly_profit(data),
                    use_container_width=True,
                    config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Monthly Summary ──
    st.markdown(
        '<div class="section-card">'
        '<div class="section-title">月別サマリー</div>',
        unsafe_allow_html=True)
    rows = []
    for i, m in enumerate(MONTHS):
        a = data["sales_incl_tax"][i]
        b = data["budget"][i]
        ly = data["last_year"][i]
        rows.append({
            "月": m,
            "売上予算": fmt_yen(b),
            "売上実績": fmt_yen(a),
            "予算達成率": f"{(a/b*100) if b else 0:.1f}%",
            "前年実績": fmt_yen(ly),
            "前年比": f"{(a/ly*100) if ly else 0:.1f}%",
            "営業利益": fmt_yen(data["operating_profit"][i]),
        })
    st.dataframe(pd.DataFrame(rows),
                 hide_index=True, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Category Breakdown ──
    st.markdown(
        '<div class="section-card">'
        '<div class="section-title">カテゴリ別 実績一覧</div>',
        unsafe_allow_html=True)
    cat_rows = []
    for cat in SALES_CATEGORIES:
        row = {"カテゴリ": cat}
        for i, m in enumerate(MONTHS):
            row[m] = fmt_yen(data["categories"][cat][i])
        row["合計"] = fmt_yen(sum(data["categories"][cat]))
        cat_rows.append(row)
    st.dataframe(pd.DataFrame(cat_rows),
                 hide_index=True, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Footer ──
    st.markdown(
        '<div style="text-align:center;padding:2rem 0 1rem;color:#bbb;'
        'font-size:.68rem;letter-spacing:.08em;">'
        '2026 予実管理ダッシュボード &mdash; Prototype</div>',
        unsafe_allow_html=True)


if __name__ == "__main__":
    main()
