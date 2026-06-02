"""
2026年度 EC部 予実管理ダッシュボード — マルチプロジェクト対応
"""
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io, re, requests, datetime, base64
from pathlib import Path

# ═══════════════════════════════════════════════════
# 0. CONFIG
# ═══════════════════════════════════════════════════
PROJECTS = {
    "MAIN":          {"label": "メインダッシュボード", "subtitle": "全プロジェクト概況",      "gid": "",          "logo": None},
    "AKIBABROADWAY": {"label": "AKIBA BROADWAY",      "subtitle": "お屋敷公演",              "gid": "0",         "logo": "assets/akibro.jpg"},
    "VIRTUAL":       {"label": "バーチャルあっとほぉーむカフェ", "subtitle": "バーチャル配信", "gid": "1173568321","logo": "assets/virtual.png"},
    "NICOLIVE":      {"label": "ニコ生",               "subtitle": "ニコニコ生放送",          "gid": "1834964074","logo": "assets/nicolive.png"},
    "AT_LIVE":       {"label": "あっとライブ",          "subtitle": "ライブイベント",          "gid": "1558766351","logo": None},
    "WEDDING":       {"label": "ウェディング",          "subtitle": "ウェディング事業",        "gid": "1014571293","logo": None},
    "PHOTOSHOT":     {"label": "撮影会",               "subtitle": "撮影会イベント",          "gid": "970613346", "logo": None},
}

# ── デフォルトのスプレッドシート URL（毎回入力不要にするため） ──
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1euOCbnz-bC-xqoMe_hePVsDLVVdO4sNaK80woswd1sM/edit"

# ── Color palette (refreshed for better visibility) ──
PINK       = "#f96cb4"   # Brand accent
PINK_L     = "#fdb8d8"   # Soft pink
PURPLE     = "#a78bfa"   # Lavender — previous year
BLUE       = "#60a5fa"   # Sky blue
TEAL       = "#34d399"   # Mint
AMBER      = "#fbbf24"   # Gold
ORANGE     = "#fb923c"   # Peach — cost
ROSE       = "#fb7185"   # Rose
SAGE       = "#a3b18a"   # Sage

PROFIT_POS = "#10b981"   # Emerald (profit)
PROFIT_NEG = "#ef4444"   # Red (loss)
COST_COLOR = "#fb923c"   # Orange (cost)

WHITE      = "#FFFFFF"
LGRAY      = "#F8F9FA"
MGRAY      = "#CBD5E1"   # Slate-200, more visible than DEE2E6
CHARCOAL   = "#1f2937"
DARK       = "#1a1a2e"

# Palette for multi-category donut / stacked charts
PALETTE = [PINK, PURPLE, BLUE, TEAL, AMBER, ORANGE, ROSE, SAGE, "#94a3b8"]

FONT   = "'Noto Sans JP','Helvetica Neue',Arial,sans-serif"
MONTHS = ["4月","5月","6月","7月","8月","9月","10月","11月","12月","1月","2月","3月"]

def fiscal_month_index():
    m = datetime.date.today().month
    return m - 4 if m >= 4 else m + 8


# ═══════════════════════════════════════════════════
# 1. CSS
# ═══════════════════════════════════════════════════
def inject_css():
    st.markdown(f"""<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;600;700&display=swap');
    html,body,p,span,div,td,th,label,input,select,textarea,
    .stMarkdown,.stDataFrame,.stSelectbox,.stTextInput,
    .stCaption,.stButton,[data-testid="stSidebar"]{{font-family:{FONT};color:{CHARCOAL};}}
    [data-testid="stIconMaterial"],.material-symbols-rounded,[class*="Icon"]{{font-family:'Material Symbols Rounded'!important;}}
    .stApp{{background:{LGRAY};}}

    /* ── Header ── */
    .dash-header{{background:linear-gradient(135deg,{DARK},#2d2d4e);
        padding:1.3rem 1.8rem 1.2rem;border-radius:0 0 1rem 1rem;
        margin:-1rem -1rem 1.2rem -1rem;display:flex;align-items:center;gap:1.2rem;}}
    .header-accent{{width:4px;height:54px;background:{PINK};border-radius:2px;flex-shrink:0;}}
    .dash-header h1{{color:{WHITE};font-size:1.4rem;font-weight:700;margin:0;letter-spacing:.02em;}}
    .dash-header p{{color:#e5e7eb;font-size:.85rem;font-weight:400;margin:.2rem 0 0;letter-spacing:.03em;}}

    /* ── Header logo wrapper ── */
    .header-logo-wrap{{background:{WHITE};border-radius:10px;
        padding:8px 14px;display:flex;align-items:center;justify-content:center;
        box-shadow:0 2px 8px rgba(0,0,0,.25);flex-shrink:0;
        height:62px;min-width:80px;max-width:240px;}}
    .header-logo{{height:46px;max-width:212px;object-fit:contain;display:block;}}

    /* ── Project card logo (main dashboard) ── */
    .proj-card-header{{display:inline-flex;align-items:center;gap:.8rem;
        margin:0 0 .8rem;padding-bottom:.3rem;border-bottom:2px solid {PINK};}}
    .proj-logo-wrap{{background:{LGRAY};border-radius:6px;
        padding:5px 10px;display:flex;align-items:center;justify-content:center;
        height:46px;min-width:60px;max-width:160px;border:1px solid rgba(0,0,0,.05);}}
    .proj-logo{{height:36px;max-width:140px;object-fit:contain;display:block;}}
    .proj-badge{{width:42px;height:42px;border-radius:50%;background:{PINK};
        color:{WHITE};display:flex;align-items:center;justify-content:center;
        font-weight:700;font-size:1.1rem;flex-shrink:0;}}
    .proj-card-title{{font-size:1rem;font-weight:600;color:{CHARCOAL};margin:0;}}

    /* ── Month selector bar ── */
    .month-bar{{background:{WHITE};border-radius:.5rem;padding:.4rem .9rem;
        box-shadow:0 1px 3px rgba(0,0,0,.05);border:1px solid rgba(0,0,0,.04);
        margin-bottom:1rem;display:flex;align-items:center;gap:.5rem;}}
    .month-bar-label{{font-size:.7rem;color:#999;letter-spacing:.05em;
        text-transform:uppercase;font-weight:500;}}

    /* ── KPI Card ── */
    .kpi-card{{background:{WHITE};border-radius:.65rem;padding:1rem 1.15rem;
        box-shadow:0 1px 4px rgba(0,0,0,.06);border:1px solid rgba(0,0,0,.04);
        transition:box-shadow .2s;height:100%;}}
    .kpi-card:hover{{box-shadow:0 4px 16px rgba(0,0,0,.08);}}
    .kpi-label{{font-size:.65rem;font-weight:500;letter-spacing:.1em;
        text-transform:uppercase;color:#999;margin-bottom:.35rem;}}
    .kpi-value{{font-size:1.35rem;font-weight:700;color:{CHARCOAL};line-height:1.15;}}
    .kpi-sub{{font-size:.7rem;color:#aaa;margin-top:.2rem;}}
    .kpi-accent{{color:{PINK};}}
    .kpi-cost{{color:{COST_COLOR};}}
    .kpi-profit-pos{{color:{PROFIT_POS};}}
    .kpi-profit-neg{{color:{PROFIT_NEG};}}

    /* ── Section Card ── */
    .section-card{{background:{WHITE};border-radius:.65rem;padding:1.4rem 1.6rem;
        box-shadow:0 1px 4px rgba(0,0,0,.06);border:1px solid rgba(0,0,0,.04);
        margin-bottom:1rem;}}
    .section-title{{font-size:.85rem;font-weight:600;color:{CHARCOAL};
        margin-bottom:.9rem;padding-bottom:.4rem;border-bottom:2px solid {PINK};
        display:inline-block;}}

    /* ── Progress bar ── */
    .prog-track{{background:{LGRAY};border-radius:6px;height:7px;
        overflow:hidden;margin-top:.4rem;}}
    .prog-fill{{height:100%;border-radius:6px;
        background:linear-gradient(90deg,{PINK},{PINK_L});transition:width .6s;}}
    .prog-fill-over{{background:linear-gradient(90deg,{PROFIT_POS},#34d399);}}

    /* ── Main page summary cards ── */
    .proj-summary-card{{background:{WHITE};border-radius:.65rem;
        padding:1.2rem 1.4rem;box-shadow:0 1px 4px rgba(0,0,0,.06);
        border:1px solid rgba(0,0,0,.04);margin-bottom:.7rem;}}
    .proj-summary-card h3{{font-size:.92rem;font-weight:600;color:{CHARCOAL};
        margin:0 0 .65rem;padding-bottom:.3rem;border-bottom:2px solid {PINK};
        display:inline-block;}}
    .proj-row{{display:grid;grid-template-columns:repeat(5,1fr);gap:1rem;}}
    .proj-metric-label{{font-size:.62rem;color:#999;letter-spacing:.05em;
        text-transform:uppercase;font-weight:500;}}
    .proj-metric-value{{font-size:1.05rem;font-weight:700;color:{CHARCOAL};
        margin-top:.15rem;}}

    /* ── Hide Streamlit chrome ── */
    #MainMenu,footer{{visibility:hidden;}}
    header{{background:transparent!important;}}
    .block-container{{padding-top:0!important;max-width:1100px;}}
    [data-testid="stSidebar"]{{background:{WHITE};}}
    </style>""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════
# 2. DATA HELPERS
# ═══════════════════════════════════════════════════
def yen(v):
    if pd.isna(v) or str(v).strip() in ("","#DIV/0!","#VALUE!"):
        return 0
    s = str(v).replace("¥","").replace(",","").replace(" ","").replace("\u3000","")
    s = s.replace("−","-").replace("ー","-")
    try:
        return int(s)
    except ValueError:
        try: return int(float(s))
        except ValueError: return 0

def fmt(v): return f"¥{v:,.0f}"

def html(s):
    """Strip newlines/indents so Streamlit's markdown doesn't treat it as a code block."""
    return re.sub(r"\s*\n\s*", "", s)

@st.cache_data
def img_b64(path):
    """画像ファイルを base64 データ URI に変換。存在しなければ None。"""
    if not path:
        return None
    # app.py と同じディレクトリを起点とする絶対パスで解決
    base_dir = Path(__file__).parent
    p = (base_dir / path) if not Path(path).is_absolute() else Path(path)
    if not p.exists():
        return None
    with open(p, "rb") as f:
        magic = f.read(12)
    if magic.startswith(b"\xff\xd8\xff"):
        mime = "image/jpeg"
    elif magic.startswith(b"\x89PNG"):
        mime = "image/png"
    elif magic.startswith(b"GIF8"):
        mime = "image/gif"
    elif magic[:4] == b"RIFF" and magic[8:12] == b"WEBP":
        mime = "image/webp"
    else:
        mime = "image/png"
    with open(p, "rb") as f:
        data = base64.b64encode(f.read()).decode()
    return f"data:{mime};base64,{data}"

def extract_ssid(url):
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    return m.group(1) if m else ""

def csv_url(ssid, gid): return f"https://docs.google.com/spreadsheets/d/{ssid}/export?format=csv&gid={gid}"

def fetch_sheet(ssid, gid):
    r = requests.get(csv_url(ssid, gid), timeout=15)
    r.raise_for_status()
    return pd.read_csv(io.StringIO(r.content.decode("utf-8-sig")), header=None, dtype=str)

def load_local():
    return pd.read_csv("data.csv", header=None, dtype=str, encoding="utf-8-sig")

def rv(raw, label):
    """Find row by first column label, return 12 monthly values."""
    for i in range(raw.shape[0]):
        if str(raw.iloc[i,0]).strip() == label:
            return [yen(raw.iloc[i,j]) for j in range(1, min(13, raw.shape[1]))]
    return [0]*12

def parse_common(raw):
    return {
        "budget":     rv(raw, "売上予算"),
        "sales_inc":  rv(raw, "売上合計(税込)"),
        "sales_exc":  rv(raw, "売上合計(税抜)"),
        "last_year":  rv(raw, "昨年売上"),
        "cost_budget":rv(raw, "変動費予算"),
        "cost_inc":   rv(raw, "原価合計(税込)"),
        "cost_exc":   rv(raw, "原価合計(税抜)"),
        "sg_a":       rv(raw, "販管費"),
        "profit":     rv(raw, "営業利益"),
    }

def parse_akibro_cats(raw):
    cats = ["物販（会場内チェキ）","物販（会場内グッズ）","ガイドツアー",
            "グッズ（MD）","グッズ通販（MD）","チケット（前売り）",
            "チケット（当日）","チケット（海外）","外部"]
    return {c: rv(raw, c) for c in cats}

def parse_virtual_cats(raw):
    cats = {}
    coin_idx = 0
    pf_names = ["SBペイメント","Apple","Google"]
    prepaid = {}
    apple_google_count = 0
    unused_end = {}
    ue_count = 0
    ue_names_coin = ["SBペイメント","Apple","Google"]
    for i in range(raw.shape[0]):
        cell = str(raw.iloc[i,0]).strip()
        vals = [yen(raw.iloc[i,j]) for j in range(1, min(13, raw.shape[1]))]
        if cell == "売上（＝コイン使用数）" and coin_idx < 3:
            cats[pf_names[coin_idx]] = vals
            coin_idx += 1
        elif cell == "売上（＝チケット使用数）":
            cats["Stripe"] = vals
        elif cell == "グッズ":
            cats["グッズ"] = vals
        elif cell == "ボイス":
            cats["ボイス"] = vals
        elif cell == "コイン購入額（前受金）":
            prepaid["SBペイメント"] = vals
        elif "コイン購入額" in cell and "手数料" in cell and apple_google_count < 2:
            prepaid[["Apple","Google"][apple_google_count]] = vals
            apple_google_count += 1
        elif cell == "チケット購入額（前受金）":
            prepaid["Stripe"] = vals
        elif "当月末" in cell and "コイン" in cell and ue_count < 3:
            unused_end[ue_names_coin[ue_count]] = vals
            ue_count += 1
        elif "当月末" in cell and "チケット" in cell:
            unused_end["Stripe"] = vals
    return cats, prepaid, unused_end

def parse_nicolive_cats(raw):
    cats = ["会員費","ギフト収入","ファンミ参加費(Bitfan)","ファンミチェキ券"]
    return {c: rv(raw, c) for c in cats}

def parse_atlive_cats(raw):
    cats = ["チケット(前売り)","チケット(当日)","チェキ券"]
    return {c: rv(raw, c) for c in cats}

def parse_wedding_cats(raw):
    cats = ["出張お給仕","フォト"]
    return {c: rv(raw, c) for c in cats}


# ═══════════════════════════════════════════════════
# 3. CHART BUILDERS
# ═══════════════════════════════════════════════════
PL = dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
          font=dict(family=FONT, size=12, color=CHARCOAL),
          margin=dict(l=50,r=30,t=40,b=40),
          hoverlabel=dict(bgcolor=WHITE, font_size=12, bordercolor=PINK))

def _hex_to_rgb(h):
    return int(h[1:3],16), int(h[3:5],16), int(h[5:7],16)

def _hl_colors(base, n, hi=None, fade=.55):
    """Color list; non-highlighted bars dimmed at the given opacity."""
    if hi is None:
        return [base]*n
    r,g,b = _hex_to_rgb(base)
    return [base if i==hi else f"rgba({r},{g},{b},{fade})" for i in range(n)]

def chart_bar_line(data, hi=None):
    """Budget(line) vs Actual(bar) vs LastYear(bar)."""
    n = len(MONTHS)
    fig = go.Figure()
    fig.add_trace(go.Bar(x=MONTHS, y=data["sales_inc"], name="実績（税込）",
                         marker_color=_hl_colors(PINK, n, hi), marker_line_width=0))
    fig.add_trace(go.Bar(x=MONTHS, y=data["last_year"], name="前年実績",
                         marker_color=_hl_colors(PURPLE, n, hi, fade=.45), marker_line_width=0, opacity=.85))
    fig.add_trace(go.Scatter(x=MONTHS, y=data["budget"], name="売上予算",
                             mode="lines+markers",
                             line=dict(color=CHARCOAL, width=2.5, dash="dot"),
                             marker=dict(size=6, color=CHARCOAL)))
    fig.update_layout(**PL, barmode="group", height=370,
        legend=dict(orientation="h",yanchor="bottom",y=1.04,xanchor="center",x=.5,font=dict(size=11)),
        yaxis=dict(gridcolor="rgba(0,0,0,.06)",zeroline=False,tickformat=",",tickprefix="¥"),
        xaxis=dict(showgrid=False))
    return fig

def chart_sales_cost(data, hi=None):
    """売上 vs 原価 のグループ棒グラフ (with profit line)."""
    n = len(MONTHS)
    fig = go.Figure()
    fig.add_trace(go.Bar(x=MONTHS, y=data["sales_inc"], name="売上（税込）",
                         marker_color=_hl_colors(PINK, n, hi), marker_line_width=0))
    fig.add_trace(go.Bar(x=MONTHS, y=data["cost_inc"], name="原価（税込）",
                         marker_color=_hl_colors(COST_COLOR, n, hi, fade=.5),
                         marker_line_width=0, opacity=.85))
    fig.add_trace(go.Scatter(x=MONTHS, y=data["profit"], name="営業利益",
                             mode="lines+markers",
                             line=dict(color=PROFIT_POS, width=2.5),
                             marker=dict(size=6, color=PROFIT_POS)))
    fig.update_layout(**PL, barmode="group", height=370,
        legend=dict(orientation="h",yanchor="bottom",y=1.04,xanchor="center",x=.5,font=dict(size=11)),
        yaxis=dict(gridcolor="rgba(0,0,0,.06)",zeroline=True,zerolinecolor="rgba(0,0,0,.1)",
                   tickformat=",",tickprefix="¥"),
        xaxis=dict(showgrid=False))
    return fig

def chart_donut(categories, mi=None):
    """Donut chart - either total or specific month."""
    labels = list(categories.keys())
    if mi is None:
        values = [sum(v) for v in categories.values()]
    else:
        values = [v[mi] for v in categories.values()]
    n = len(labels)
    colors = [PALETTE[i % len(PALETTE)] for i in range(n)]
    fig = go.Figure(go.Pie(labels=labels,values=values,hole=.58,
        marker=dict(colors=colors,line=dict(color=WHITE,width=2)),
        textinfo="percent",textfont=dict(size=11,color=WHITE),
        hovertemplate="%{label}<br>¥%{value:,}<br>%{percent}<extra></extra>",sort=False))
    fig.update_layout(**{**PL,"margin":dict(l=20,r=160,t=30,b=20)},showlegend=True,
        legend=dict(orientation="v",yanchor="middle",y=.5,xanchor="left",x=1.02,font=dict(size=10)),
        height=360)
    return fig

def chart_achievement_monthly(data, hi=None):
    """月別 予算達成率 棒グラフ - 目立つようにメインエリアに置く用."""
    n = len(MONTHS)
    budget = data["budget"]; actual = data["sales_inc"]
    rates = [(actual[i]/budget[i]*100 if budget[i] else 0) for i in range(n)]
    has_budget = [budget[i] > 0 for i in range(n)]

    colors = []
    for i in range(n):
        if not has_budget[i]:
            colors.append("#e5e7eb")  # No budget — neutral gray
        elif rates[i] >= 100:
            base = PROFIT_POS
            if hi is not None and i != hi:
                r,g,b = _hex_to_rgb(base)
                colors.append(f"rgba({r},{g},{b},.55)")
            else:
                colors.append(base)
        else:
            base = PINK
            if hi is not None and i != hi:
                r,g,b = _hex_to_rgb(base)
                colors.append(f"rgba({r},{g},{b},.55)")
            else:
                colors.append(base)

    fig = go.Figure(go.Bar(x=MONTHS, y=rates, marker_color=colors,
                           marker_line_width=0,
                           text=[f"{r:.0f}%" if has_budget[i] else "—"
                                 for i,r in enumerate(rates)],
                           textposition="outside",
                           textfont=dict(size=10,color=CHARCOAL)))
    fig.add_hline(y=100, line_dash="dot", line_color=CHARCOAL, line_width=1.5,
                  annotation_text="目標 100%", annotation_position="top right",
                  annotation_font_size=10, annotation_font_color=CHARCOAL)
    fig.update_layout(**PL, height=300,
        yaxis=dict(gridcolor="rgba(0,0,0,.06)", zeroline=False, ticksuffix="%",
                   range=[0, max(max(rates)*1.25, 130)]),
        xaxis=dict(showgrid=False))
    return fig

def chart_profit(data, hi=None):
    """月別 営業利益 棒グラフ."""
    p = data["profit"]
    n = len(MONTHS)
    if hi is None:
        colors = [PROFIT_POS if v>=0 else PROFIT_NEG for v in p]
    else:
        colors = []
        for i,v in enumerate(p):
            base = PROFIT_POS if v>=0 else PROFIT_NEG
            if i == hi:
                colors.append(base)
            else:
                r,g,b = _hex_to_rgb(base)
                colors.append(f"rgba({r},{g},{b},.55)")
    fig = go.Figure(go.Bar(x=MONTHS,y=p,marker_color=colors,marker_line_width=0))
    fig.update_layout(**PL,height=300,
        yaxis=dict(gridcolor="rgba(0,0,0,.06)",zeroline=True,
                   zerolinecolor="rgba(0,0,0,.12)",tickformat=",",tickprefix="¥"),
        xaxis=dict(showgrid=False))
    return fig

def chart_stacked_bar(categories, hi=None):
    """Stacked bar with per-month highlighting via opacity."""
    n_months = len(MONTHS)
    fig = go.Figure()
    for idx, (name, vals) in enumerate(categories.items()):
        color = PALETTE[idx % len(PALETTE)]
        # If hi is set, fade non-highlighted columns (whole stack)
        if hi is None:
            colors = [color]*n_months
        else:
            r,g,b = _hex_to_rgb(color)
            colors = [color if i==hi else f"rgba({r},{g},{b},.5)" for i in range(n_months)]
        fig.add_trace(go.Bar(x=MONTHS, y=vals, name=name, marker_color=colors, marker_line_width=0))
    fig.update_layout(**PL, barmode="stack", height=370,
        legend=dict(orientation="h",yanchor="bottom",y=1.04,xanchor="center",x=.5,font=dict(size=10)),
        yaxis=dict(gridcolor="rgba(0,0,0,.06)",zeroline=False,tickformat=",",tickprefix="¥"),
        xaxis=dict(showgrid=False))
    return fig

def chart_profit_with_margin(data, hi=None):
    """営業利益バー + 利益率ライン."""
    p = data["profit"]
    s = data["sales_inc"]
    margins = [(p[i]/s[i]*100 if s[i] else 0) for i in range(12)]
    n = len(MONTHS)
    if hi is None:
        colors = [PROFIT_POS if v>=0 else PROFIT_NEG for v in p]
    else:
        colors = []
        for i,v in enumerate(p):
            base = PROFIT_POS if v>=0 else PROFIT_NEG
            if i == hi: colors.append(base)
            else:
                r,g,b = _hex_to_rgb(base)
                colors.append(f"rgba({r},{g},{b},.55)")
    fig = go.Figure()
    fig.add_trace(go.Bar(x=MONTHS,y=p,name="営業利益",marker_color=colors,marker_line_width=0))
    fig.add_trace(go.Scatter(x=MONTHS,y=margins,name="利益率(%)",
        mode="lines+markers",line=dict(color=CHARCOAL,width=2),
        marker=dict(size=5,color=CHARCOAL),yaxis="y2"))
    fig.update_layout(**PL,height=320,
        yaxis=dict(gridcolor="rgba(0,0,0,.06)",zeroline=True,
                   zerolinecolor="rgba(0,0,0,.12)",tickformat=",",tickprefix="¥"),
        yaxis2=dict(overlaying="y",side="right",showgrid=False,ticksuffix="%",zeroline=False),
        xaxis=dict(showgrid=False),
        legend=dict(orientation="h",yanchor="bottom",y=1.04,xanchor="center",x=.5,font=dict(size=11)))
    return fig

def chart_cumulative(data):
    s = data["sales_inc"]
    cum = []; total = 0
    for v in s:
        total += v; cum.append(total)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=MONTHS,y=cum,mode="lines+markers+text",
        line=dict(color=PINK,width=2.5),marker=dict(size=6,color=PINK),
        text=[fmt(v) if v>0 else "" for v in cum],textposition="top center",
        textfont=dict(size=9,color=CHARCOAL)))
    fig.update_layout(**PL,height=300,
        yaxis=dict(gridcolor="rgba(0,0,0,.06)",zeroline=False,tickformat=",",tickprefix="¥"),
        xaxis=dict(showgrid=False))
    return fig


# ═══════════════════════════════════════════════════
# 4. UI COMPONENTS
# ═══════════════════════════════════════════════════
def render_header(key):
    """シンプルで確実に表示されるヘッダー（白背景・ピンクアクセント）"""
    p = PROJECTS[key]

    # ロゴパスを解決
    logo_path = None
    if p.get("logo"):
        base_dir = Path(__file__).parent
        candidate = base_dir / p["logo"]
        if candidate.exists():
            logo_path = str(candidate)

    # 上部にピンクのアクセントライン
    st.markdown(
        f'<div style="height:4px;background:{PINK};margin:-1rem -1rem 0 -1rem;'
        f'border-radius:0;"></div>',
        unsafe_allow_html=True,
    )

    # ヘッダー本体: 白背景にロゴ + タイトル
    st.markdown(
        f'<div style="background:#ffffff;padding:1.2rem 1.5rem;'
        f'margin:0 -1rem 1.2rem -1rem;border-bottom:1px solid #e5e7eb;'
        f'box-shadow:0 2px 4px rgba(0,0,0,.04);">',
        unsafe_allow_html=True,
    )

    if logo_path:
        c1, c2 = st.columns([1, 5], gap="medium")
        with c1:
            st.image(logo_path, width=130)
        with c2:
            st.markdown(
                f'<div style="padding-top:.4rem;">'
                f'<div style="color:#1f2937;font-size:1.45rem;font-weight:700;'
                f'line-height:1.2;letter-spacing:.02em;">{p["label"]}</div>'
                f'<div style="color:#6b7280;font-size:.9rem;font-weight:400;'
                f'margin-top:.35rem;line-height:1.3;">{p["subtitle"]} — 2026年度 予実管理</div>'
                f'</div>',
                unsafe_allow_html=True,
            )
    else:
        st.markdown(
            f'<div style="color:#1f2937;font-size:1.45rem;font-weight:700;'
            f'line-height:1.2;letter-spacing:.02em;">{p["label"]}</div>'
            f'<div style="color:#6b7280;font-size:.9rem;font-weight:400;'
            f'margin-top:.35rem;line-height:1.3;">{p["subtitle"]} — 2026年度 予実管理</div>',
            unsafe_allow_html=True,
        )

    st.markdown('</div>', unsafe_allow_html=True)

def page_month_selector(default_index=0, key_suffix=""):
    """ページ内の月セレクター。0=全期間、1〜12=各月のインデックス+1。Returns mi (None or 0-11)."""
    opts = ["全期間"] + MONTHS
    col_l, col_r = st.columns([3, 2])
    with col_l:
        st.markdown('<div style="display:flex;align-items:center;height:42px;">'
                    f'<span style="color:#888;font-size:.8rem;">▸ 表示する期間を選択</span></div>',
                    unsafe_allow_html=True)
    with col_r:
        sel = st.selectbox("表示月", opts, index=default_index,
                          label_visibility="collapsed", key=f"month_sel_{key_suffix}")
    return None if sel == "全期間" else MONTHS.index(sel)

def kpi(label, value, sub="", color_class=""):
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    cls = f' kpi-{color_class}' if color_class else ""
    st.markdown(html(f"""<div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value{cls}">{value}</div>{sub_html}
    </div>"""), unsafe_allow_html=True)

def kpi_progress(label, value, pct, sub=""):
    fill_cls = "prog-fill prog-fill-over" if pct > 100 else "prog-fill"
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    st.markdown(html(f"""<div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value kpi-accent">{value}</div>
        <div class="prog-track"><div class="{fill_cls}" style="width:{min(pct,100):.1f}%"></div></div>
        {sub_html}
    </div>"""), unsafe_allow_html=True)

def section(title):
    st.markdown(f'<div class="section-card"><div class="section-title">{title}</div>',
                unsafe_allow_html=True)

def section_end():
    st.markdown('</div>', unsafe_allow_html=True)

def profit_class(v):
    return "profit-pos" if v >= 0 else "profit-neg"


# ═══════════════════════════════════════════════════
# 5. KPI ROW BUILDERS (project-specific)
# ═══════════════════════════════════════════════════
def kpi_row_with_budget(data, mi, has_yoy=True):
    """予算ありプロジェクト用 (AKIBA, AT_LIVE) — 売上/原価/営業利益/予算達成率."""
    if mi is not None:
        s,c,b,ly,pr = (data["sales_inc"][mi], data["cost_inc"][mi],
                       data["budget"][mi], data["last_year"][mi], data["profit"][mi])
        sfx = f"（{MONTHS[mi]}）"
    else:
        s,c,b,ly,pr = (sum(data["sales_inc"]), sum(data["cost_inc"]),
                       sum(data["budget"]), sum(data["last_year"]), sum(data["profit"]))
        sfx = "（通期）"

    cost_rate = (c/s*100) if s else 0
    margin    = (pr/s*100) if s else 0
    ach       = (s/b*100) if b else 0

    cols = st.columns(4, gap="medium")
    with cols[0]: kpi(f"売上実績{sfx}", fmt(s), f"予算 {fmt(b)}" if b else "予算未設定")
    with cols[1]: kpi(f"原価{sfx}", fmt(c), f"原価率 {cost_rate:.1f}%", color_class="cost")
    with cols[2]:
        cls = "profit-pos" if pr >= 0 else "profit-neg"
        kpi(f"営業利益{sfx}", fmt(pr), f"利益率 {margin:.1f}%", color_class=cls)
    with cols[3]:
        if b:
            yoy_sub = f"前年比 {(s/ly*100):.1f}%" if (has_yoy and ly) else ""
            kpi_progress(f"予算達成率{sfx}", f"{ach:.1f}%", ach, yoy_sub)
        else:
            yoy_sub = f"前年 {fmt(ly)}" if ly else ""
            kpi(f"前年比{sfx}", f"{(s/ly*100):.1f}%" if ly else "—", yoy_sub)

def kpi_row_no_budget(data, mi, extra_label=None, extra_value=None, extra_sub=""):
    """予算なしプロジェクト用 (VIRTUAL, NICOLIVE, WEDDING) — 売上/原価/営業利益/その他."""
    if mi is not None:
        s,c,pr = data["sales_inc"][mi], data["cost_inc"][mi], data["profit"][mi]
        sfx = f"（{MONTHS[mi]}）"
    else:
        s,c,pr = sum(data["sales_inc"]), sum(data["cost_inc"]), sum(data["profit"])
        sfx = "（通期）"

    cost_rate = (c/s*100) if s else 0
    margin    = (pr/s*100) if s else 0

    cols = st.columns(4, gap="medium")
    with cols[0]: kpi(f"売上実績{sfx}", fmt(s))
    with cols[1]: kpi(f"原価{sfx}", fmt(c), f"原価率 {cost_rate:.1f}%", color_class="cost")
    with cols[2]:
        cls = "profit-pos" if pr >= 0 else "profit-neg"
        kpi(f"営業利益{sfx}", fmt(pr), f"利益率 {margin:.1f}%", color_class=cls)
    with cols[3]:
        if extra_label:
            kpi(extra_label, extra_value, extra_sub)
        else:
            kpi("—", "—")


# ═══════════════════════════════════════════════════
# 6. PAGE RENDERERS
# ═══════════════════════════════════════════════════

# ── MAIN PAGE ──
def render_main_page(ssid):
    render_header("MAIN")

    # Month selector — defaults to current fiscal month
    mi_default = fiscal_month_index() + 1  # +1 because index 0 is "全期間"
    mi = page_month_selector(default_index=mi_default, key_suffix="main")
    if mi is None:
        period_label = "通期"
    else:
        period_label = MONTHS[mi]

    # Load all projects
    all_data = {}
    for key, proj in PROJECTS.items():
        if key == "MAIN": continue
        gid = st.session_state.get(f"gid_{key}", proj["gid"])
        if not gid or not ssid: continue
        try:
            raw = fetch_sheet(ssid, gid)
            all_data[key] = parse_common(raw)
        except Exception:
            pass

    if not all_data:
        st.info("スプレッドシートの URL と各シートの gid を設定してください。")
        return

    # Per-project summary cards
    total_sales = total_cost = total_profit = 0
    for key, d in all_data.items():
        p = PROJECTS[key]
        if mi is not None:
            s,c,pr,b = d["sales_inc"][mi], d["cost_inc"][mi], d["profit"][mi], d["budget"][mi]
        else:
            s,c,pr,b = sum(d["sales_inc"]), sum(d["cost_inc"]), sum(d["profit"]), sum(d["budget"])
        total_sales += s; total_cost += c; total_profit += pr
        ach = (s/b*100) if b else 0
        margin = (pr/s*100) if s else 0
        pr_color = PROFIT_POS if pr >= 0 else PROFIT_NEG

        ach_html = (f'<div><div class="proj-metric-label">予算達成率</div>'
                    f'<div class="proj-metric-value">{ach:.1f}%</div></div>'
                    if b else
                    f'<div><div class="proj-metric-label">予算</div>'
                    f'<div class="proj-metric-value" style="color:#bbb;">—</div></div>')

        # Header with logo or badge
        logo_b64 = img_b64(p.get("logo"))
        if logo_b64:
            head_inner = (f'<div class="proj-logo-wrap"><img src="{logo_b64}" class="proj-logo" alt=""></div>'
                          f'<div class="proj-card-title">{p["label"]}</div>')
        else:
            initial = p["label"][:1]
            head_inner = (f'<div class="proj-badge">{initial}</div>'
                          f'<div class="proj-card-title">{p["label"]}</div>')

        st.markdown(html(f"""<div class="proj-summary-card">
            <div class="proj-card-header">{head_inner}</div>
            <div class="proj-row">
                <div>
                    <div class="proj-metric-label">売上（税込）</div>
                    <div class="proj-metric-value">{fmt(s)}</div>
                </div>
                <div>
                    <div class="proj-metric-label">原価（税込）</div>
                    <div class="proj-metric-value" style="color:{COST_COLOR};">{fmt(c)}</div>
                </div>
                <div>
                    <div class="proj-metric-label">営業利益</div>
                    <div class="proj-metric-value" style="color:{pr_color};">{fmt(pr)}</div>
                </div>
                <div>
                    <div class="proj-metric-label">利益率</div>
                    <div class="proj-metric-value" style="color:{pr_color};">{margin:.1f}%</div>
                </div>
                {ach_html}
            </div>
        </div>"""), unsafe_allow_html=True)

    # Total
    pr_color = PROFIT_POS if total_profit >= 0 else PROFIT_NEG
    total_margin = (total_profit/total_sales*100) if total_sales else 0
    st.markdown(html(f"""<div class="proj-summary-card" style="background:linear-gradient(135deg,{DARK},#2d2d4e);margin-top:.5rem;">
        <div class="proj-card-header" style="border-color:{PINK};">
            <div class="proj-card-title" style="color:{WHITE};font-size:1.05rem;">全プロジェクト合計（{period_label}）</div>
        </div>
        <div class="proj-row">
            <div>
                <div class="proj-metric-label" style="color:#e5e7eb;">売上合計</div>
                <div class="proj-metric-value" style="color:{WHITE};">{fmt(total_sales)}</div>
            </div>
            <div>
                <div class="proj-metric-label" style="color:#e5e7eb;">原価合計</div>
                <div class="proj-metric-value" style="color:{COST_COLOR};">{fmt(total_cost)}</div>
            </div>
            <div>
                <div class="proj-metric-label" style="color:#e5e7eb;">営業利益合計</div>
                <div class="proj-metric-value" style="color:{pr_color};">{fmt(total_profit)}</div>
            </div>
            <div>
                <div class="proj-metric-label" style="color:#e5e7eb;">利益率</div>
                <div class="proj-metric-value" style="color:{pr_color};">{total_margin:.1f}%</div>
            </div>
            <div></div>
        </div>
    </div>"""), unsafe_allow_html=True)

    # Comparison chart — sales vs cost vs profit by project
    section(f"プロジェクト別 売上 / 原価 / 営業利益（{period_label}）")
    names = [PROJECTS[k]["label"] for k in all_data]
    if mi is not None:
        sales_v = [all_data[k]["sales_inc"][mi] for k in all_data]
        cost_v  = [all_data[k]["cost_inc"][mi]  for k in all_data]
        profit_v= [all_data[k]["profit"][mi]    for k in all_data]
    else:
        sales_v = [sum(all_data[k]["sales_inc"]) for k in all_data]
        cost_v  = [sum(all_data[k]["cost_inc"])  for k in all_data]
        profit_v= [sum(all_data[k]["profit"])    for k in all_data]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=names,y=sales_v,name="売上（税込）",marker_color=PINK,marker_line_width=0))
    fig.add_trace(go.Bar(x=names,y=cost_v,name="原価（税込）",marker_color=COST_COLOR,marker_line_width=0,opacity=.85))
    fig.add_trace(go.Bar(x=names,y=profit_v,name="営業利益",
        marker_color=[PROFIT_POS if v>=0 else PROFIT_NEG for v in profit_v],
        marker_line_width=0,opacity=.85))
    fig.update_layout(**PL,barmode="group",height=380,
        legend=dict(orientation="h",yanchor="bottom",y=1.04,xanchor="center",x=.5,font=dict(size=11)),
        yaxis=dict(gridcolor="rgba(0,0,0,.06)",zeroline=True,zerolinecolor="rgba(0,0,0,.12)",tickformat=",",tickprefix="¥"),
        xaxis=dict(showgrid=False))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})
    section_end()


# ── PROJECT PAGE COMMON: summary table ──
def render_summary_table(data, with_budget=True, with_yoy=True):
    section("月別サマリー")
    rows = []
    for i, m in enumerate(MONTHS):
        a,c,pr = data["sales_inc"][i], data["cost_inc"][i], data["profit"][i]
        b,ly = data["budget"][i], data["last_year"][i]
        cost_rate = (c/a*100) if a else 0
        margin    = (pr/a*100) if a else 0
        ach       = (a/b*100) if b else 0
        yoy       = (a/ly*100) if ly else 0
        row = {"月":m}
        if with_budget:
            row["売上予算"] = fmt(b) if b else "—"
        row["売上実績"] = fmt(a)
        if with_budget:
            row["予算達成率"] = f"{ach:.1f}%" if b else "—"
        row["原価"]     = fmt(c)
        row["原価率"]   = f"{cost_rate:.1f}%" if a else "—"
        row["営業利益"] = fmt(pr)
        row["利益率"]   = f"{margin:.1f}%" if a else "—"
        if with_yoy:
            row["前年実績"] = fmt(ly) if ly else "—"
            row["前年比"]   = f"{yoy:.1f}%" if ly else "—"
        rows.append(row)
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
    section_end()


# ── AKIBA BROADWAY ──
def render_akibro(data, cats):
    render_header("AKIBABROADWAY")
    mi = page_month_selector(key_suffix="akibro")
    kpi_row_with_budget(data, mi, has_yoy=True)
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    section("月別 予算達成率")
    st.plotly_chart(chart_achievement_monthly(data, mi),
                    use_container_width=True, config={"displayModeBar":False})
    section_end()

    left, right = st.columns([3,2], gap="medium")
    with left:
        section("月別 予算 vs 実績 vs 前年")
        st.plotly_chart(chart_bar_line(data, mi),
                        use_container_width=True, config={"displayModeBar":False})
        section_end()
    with right:
        section("売上内訳" + (f" {MONTHS[mi]}" if mi is not None else " 通期"))
        st.plotly_chart(chart_donut(cats, mi),
                        use_container_width=True, config={"displayModeBar":False})
        section_end()

    section("売上 vs 原価 vs 営業利益")
    st.plotly_chart(chart_sales_cost(data, mi),
                    use_container_width=True, config={"displayModeBar":False})
    section_end()

    render_summary_table(data, with_budget=True, with_yoy=True)


# ── VIRTUAL ──
def render_virtual(data, cats, prepaid, unused_end):
    render_header("VIRTUAL")
    mi = page_month_selector(key_suffix="virtual")

    if mi is not None:
        pp_total = sum(v[mi] for v in prepaid.values()) if prepaid else 0
        ue_total = sum(v[mi] for v in unused_end.values()) if unused_end else 0
    else:
        pp_total = sum(sum(v) for v in prepaid.values()) if prepaid else 0
        # Latest non-zero unused balance
        ue_total = 0
        if unused_end:
            for v in unused_end.values():
                for x in reversed(v):
                    if x != 0:
                        ue_total += x; break
    kpi_row_no_budget(data, mi,
        extra_label="前受金（コイン/チケット購入額）" if mi is not None else "前受金（通期合計）",
        extra_value=fmt(pp_total),
        extra_sub=f"未使用残高 {fmt(ue_total)}")
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    left, right = st.columns([3,2], gap="medium")
    with left:
        section("プラットフォーム別 月次売上推移")
        st.plotly_chart(chart_stacked_bar(cats, mi),
                        use_container_width=True, config={"displayModeBar":False})
        section_end()
    with right:
        section("売上構成比" + (f" {MONTHS[mi]}" if mi is not None else " 通期"))
        st.plotly_chart(chart_donut(cats, mi),
                        use_container_width=True, config={"displayModeBar":False})
        section_end()

    section("売上 vs 原価 vs 営業利益")
    st.plotly_chart(chart_sales_cost(data, mi),
                    use_container_width=True, config={"displayModeBar":False})
    section_end()

    section("営業利益 / 利益率推移")
    st.plotly_chart(chart_profit_with_margin(data, mi),
                    use_container_width=True, config={"displayModeBar":False})
    section_end()

    render_summary_table(data, with_budget=False, with_yoy=True)


# ── NICOLIVE ──
def render_nicolive(data, cats):
    render_header("NICOLIVE")
    mi = page_month_selector(key_suffix="nicolive")

    if mi is not None:
        membership = cats.get("会員費",[0]*12)[mi]
        s_m = data["sales_inc"][mi]
    else:
        membership = sum(cats.get("会員費",[0]*12))
        s_m = sum(data["sales_inc"])
    stock_ratio = (membership/s_m*100) if s_m else 0

    kpi_row_no_budget(data, mi,
        extra_label="会員費（ストック）",
        extra_value=fmt(membership),
        extra_sub=f"ストック比率 {stock_ratio:.1f}%")
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    left, right = st.columns([3,2], gap="medium")
    with left:
        section("カテゴリ別 月次売上推移")
        st.plotly_chart(chart_stacked_bar(cats, mi),
                        use_container_width=True, config={"displayModeBar":False})
        section_end()
    with right:
        section("売上構成比" + (f" {MONTHS[mi]}" if mi is not None else " 通期"))
        st.plotly_chart(chart_donut(cats, mi),
                        use_container_width=True, config={"displayModeBar":False})
        section_end()

    section("売上 vs 原価 vs 営業利益")
    st.plotly_chart(chart_sales_cost(data, mi),
                    use_container_width=True, config={"displayModeBar":False})
    section_end()

    section("営業利益 / 利益率推移")
    st.plotly_chart(chart_profit_with_margin(data, mi),
                    use_container_width=True, config={"displayModeBar":False})
    section_end()

    render_summary_table(data, with_budget=False, with_yoy=False)


# ── AT LIVE ──
def render_atlive(data, cats):
    render_header("AT_LIVE")
    mi = page_month_selector(key_suffix="atlive")
    kpi_row_with_budget(data, mi, has_yoy=True)
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    section("月別 予算達成率（開催月のみハイライト）")
    st.plotly_chart(chart_achievement_monthly(data, mi),
                    use_container_width=True, config={"displayModeBar":False})
    section_end()

    left, right = st.columns([3,2], gap="medium")
    with left:
        section("月別 予算 vs 実績 vs 前年")
        st.plotly_chart(chart_bar_line(data, mi),
                        use_container_width=True, config={"displayModeBar":False})
        section_end()
    with right:
        section("売上構成比" + (f" {MONTHS[mi]}" if mi is not None else " 通期"))
        st.plotly_chart(chart_donut(cats, mi),
                        use_container_width=True, config={"displayModeBar":False})
        section_end()

    section("売上 vs 原価 vs 営業利益")
    st.plotly_chart(chart_sales_cost(data, mi),
                    use_container_width=True, config={"displayModeBar":False})
    section_end()

    render_summary_table(data, with_budget=True, with_yoy=True)


# ── WEDDING ──
def render_wedding(data, cats):
    render_header("WEDDING")
    mi = page_month_selector(key_suffix="wedding")

    cum = sum(data["sales_inc"])
    kpi_row_no_budget(data, mi,
        extra_label="累計売上（通期）", extra_value=fmt(cum),
        extra_sub=f"累計利益 {fmt(sum(data['profit']))}")
    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    left, right = st.columns([3,2], gap="medium")
    with left:
        section("月別 売上推移")
        st.plotly_chart(chart_stacked_bar(cats, mi),
                        use_container_width=True, config={"displayModeBar":False})
        section_end()
    with right:
        section("売上構成比" + (f" {MONTHS[mi]}" if mi is not None else " 通期"))
        st.plotly_chart(chart_donut(cats, mi),
                        use_container_width=True, config={"displayModeBar":False})
        section_end()

    section("売上 vs 原価 vs 営業利益")
    st.plotly_chart(chart_sales_cost(data, mi),
                    use_container_width=True, config={"displayModeBar":False})
    section_end()

    section("累計売上推移")
    st.plotly_chart(chart_cumulative(data),
                    use_container_width=True, config={"displayModeBar":False})
    section_end()

    render_summary_table(data, with_budget=False, with_yoy=True)


# ── PHOTOSHOT (placeholder) ──
def render_photoshot(data):
    render_header("PHOTOSHOT")
    mi = page_month_selector(key_suffix="photoshot")
    st.info("撮影会の売上構成データが入力され次第、ダッシュボードを更新します。")
    kpi_row_no_budget(data, mi)
    render_summary_table(data, with_budget=False, with_yoy=True)


# ═══════════════════════════════════════════════════
# 7. MAIN
# ═══════════════════════════════════════════════════
def main():
    st.set_page_config(page_title="予実管理 2026", layout="wide", initial_sidebar_state="expanded")
    inject_css()

    # ── Sidebar ──
    with st.sidebar:
        st.markdown(f'<div style="display:flex;align-items:center;gap:.6rem;margin-bottom:1rem;">'
                    f'<div style="width:3px;height:24px;background:{PINK};border-radius:2px;"></div>'
                    f'<span style="font-size:.9rem;font-weight:600;color:{CHARCOAL};">予実管理 2026</span></div>',
                    unsafe_allow_html=True)

        st.markdown('<p style="font-size:.65rem;font-weight:500;letter-spacing:.1em;text-transform:uppercase;color:#999;margin-bottom:.15rem;">プロジェクト選択</p>',
                    unsafe_allow_html=True)
        keys = list(PROJECTS.keys())
        labels = {k: v["label"] for k,v in PROJECTS.items()}
        sel = st.selectbox("プロジェクト", keys, format_func=lambda k: labels[k],
                          label_visibility="collapsed")
        st.markdown("---")

        auto_sec = st.selectbox("自動更新", [0,10,30,60],
                                format_func=lambda x: "手動" if x==0 else f"{x}秒")
        st.markdown("---")

        # ── 高度な設定（普段は折りたたみ） ──
        with st.expander("高度な設定（URL / gid）"):
            st.markdown('<p style="font-size:.65rem;font-weight:500;letter-spacing:.1em;text-transform:uppercase;color:#999;margin-bottom:.15rem;">スプレッドシート URL</p>',
                        unsafe_allow_html=True)
            sheet_url_input = st.text_input(
                "URL",
                value=st.session_state.get("sheet_url_override", ""),
                placeholder="既定のURLを使う場合は空欄のままでOK",
                label_visibility="collapsed",
                help=f"既定: {DEFAULT_SHEET_URL[:60]}...",
            )
            if sheet_url_input:
                st.session_state["sheet_url_override"] = sheet_url_input
            elif "sheet_url_override" in st.session_state and not sheet_url_input:
                st.session_state.pop("sheet_url_override", None)

            st.markdown('<p style="font-size:.65rem;color:#aaa;line-height:1.4;margin-top:.5rem;">各シート URL 末尾の「#gid=数字」の数字</p>',
                        unsafe_allow_html=True)
            for k,p in PROJECTS.items():
                if k == "MAIN": continue
                g = st.text_input(p["label"],
                                  value=st.session_state.get(f"gid_{k}", p["gid"]),
                                  placeholder="例: 0", key=f"gid_input_{k}")
                st.session_state[f"gid_{k}"] = g

        # 実際に使う URL は、入力値があればそれ、なければデフォルト
        sheet_url = st.session_state.get("sheet_url_override", "") or DEFAULT_SHEET_URL

        # ── ロゴファイル状態の確認 ──
        with st.expander("ロゴファイル状態"):
            base_dir = Path(__file__).parent
            for k, p in PROJECTS.items():
                if not p.get("logo"):
                    continue
                logo_path = base_dir / p["logo"]
                exists = logo_path.exists()
                icon = "✓" if exists else "✗"
                color = "#10b981" if exists else "#ef4444"
                st.markdown(
                    f'<div style="font-size:.72rem;color:{color};margin:.2rem 0;">'
                    f'{icon} {p["label"]}: <code style="font-size:.68rem;">{p["logo"]}</code></div>',
                    unsafe_allow_html=True,
                )
            st.markdown(
                f'<p style="font-size:.65rem;color:#aaa;margin-top:.5rem;line-height:1.4;">'
                f'実行ディレクトリ: <code style="font-size:.6rem;">{base_dir}</code></p>',
                unsafe_allow_html=True,
            )

    if auto_sec > 0:
        st.markdown(f'<meta http-equiv="refresh" content="{auto_sec}">',
                    unsafe_allow_html=True)

    # ── Top toolbar (reload) ──
    col_r, col_s, _ = st.columns([1,3,2])
    with col_r:
        if st.button("データを再読み込み"): st.rerun()

    ssid = extract_ssid(sheet_url.strip()) if sheet_url.strip() else ""

    # ── Route ──
    if sel == "MAIN":
        with col_s:
            if ssid: st.caption("Google Sheets から取得")
        render_main_page(ssid)
        return

    # Individual project
    gid = st.session_state.get(f"gid_{sel}", PROJECTS[sel]["gid"])
    raw = None; err = None
    if ssid and gid:
        try: raw = fetch_sheet(ssid, gid)
        except Exception as e: err = str(e)
    elif ssid and not gid:
        err = f"「{PROJECTS[sel]['label']}」の gid が未設定です。"

    if raw is None:
        try:
            raw = load_local()
            if err: st.warning(f"スプレッドシート読み込み失敗。ローカル CSV 表示中。\n\n{err}")
        except FileNotFoundError:
            if err: st.error(err)
            else: st.error("データがありません。URL と gid を設定してください。")
            st.stop()

    with col_s:
        if ssid and gid and err is None:
            st.caption(f"Google Sheets (gid={gid}) から取得済み")

    data = parse_common(raw)

    if sel == "AKIBABROADWAY":
        render_akibro(data, parse_akibro_cats(raw))
    elif sel == "VIRTUAL":
        cats, prepaid, unused_end = parse_virtual_cats(raw)
        render_virtual(data, cats, prepaid, unused_end)
    elif sel == "NICOLIVE":
        render_nicolive(data, parse_nicolive_cats(raw))
    elif sel == "AT_LIVE":
        render_atlive(data, parse_atlive_cats(raw))
    elif sel == "WEDDING":
        render_wedding(data, parse_wedding_cats(raw))
    elif sel == "PHOTOSHOT":
        render_photoshot(data)

    st.markdown(f'<div style="text-align:center;padding:1.5rem 0 .5rem;color:#bbb;font-size:.65rem;letter-spacing:.08em;">2026 予実管理ダッシュボード — Prototype</div>',
                unsafe_allow_html=True)


if __name__ == "__main__":
    main()
