"""
2026年度 EC部 予実管理ダッシュボード — マルチプロジェクト対応
"""
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io, re, requests, datetime
 
# ═══════════════════════════════════════════════════
# 0. CONFIG
# ═══════════════════════════════════════════════════
PROJECTS = {
    "MAIN":          {"label": "メインダッシュボード", "subtitle": "全プロジェクト概況",      "gid": ""},
    "AKIBABROADWAY": {"label": "AKIBA BROADWAY",      "subtitle": "お屋敷公演",              "gid": "0"},
    "VIRTUAL":       {"label": "バーチャルあっとほぉーむカフェ", "subtitle": "バーチャル配信", "gid": "1173568321"},
    "NICOLIVE":      {"label": "ニコ生",               "subtitle": "ニコニコ生放送",          "gid": "1834964074"},
    "AT_LIVE":       {"label": "あっとライブ",          "subtitle": "ライブイベント",          "gid": "1558766351"},
    "WEDDING":       {"label": "ウェディング",          "subtitle": "ウェディング事業",        "gid": "1014571293"},
    "PHOTOSHOT":     {"label": "撮影会",               "subtitle": "撮影会イベント",          "gid": "970613346"},
}
 
PINK      = "#f96cb4"
PINK_L    = "#fdb8d8"
WHITE     = "#FFFFFF"
LGRAY     = "#F8F9FA"
MGRAY     = "#DEE2E6"
CHARCOAL  = "#333333"
DARK      = "#1a1a2e"
FONT      = "'Noto Sans JP','Helvetica Neue',Arial,sans-serif"
MONTHS    = ["4月","5月","6月","7月","8月","9月","10月","11月","12月","1月","2月","3月"]
 
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
    .dash-header{{background:linear-gradient(135deg,{DARK},#2d2d4e);padding:1.2rem 2rem 1rem;border-radius:0 0 1rem 1rem;margin:-1rem -1rem 1.2rem -1rem;display:flex;align-items:center;gap:1rem;}}
    .header-accent{{width:3px;height:32px;background:{PINK};border-radius:2px;}}
    .dash-header h1{{color:{WHITE};font-size:1.2rem;font-weight:600;margin:0;}}
    .dash-header p{{color:rgba(255,255,255,.5);font-size:.72rem;font-weight:300;margin:0;}}
    .kpi-card{{background:{WHITE};border-radius:.65rem;padding:1.1rem 1.2rem;box-shadow:0 1px 4px rgba(0,0,0,.06);border:1px solid rgba(0,0,0,.04);transition:box-shadow .2s;}}
    .kpi-card:hover{{box-shadow:0 4px 16px rgba(0,0,0,.08);}}
    .kpi-label{{font-size:.68rem;font-weight:500;letter-spacing:.1em;text-transform:uppercase;color:#999;margin-bottom:.4rem;}}
    .kpi-value{{font-size:1.45rem;font-weight:700;color:{CHARCOAL};line-height:1.15;}}
    .kpi-sub{{font-size:.7rem;color:#aaa;margin-top:.2rem;}}
    .kpi-accent{{color:{PINK};}}
    .section-card{{background:{WHITE};border-radius:.65rem;padding:1.4rem 1.6rem;box-shadow:0 1px 4px rgba(0,0,0,.06);border:1px solid rgba(0,0,0,.04);margin-bottom:1rem;}}
    .section-title{{font-size:.85rem;font-weight:600;color:{CHARCOAL};margin-bottom:.9rem;padding-bottom:.4rem;border-bottom:2px solid {PINK};display:inline-block;}}
    .prog-track{{background:{LGRAY};border-radius:6px;height:7px;overflow:hidden;margin-top:.4rem;}}
    .prog-fill{{height:100%;border-radius:6px;background:linear-gradient(90deg,{PINK},{PINK_L});transition:width .6s;}}
    .proj-summary-card{{background:{WHITE};border-radius:.65rem;padding:1.2rem 1.4rem;box-shadow:0 1px 4px rgba(0,0,0,.06);border:1px solid rgba(0,0,0,.04);margin-bottom:.8rem;}}
    .proj-summary-card h3{{font-size:.9rem;font-weight:600;color:{CHARCOAL};margin:0 0 .6rem;padding-bottom:.3rem;border-bottom:2px solid {PINK};display:inline-block;}}
    .proj-row{{display:flex;gap:1.5rem;flex-wrap:wrap;margin-top:.4rem;}}
    .proj-metric{{min-width:100px;}}
    .proj-metric-label{{font-size:.65rem;color:#999;letter-spacing:.05em;}}
    .proj-metric-value{{font-size:1.1rem;font-weight:700;color:{CHARCOAL};}}
    .profit-positive{{color:#2ecc71;}}
    .profit-negative{{color:#e74c3c;}}
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
 
# ── Common parser ──
def parse_common(raw):
    d = {}
    d["budget"]          = rv(raw, "売上予算")
    d["sales_inc"]       = rv(raw, "売上合計(税込)")
    d["sales_exc"]       = rv(raw, "売上合計(税抜)")
    d["last_year"]       = rv(raw, "昨年売上")
    d["cost_budget"]     = rv(raw, "変動費予算")
    d["cost_inc"]        = rv(raw, "原価合計(税込)")
    d["cost_exc"]        = rv(raw, "原価合計(税抜)")
    d["sg_a"]            = rv(raw, "販管費")
    d["profit"]          = rv(raw, "営業利益")
    return d
 
# ── Project-specific category parsers ──
def parse_akibro_cats(raw):
    cats = ["物販（会場内チェキ）","物販（会場内グッズ）","ガイドツアー",
            "グッズ（MD）","グッズ通販（MD）","チケット（前売り）",
            "チケット（当日）","チケット（海外）","外部"]
    return {c: rv(raw, c) for c in cats}
 
def parse_virtual_cats(raw):
    cats = {}
    coin_idx = 0
    pf_names = ["SBペイメント","Apple","Google"]
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
    # 前受金（キャッシュイン参考）
    prepaid = {}
    prepaid_labels = {
        "コイン購入額（前受金）": "SBペイメント",
        "コイン購入額 − 手数料（前受金）": None,  # Apple, Google の順
        "チケット購入額（前受金）": "Stripe",
    }
    apple_google_count = 0
    for i in range(raw.shape[0]):
        cell = str(raw.iloc[i,0]).strip()
        vals = [yen(raw.iloc[i,j]) for j in range(1, min(13, raw.shape[1]))]
        if cell == "コイン購入額（前受金）":
            prepaid["SBペイメント"] = vals
        elif "コイン購入額" in cell and "手数料" in cell:
            name = ["Apple","Google"][apple_google_count] if apple_google_count < 2 else "Other"
            prepaid[name] = vals
            apple_google_count += 1
        elif cell == "チケット購入額（前受金）":
            prepaid["Stripe"] = vals
    # 未使用残高
    unused_end = {}
    ue_count = {"coin":0, "ticket":0}
    ue_names_coin = ["SBペイメント","Apple","Google"]
    for i in range(raw.shape[0]):
        cell = str(raw.iloc[i,0]).strip()
        vals = [yen(raw.iloc[i,j]) for j in range(1, min(13, raw.shape[1]))]
        if "当月末" in cell and "コイン" in cell and ue_count["coin"] < 3:
            unused_end[ue_names_coin[ue_count["coin"]]] = vals
            ue_count["coin"] += 1
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
 
def _hl_colors(base, n, hi=None):
    """Return list of colors; dim non-highlighted bars."""
    cols = [base]*n
    if hi is not None:
        cols = [base if i==hi else f"rgba({int(base[1:3],16)},{int(base[3:5],16)},{int(base[5:7],16)},.3)" for i in range(n)]
    return cols
 
def chart_bar_line(data, hi=None):
    """Budget(line) vs Actual(bar) vs LastYear(bar)."""
    n = len(MONTHS)
    fig = go.Figure()
    fig.add_trace(go.Bar(x=MONTHS, y=data["sales_inc"], name="実績（税込）",
                         marker_color=_hl_colors(PINK, n, hi), marker_line_width=0, opacity=.85))
    fig.add_trace(go.Bar(x=MONTHS, y=data["last_year"], name="前年実績",
                         marker_color=_hl_colors(MGRAY, n, hi), marker_line_width=0, opacity=.55))
    fig.add_trace(go.Scatter(x=MONTHS, y=data["budget"], name="売上予算", mode="lines+markers",
                             line=dict(color=CHARCOAL, width=2, dash="dot"),
                             marker=dict(size=5, color=CHARCOAL)))
    fig.update_layout(**PL, barmode="group", height=370,
        legend=dict(orientation="h",yanchor="bottom",y=1.04,xanchor="center",x=.5,font=dict(size=11)),
        yaxis=dict(gridcolor="rgba(0,0,0,.06)",zeroline=False,tickformat=",",tickprefix="¥"),
        xaxis=dict(showgrid=False))
    return fig
 
def chart_donut(categories, title_center=""):
    labels = list(categories.keys())
    values = [sum(v) for v in categories.values()]
    n = len(labels)
    colors = [f"rgb({int(249*(1-i/max(n-1,1))+222*(i/max(n-1,1)))},"
              f"{int(108*(1-i/max(n-1,1))+226*(i/max(n-1,1)))},"
              f"{int(180*(1-i/max(n-1,1))+230*(i/max(n-1,1)))})" for i in range(n)]
    fig = go.Figure(go.Pie(labels=labels,values=values,hole=.58,
        marker=dict(colors=colors,line=dict(color=WHITE,width=2)),
        textinfo="percent",textfont=dict(size=11,color=CHARCOAL),
        hovertemplate="%{label}<br>¥%{value:,}<br>%{percent}<extra></extra>",sort=False))
    fig.update_layout(**{**PL,"margin":dict(l=20,r=160,t=30,b=20)},showlegend=True,
        legend=dict(orientation="v",yanchor="middle",y=.5,xanchor="left",x=1.02,font=dict(size=10)),height=360)
    return fig
 
def chart_donut_month(categories, mi):
    labels = list(categories.keys())
    values = [v[mi] for v in categories.values()]
    n = len(labels)
    colors = [f"rgb({int(249*(1-i/max(n-1,1))+222*(i/max(n-1,1)))},"
              f"{int(108*(1-i/max(n-1,1))+226*(i/max(n-1,1)))},"
              f"{int(180*(1-i/max(n-1,1))+230*(i/max(n-1,1)))})" for i in range(n)]
    fig = go.Figure(go.Pie(labels=labels,values=values,hole=.58,
        marker=dict(colors=colors,line=dict(color=WHITE,width=2)),
        textinfo="percent",textfont=dict(size=11,color=CHARCOAL),
        hovertemplate="%{label}<br>¥%{value:,}<br>%{percent}<extra></extra>",sort=False))
    fig.update_layout(**{**PL,"margin":dict(l=20,r=160,t=30,b=20)},showlegend=True,
        legend=dict(orientation="v",yanchor="middle",y=.5,xanchor="left",x=1.02,font=dict(size=10)),height=360)
    return fig
 
def chart_profit(data, hi=None):
    p = data["profit"]
    n = len(MONTHS)
    base_colors = [PINK if v>=0 else "#e74c3c" for v in p]
    if hi is not None:
        colors = []
        for i,c in enumerate(base_colors):
            if i == hi:
                colors.append(c)
            else:
                r,g,b = int(c[1:3],16),int(c[3:5],16),int(c[5:7],16)
                colors.append(f"rgba({r},{g},{b},.3)")
    else:
        colors = base_colors
    fig = go.Figure(go.Bar(x=MONTHS,y=p,marker_color=colors,marker_line_width=0,opacity=.85))
    fig.update_layout(**PL,height=300,
        yaxis=dict(gridcolor="rgba(0,0,0,.06)",zeroline=True,zerolinecolor="rgba(0,0,0,.12)",tickformat=",",tickprefix="¥"),
        xaxis=dict(showgrid=False))
    return fig
 
def chart_stacked_bar(categories, hi=None):
    """Stacked bar for platform/category breakdown by month."""
    fig = go.Figure()
    n_cats = len(categories)
    for idx, (name, vals) in enumerate(categories.items()):
        t = idx / max(n_cats-1,1)
        r = int(249*(1-t)+222*t); g = int(108*(1-t)+226*t); b = int(180*(1-t)+230*t)
        color = f"rgb({r},{g},{b})"
        fig.add_trace(go.Bar(x=MONTHS, y=vals, name=name, marker_color=color, marker_line_width=0))
    fig.update_layout(**PL, barmode="stack", height=370,
        legend=dict(orientation="h",yanchor="bottom",y=1.04,xanchor="center",x=.5,font=dict(size=10)),
        yaxis=dict(gridcolor="rgba(0,0,0,.06)",zeroline=False,tickformat=",",tickprefix="¥"),
        xaxis=dict(showgrid=False))
    return fig
 
def chart_profit_with_margin(data):
    """Profit bars + profit margin line overlay."""
    p = data["profit"]
    s = data["sales_inc"]
    margins = [(p[i]/s[i]*100 if s[i] else 0) for i in range(12)]
    colors = [PINK if v>=0 else "#e74c3c" for v in p]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=MONTHS,y=p,name="営業利益",marker_color=colors,marker_line_width=0,opacity=.85))
    fig.add_trace(go.Scatter(x=MONTHS,y=margins,name="利益率(%)",
        mode="lines+markers",line=dict(color=CHARCOAL,width=2),
        marker=dict(size=5,color=CHARCOAL),yaxis="y2"))
    fig.update_layout(**PL,height=320,
        yaxis=dict(gridcolor="rgba(0,0,0,.06)",zeroline=True,zerolinecolor="rgba(0,0,0,.12)",tickformat=",",tickprefix="¥"),
        yaxis2=dict(overlaying="y",side="right",showgrid=False,ticksuffix="%",zeroline=False),
        xaxis=dict(showgrid=False),
        legend=dict(orientation="h",yanchor="bottom",y=1.04,xanchor="center",x=.5,font=dict(size=11)))
    return fig
 
def chart_event_achievement(data):
    """Achievement rate bars highlighting event months."""
    budget = data["budget"]
    actual = data["sales_inc"]
    rates = [(actual[i]/budget[i]*100 if budget[i] else 0) for i in range(12)]
    colors = [PINK if budget[i]>0 else MGRAY for i in range(12)]
    fig = go.Figure(go.Bar(x=MONTHS,y=rates,marker_color=colors,marker_line_width=0,opacity=.85))
    fig.add_hline(y=100,line_dash="dot",line_color=CHARCOAL,line_width=1.5,annotation_text="目標100%",annotation_position="top right")
    fig.update_layout(**PL,height=320,
        yaxis=dict(gridcolor="rgba(0,0,0,.06)",zeroline=False,ticksuffix="%"),
        xaxis=dict(showgrid=False))
    return fig
 
def chart_cumulative_line(data):
    """Cumulative sales line."""
    s = data["sales_inc"]
    cum = []
    total = 0
    for v in s:
        total += v
        cum.append(total)
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
    p = PROJECTS[key]
    st.markdown(f"""<div class="dash-header">
        <div class="header-accent"></div>
        <div><h1>{p["label"]}</h1><p>{p["subtitle"]} — 2026年度 予実管理</p></div>
    </div>""", unsafe_allow_html=True)
 
def kpi_card(label, value_html, sub=""):
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value_html}</div>{sub_html}
    </div>""", unsafe_allow_html=True)
 
def kpi_card_progress(label, value_html, pct):
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value kpi-accent">{value_html}</div>
        <div class="prog-track"><div class="prog-fill" style="width:{min(pct,100):.1f}%"></div></div>
    </div>""", unsafe_allow_html=True)
 
def month_selector():
    opts = ["全期間"] + MONTHS
    sel = st.selectbox("表示月", opts, index=0, label_visibility="collapsed")
    if sel == "全期間":
        return None
    return MONTHS.index(sel)
 
def section(title):
    st.markdown(f'<div class="section-card"><div class="section-title">{title}</div>', unsafe_allow_html=True)
 
def section_end():
    st.markdown('</div>', unsafe_allow_html=True)
 
# ═══════════════════════════════════════════════════
# 5. PAGE RENDERERS
# ═══════════════════════════════════════════════════
 
# ── MAIN PAGE ──
def render_main_page(ssid, sheet_url):
    render_header("MAIN")
    mi = fiscal_month_index()
    month_label = MONTHS[mi]
 
    st.markdown(f'<p style="font-size:.82rem;color:#999;margin-bottom:1rem;">表示月: <b style="color:{CHARCOAL};">{month_label}</b></p>', unsafe_allow_html=True)
 
    all_data = {}
    for key, proj in PROJECTS.items():
        if key == "MAIN": continue
        gid = st.session_state.get(f"gid_{key}", proj["gid"])
        if not gid: continue
        try:
            if ssid:
                raw = fetch_sheet(ssid, gid)
            else:
                continue
            all_data[key] = parse_common(raw)
        except Exception:
            pass
 
    if not all_data:
        st.info("スプレッドシートの URL と各シートの gid を設定してください。")
        return
 
    # Summary cards for each project
    total_sales = 0
    total_profit = 0
    for key, d in all_data.items():
        p = PROJECTS[key]
        sales_m  = d["sales_inc"][mi]
        profit_m = d["profit"][mi]
        budget_m = d["budget"][mi]
        cost_m   = d["cost_inc"][mi]
        total_sales  += sales_m
        total_profit += profit_m
        ach = (sales_m/budget_m*100) if budget_m else 0
        pr_class = "profit-positive" if profit_m >= 0 else "profit-negative"
        margin = (profit_m/sales_m*100) if sales_m else 0
 
        st.markdown(f"""<div class="proj-summary-card">
            <h3>{p["label"]}</h3>
            <div class="proj-row">
                <div class="proj-metric">
                    <div class="proj-metric-label">売上（税込）</div>
                    <div class="proj-metric-value">{fmt(sales_m)}</div>
                </div>
                <div class="proj-metric">
                    <div class="proj-metric-label">原価（税込）</div>
                    <div class="proj-metric-value">{fmt(cost_m)}</div>
                </div>
                <div class="proj-metric">
                    <div class="proj-metric-label">営業利益</div>
                    <div class="proj-metric-value {pr_class}">{fmt(profit_m)}</div>
                </div>
                <div class="proj-metric">
                    <div class="proj-metric-label">利益率</div>
                    <div class="proj-metric-value {pr_class}">{margin:.1f}%</div>
                </div>
                {"" if not budget_m else f'''<div class="proj-metric">
                    <div class="proj-metric-label">予算達成率</div>
                    <div class="proj-metric-value">{ach:.1f}%</div>
                </div>'''}
            </div>
        </div>""", unsafe_allow_html=True)
 
    # Total bar
    pr_class = "profit-positive" if total_profit >= 0 else "profit-negative"
    st.markdown(f"""<div class="proj-summary-card" style="background:linear-gradient(135deg,{DARK},#2d2d4e);margin-top:.5rem;">
        <h3 style="color:{WHITE};border-color:{PINK};">全プロジェクト合計（{month_label}）</h3>
        <div class="proj-row">
            <div class="proj-metric"><div class="proj-metric-label" style="color:rgba(255,255,255,.5);">売上合計</div>
                <div class="proj-metric-value" style="color:{WHITE};">{fmt(total_sales)}</div></div>
            <div class="proj-metric"><div class="proj-metric-label" style="color:rgba(255,255,255,.5);">営業利益合計</div>
                <div class="proj-metric-value {pr_class}">{fmt(total_profit)}</div></div>
        </div>
    </div>""", unsafe_allow_html=True)
 
    # Comparison bar chart
    section("プロジェクト別 当月売上比較")
    names = [PROJECTS[k]["label"] for k in all_data]
    sales_vals = [all_data[k]["sales_inc"][mi] for k in all_data]
    profit_vals = [all_data[k]["profit"][mi] for k in all_data]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=names,y=sales_vals,name="売上（税込）",marker_color=PINK,marker_line_width=0,opacity=.85))
    fig.add_trace(go.Bar(x=names,y=profit_vals,name="営業利益",
        marker_color=[PINK_L if v>=0 else "#e74c3c" for v in profit_vals],marker_line_width=0,opacity=.7))
    fig.update_layout(**PL,barmode="group",height=350,
        legend=dict(orientation="h",yanchor="bottom",y=1.04,xanchor="center",x=.5,font=dict(size=11)),
        yaxis=dict(gridcolor="rgba(0,0,0,.06)",zeroline=True,zerolinecolor="rgba(0,0,0,.12)",tickformat=",",tickprefix="¥"),
        xaxis=dict(showgrid=False))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})
    section_end()
 
 
# ── AKIBA BROADWAY ──
def render_akibro(data, cats, mi):
    render_header("AKIBABROADWAY")
    # KPI
    if mi is not None:
        s,b,ly,pr = data["sales_inc"][mi],data["budget"][mi],data["last_year"][mi],data["profit"][mi]
        label_suffix = f"（{MONTHS[mi]}）"
    else:
        s,b,ly,pr = sum(data["sales_inc"]),sum(data["budget"]),sum(data["last_year"]),sum(data["profit"])
        label_suffix = "（通期）"
    ach = (s/b*100) if b else 0
    yoy = (s/ly*100) if ly else 0
    cols = st.columns(4, gap="medium")
    with cols[0]: kpi_card(f"売上実績{label_suffix}", fmt(s), f"予算 {fmt(b)}")
    with cols[1]: kpi_card_progress(f"予算達成率{label_suffix}", f"{ach:.1f}%", ach)
    with cols[2]: kpi_card(f"前年比{label_suffix}", f"{yoy:.1f}%", f"前年 {fmt(ly)}")
    with cols[3]:
        pr_rate = (pr/s*100) if s else 0
        kpi_card(f"営業利益{label_suffix}", fmt(pr), f"利益率 {pr_rate:.1f}%")
    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
 
    left, right = st.columns([3,2], gap="medium")
    with left:
        section("月別 予算 vs 実績 vs 前年")
        st.plotly_chart(chart_bar_line(data, mi), use_container_width=True, config={"displayModeBar":False})
        section_end()
    with right:
        section("売上内訳" + (f" {MONTHS[mi]}" if mi is not None else " 通期"))
        if mi is not None:
            st.plotly_chart(chart_donut_month(cats, mi), use_container_width=True, config={"displayModeBar":False})
        else:
            st.plotly_chart(chart_donut(cats), use_container_width=True, config={"displayModeBar":False})
        section_end()
    section("月別 営業利益")
    st.plotly_chart(chart_profit(data, mi), use_container_width=True, config={"displayModeBar":False})
    section_end()
 
 
# ── VIRTUAL ──
def render_virtual(data, cats, prepaid, unused_end, mi):
    render_header("VIRTUAL")
    if mi is not None:
        s,pr = data["sales_inc"][mi],data["profit"][mi]
        pp_total = sum(v[mi] for v in prepaid.values()) if prepaid else 0
        ue_total = sum(v[mi] for v in unused_end.values()) if unused_end else 0
        label_suffix = f"（{MONTHS[mi]}）"
    else:
        s,pr = sum(data["sales_inc"]),sum(data["profit"])
        pp_total = sum(sum(v) for v in prepaid.values()) if prepaid else 0
        ue_total = sum(v[0] for v in unused_end.values()) if unused_end else 0  # latest snapshot
        label_suffix = "（通期）"
    margin = (pr/s*100) if s else 0
    cols = st.columns(4, gap="medium")
    with cols[0]: kpi_card(f"売上実績{label_suffix}", fmt(s))
    with cols[1]:
        pr_cls = "" if pr >= 0 else ' style="color:#e74c3c"'
        kpi_card(f"営業利益{label_suffix}", f'<span{pr_cls}>{fmt(pr)}</span>', f"利益率 {margin:.1f}%")
    with cols[2]: kpi_card(f"前受金{label_suffix}", fmt(pp_total), "コイン/チケット購入額")
    with cols[3]: kpi_card("未使用残高", fmt(ue_total), "将来の売上ポテンシャル")
    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
 
    left, right = st.columns([3,2], gap="medium")
    with left:
        section("プラットフォーム別 月次売上推移")
        st.plotly_chart(chart_stacked_bar(cats, mi), use_container_width=True, config={"displayModeBar":False})
        section_end()
    with right:
        section("売上構成比" + (f" {MONTHS[mi]}" if mi is not None else " 通期"))
        if mi is not None:
            st.plotly_chart(chart_donut_month(cats, mi), use_container_width=True, config={"displayModeBar":False})
        else:
            st.plotly_chart(chart_donut(cats), use_container_width=True, config={"displayModeBar":False})
        section_end()
    section("営業利益 / 利益率推移")
    st.plotly_chart(chart_profit_with_margin(data), use_container_width=True, config={"displayModeBar":False})
    section_end()
 
 
# ── NICOLIVE ──
def render_nicolive(data, cats, mi):
    render_header("NICOLIVE")
    if mi is not None:
        s,pr = data["sales_inc"][mi],data["profit"][mi]
        membership = cats.get("会員費",[0]*12)[mi]
        label_suffix = f"（{MONTHS[mi]}）"
    else:
        s,pr = sum(data["sales_inc"]),sum(data["profit"])
        membership = sum(cats.get("会員費",[0]*12))
        label_suffix = "（通期）"
    margin = (pr/s*100) if s else 0
    stock_ratio = (membership/s*100) if s else 0
    cols = st.columns(4, gap="medium")
    with cols[0]: kpi_card(f"売上実績{label_suffix}", fmt(s))
    with cols[1]: kpi_card(f"営業利益{label_suffix}", fmt(pr), f"利益率 {margin:.1f}%")
    with cols[2]: kpi_card(f"会員費{label_suffix}", fmt(membership), f"ストック比率 {stock_ratio:.1f}%")
    with cols[3]: kpi_card_progress("利益率", f"{margin:.1f}%", max(margin,0))
    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
 
    left, right = st.columns([3,2], gap="medium")
    with left:
        section("カテゴリ別 月次売上推移")
        st.plotly_chart(chart_stacked_bar(cats, mi), use_container_width=True, config={"displayModeBar":False})
        section_end()
    with right:
        section("売上構成比" + (f" {MONTHS[mi]}" if mi is not None else " 通期"))
        if mi is not None:
            st.plotly_chart(chart_donut_month(cats, mi), use_container_width=True, config={"displayModeBar":False})
        else:
            st.plotly_chart(chart_donut(cats), use_container_width=True, config={"displayModeBar":False})
        section_end()
    section("営業利益 / 利益率推移")
    st.plotly_chart(chart_profit_with_margin(data), use_container_width=True, config={"displayModeBar":False})
    section_end()
 
 
# ── AT LIVE ──
def render_atlive(data, cats, mi):
    render_header("AT_LIVE")
    if mi is not None:
        s,b,pr = data["sales_inc"][mi],data["budget"][mi],data["profit"][mi]
        label_suffix = f"（{MONTHS[mi]}）"
    else:
        s,b,pr = sum(data["sales_inc"]),sum(data["budget"]),sum(data["profit"])
        label_suffix = "（通期）"
    ach = (s/b*100) if b else 0
    cols = st.columns(4, gap="medium")
    with cols[0]: kpi_card(f"売上実績{label_suffix}", fmt(s), f"予算 {fmt(b)}" if b else "")
    with cols[1]: kpi_card_progress(f"予算達成率{label_suffix}", f"{ach:.1f}%", ach) if b else kpi_card(f"予算達成率{label_suffix}", "—")
    with cols[2]: kpi_card(f"営業利益{label_suffix}", fmt(pr))
    with cols[3]:
        margin = (pr/s*100) if s else 0
        kpi_card("利益率", f"{margin:.1f}%")
    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
 
    left, right = st.columns([3,2], gap="medium")
    with left:
        section("月別 予算達成率")
        st.plotly_chart(chart_event_achievement(data), use_container_width=True, config={"displayModeBar":False})
        section_end()
    with right:
        section("売上構成比" + (f" {MONTHS[mi]}" if mi is not None else " 通期"))
        if mi is not None:
            st.plotly_chart(chart_donut_month(cats, mi), use_container_width=True, config={"displayModeBar":False})
        else:
            st.plotly_chart(chart_donut(cats), use_container_width=True, config={"displayModeBar":False})
        section_end()
    section("月別 営業利益")
    st.plotly_chart(chart_profit(data, mi), use_container_width=True, config={"displayModeBar":False})
    section_end()
 
 
# ── WEDDING ──
def render_wedding(data, cats, mi):
    render_header("WEDDING")
    if mi is not None:
        s,pr = data["sales_inc"][mi],data["profit"][mi]
        label_suffix = f"（{MONTHS[mi]}）"
    else:
        s,pr = sum(data["sales_inc"]),sum(data["profit"])
        label_suffix = "（通期）"
    margin = (pr/s*100) if s else 0
    cols = st.columns(4, gap="medium")
    with cols[0]: kpi_card(f"売上実績{label_suffix}", fmt(s))
    with cols[1]: kpi_card(f"営業利益{label_suffix}", fmt(pr), f"利益率 {margin:.1f}%")
    with cols[2]:
        # 累計
        cum = sum(data["sales_inc"])
        kpi_card("累計売上", fmt(cum))
    with cols[3]:
        cum_pr = sum(data["profit"])
        kpi_card("累計利益", fmt(cum_pr))
    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
 
    left, right = st.columns([3,2], gap="medium")
    with left:
        section("月別 売上推移")
        st.plotly_chart(chart_stacked_bar(cats, mi), use_container_width=True, config={"displayModeBar":False})
        section_end()
    with right:
        section("売上構成比" + (f" {MONTHS[mi]}" if mi is not None else " 通期"))
        if mi is not None:
            st.plotly_chart(chart_donut_month(cats, mi), use_container_width=True, config={"displayModeBar":False})
        else:
            st.plotly_chart(chart_donut(cats), use_container_width=True, config={"displayModeBar":False})
        section_end()
    section("累計売上推移")
    st.plotly_chart(chart_cumulative_line(data), use_container_width=True, config={"displayModeBar":False})
    section_end()
 
 
# ── PHOTOSHOT (placeholder) ──
def render_photoshot(data, mi):
    render_header("PHOTOSHOT")
    st.info("撮影会の売上構成データが入力され次第、ダッシュボードを更新します。")
    if mi is not None:
        s,pr = data["sales_inc"][mi],data["profit"][mi]
        label_suffix = f"（{MONTHS[mi]}）"
    else:
        s,pr = sum(data["sales_inc"]),sum(data["profit"])
        label_suffix = "（通期）"
    cols = st.columns(4, gap="medium")
    with cols[0]: kpi_card(f"売上実績{label_suffix}", fmt(s))
    with cols[1]: kpi_card(f"営業利益{label_suffix}", fmt(pr))
    with cols[2]: pass
    with cols[3]: pass
 
 
# ═══════════════════════════════════════════════════
# 6. MAIN
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
 
        st.markdown('<p style="font-size:.65rem;font-weight:500;letter-spacing:.1em;text-transform:uppercase;color:#999;margin-bottom:.15rem;">プロジェクト選択</p>', unsafe_allow_html=True)
        keys = list(PROJECTS.keys())
        labels = {k: v["label"] for k,v in PROJECTS.items()}
        sel = st.selectbox("プロジェクト", keys, format_func=lambda k: labels[k], label_visibility="collapsed")
        st.markdown("---")
 
        st.markdown('<p style="font-size:.65rem;font-weight:500;letter-spacing:.1em;text-transform:uppercase;color:#999;margin-bottom:.15rem;">スプレッドシート URL</p>', unsafe_allow_html=True)
        sheet_url = st.text_input("URL", value=st.session_state.get("sheet_url",""),
            placeholder="https://docs.google.com/spreadsheets/d/xxxxx/edit", label_visibility="collapsed")
        if sheet_url: st.session_state["sheet_url"] = sheet_url
        st.markdown('<p style="font-size:.65rem;color:#aaa;line-height:1.5;margin-top:.15rem;">共有→「リンクを知っている全員：閲覧者」</p>', unsafe_allow_html=True)
        st.markdown("---")
 
        if sel != "MAIN":
            st.markdown('<p style="font-size:.65rem;font-weight:500;letter-spacing:.1em;text-transform:uppercase;color:#999;margin-bottom:.15rem;">表示月</p>', unsafe_allow_html=True)
            mi = month_selector()
        else:
            mi = None
 
        st.markdown("---")
        auto_sec = st.selectbox("自動更新", [0,10,30,60], format_func=lambda x: "手動" if x==0 else f"{x}秒")
        st.markdown("---")
        with st.expander("各シートの gid"):
            st.markdown('<p style="font-size:.65rem;color:#aaa;line-height:1.4;">各シートURL末尾の「#gid=数字」の数字</p>', unsafe_allow_html=True)
            for k,p in PROJECTS.items():
                if k == "MAIN": continue
                g = st.text_input(p["label"], value=st.session_state.get(f"gid_{k}", p["gid"]),
                                  placeholder="例: 0", key=f"gid_input_{k}")
                st.session_state[f"gid_{k}"] = g
 
    if auto_sec > 0:
        st.markdown(f'<meta http-equiv="refresh" content="{auto_sec}">', unsafe_allow_html=True)
 
    # ── Load ──
    ssid = extract_ssid(sheet_url.strip()) if sheet_url.strip() else ""
 
    if sel == "MAIN":
        col_r, _ = st.columns([1,5])
        with col_r:
            if st.button("データを再読み込み"): st.rerun()
        render_main_page(ssid, sheet_url)
        return
 
    # Individual project page
    gid = st.session_state.get(f"gid_{sel}", PROJECTS[sel]["gid"])
    raw = None
    err = None
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
 
    data = parse_common(raw)
 
    col_r, col_s, _ = st.columns([1,3,2])
    with col_r:
        if st.button("データを再読み込み"): st.rerun()
    with col_s:
        if ssid and gid and err is None: st.caption(f"Google Sheets (gid={gid}) から取得済み")
 
    # Dispatch
    if sel == "AKIBABROADWAY":
        cats = parse_akibro_cats(raw)
        data["categories"] = cats
        render_akibro(data, cats, mi)
    elif sel == "VIRTUAL":
        cats, prepaid, unused_end = parse_virtual_cats(raw)
        data["categories"] = cats
        render_virtual(data, cats, prepaid, unused_end, mi)
    elif sel == "NICOLIVE":
        cats = parse_nicolive_cats(raw)
        data["categories"] = cats
        render_nicolive(data, cats, mi)
    elif sel == "AT_LIVE":
        cats = parse_atlive_cats(raw)
        data["categories"] = cats
        render_atlive(data, cats, mi)
    elif sel == "WEDDING":
        cats = parse_wedding_cats(raw)
        data["categories"] = cats
        render_wedding(data, cats, mi)
    elif sel == "PHOTOSHOT":
        render_photoshot(data, mi)
 
    # ── Summary Table ──
    st.markdown("<div style='height:.5rem'></div>", unsafe_allow_html=True)
    section("月別サマリー")
    rows = []
    for i, m in enumerate(MONTHS):
        a,b,ly = data["sales_inc"][i],data["budget"][i],data["last_year"][i]
        rows.append({"月":m, "売上予算":fmt(b), "売上実績":fmt(a),
            "予算達成率":f"{(a/b*100) if b else 0:.1f}%",
            "前年実績":fmt(ly), "前年比":f"{(a/ly*100) if ly else 0:.1f}%",
            "営業利益":fmt(data["profit"][i])})
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
    section_end()
 
    # Footer
    st.markdown(f'<div style="text-align:center;padding:1.5rem 0 .5rem;color:#bbb;font-size:.65rem;letter-spacing:.08em;">2026 予実管理ダッシュボード — Prototype</div>', unsafe_allow_html=True)
 
 
if __name__ == "__main__":
    main()
