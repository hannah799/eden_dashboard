import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from collections import defaultdict
import openpyxl

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Eden Cannabis | Dashboard",
    page_icon="🌿",
    layout="wide",
)

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #0a1628; }
  [data-testid="stSidebar"] { background: #0d1f3c; }
  .main .block-container { padding-top: 1.5rem; max-width: 1400px; }

  .eden-header {
    background: linear-gradient(135deg, #0d1f3c 0%, #1a3a5c 100%);
    border-radius: 12px; padding: 24px 32px; margin-bottom: 24px;
    border: 1px solid #1e4d7b;
  }
  .eden-header h1 { color: #e8f4e8; margin: 0; font-size: 2rem; font-weight: 700; }
  .eden-header p  { color: #7ab8a0; margin: 4px 0 0; font-size: 0.95rem; }

  .kpi-card {
    background: linear-gradient(135deg, #0d1f3c 0%, #1a3a5c 100%);
    border: 1px solid #1e4d7b; border-radius: 10px;
    padding: 20px; text-align: center;
  }
  .kpi-label { color: #7ab8a0; font-size: 0.78rem; font-weight: 600;
               text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 6px; }
  .kpi-value { color: #e8f4e8; font-size: 1.9rem; font-weight: 700; line-height: 1; }
  .kpi-sub   { color: #5a8fa0; font-size: 0.75rem; margin-top: 4px; }

  .section-header {
    color: #7ab8a0; font-size: 0.75rem; font-weight: 700;
    text-transform: uppercase; letter-spacing: 0.1em;
    margin: 24px 0 12px; padding-bottom: 6px; border-bottom: 1px solid #1e4d7b;
  }
  .alert-box {
    background: #1a3a1a; border: 1px solid #3a7a3a;
    border-left: 4px solid #5ab85a; border-radius: 8px;
    padding: 12px 16px; margin-bottom: 16px;
  }
  .alert-box .alert-title { color: #5ab85a; font-weight: 700; font-size: 0.85rem; }
  .alert-box .alert-body  { color: #a0c8a0; font-size: 0.82rem; margin-top: 4px; }

  h2, h3 { color: #e8f4e8 !important; }
  p, li  { color: #b0ccd8; }
  hr     { border-color: #1e4d7b; }

  /* Tab styling */
  .stTabs [data-baseweb="tab-list"] { gap: 8px; }
  .stTabs [data-baseweb="tab"] {
    background: #0d1f3c; border: 1px solid #1e4d7b;
    border-radius: 8px 8px 0 0; color: #7ab8a0;
    padding: 8px 20px; font-weight: 600;
  }
  .stTabs [aria-selected="true"] {
    background: #1a3a5c !important; color: #e8f4e8 !important;
    border-bottom-color: #1a3a5c !important;
  }
</style>
""", unsafe_allow_html=True)

# ── Color palette ─────────────────────────────────────────────────────────────
COLORS = {
    "Flower":        "#4a9e6b",
    "Concentrates":  "#2e7bc4",
    "Carts":         "#7b5ea7",
    "Edibles":       "#e07b39",
    "Pre-rolls":     "#c4a244",
    "Tinctures":     "#3aacb8",
    "Topical":       "#c45e8a",
}
ROOM_COLORS = {
    "F1-Trailer": "#4a9e6b",
    "F2-Trailer": "#2e7bc4",
    "F3-Trailer": "#e07b39",
    "F4-Trailer": "#7b5ea7",
}
CHART_BG   = "#0d1f3c"
GRID_COLOR = "#1e3a5c"
TEXT_COLOR = "#b0ccd8"

G_PER_LB = 453.592


def chart_layout(fig, title="", height=340):
    fig.update_layout(
        title=dict(text=title, font=dict(color="#e8f4e8", size=14), x=0.01, xanchor="left"),
        paper_bgcolor=CHART_BG, plot_bgcolor=CHART_BG,
        font=dict(color=TEXT_COLOR, family="sans-serif"),
        height=height, margin=dict(l=10, r=10, t=40 if title else 10, b=10),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=TEXT_COLOR, size=11)),
        xaxis=dict(gridcolor=GRID_COLOR, zeroline=False, tickfont=dict(color=TEXT_COLOR)),
        yaxis=dict(gridcolor=GRID_COLOR, zeroline=False, tickfont=dict(color=TEXT_COLOR)),
    )
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=300)
def load_data():
    wb = openpyxl.load_workbook("eden_supply_reports.xlsx", read_only=True, data_only=True)

    # ── 1. Sweed Inventory (AdvancedReport) ───────────────────────────────────
    ws = wb["AdvancedReport"]
    adv_rows = list(ws.iter_rows(values_only=True))
    adv_data = [r for r in adv_rows[11:] if r[0] is not None]
    inv_records = []
    for r in adv_data:
        inv_records.append({
            "product":  r[0], "category": r[1] or "Other", "subcat": r[2], "type": r[3],
            "size": r[6], "price": float(r[7] or 0), "brand": r[8], "quality": r[9],
            "expiry": r[10], "cbd_pct": r[11], "thc_pct": r[12],
            "location": r[13], "org": r[15] or "Unknown", "qty": int(r[16] or 0),
        })
    inv_df = pd.DataFrame(inv_records)

    # ── 2. Biotrack Inventory ─────────────────────────────────────────────────
    ws2 = wb["Biotrack inventory"]
    bio_rows = list(ws2.iter_rows(values_only=True))
    bio_records = []
    for r in bio_rows[1:]:
        bio_records.append({
            "location": r[0], "room": r[1], "product": r[3], "inv_id": r[4],
            "count": float(r[5] or 0), "category": r[6], "strain": r[7], "type": r[8],
        })
    bio_df = pd.DataFrame(bio_records)

    # ── 3. Sweed Sales ────────────────────────────────────────────────────────
    ws3 = wb["Sweed-Sales YTD"]
    sw_rows = list(ws3.iter_rows(values_only=True))
    sw_records = []
    for r in sw_rows[1:]:
        if not r[1]: continue
        sw_records.append({
            "store": r[0] or "Unknown", "ts": r[1], "category": r[2] or "Other",
            "gross": float(r[4] or 0), "net": float(r[5] or 0), "qty": float(r[6] or 0),
        })
    sw_df = pd.DataFrame(sw_records)
    if not sw_df.empty:
        sw_df["date"] = pd.to_datetime(sw_df["ts"])
        sw_df["week"]  = sw_df["date"].dt.to_period("W").apply(lambda x: x.start_time)
        sw_df["month"] = sw_df["date"].dt.to_period("M").apply(lambda x: x.start_time)

    # ── 4. Alleaves Sales ─────────────────────────────────────────────────────
    ws4 = wb["Alleaves Sales -YTD"]
    al_rows = list(ws4.iter_rows(values_only=True))
    al_records = []
    for r in al_rows[2:]:
        if not r[3]: continue
        al_records.append({
            "location": r[0] or "Unknown", "category": (r[1] or "").strip(),
            "product": r[3], "qty": float(r[4] or 0),
            "gross": float(r[6] or 0), "net": float(r[7] or 0),
        })
    al_df = pd.DataFrame(al_records)

    # ── 5. Harvest Tracker ────────────────────────────────────────────────────
    HARVEST_PATH = "harvest_tracker.xlsx"
    wb2 = openpyxl.load_workbook(HARVEST_PATH, read_only=True, data_only=True)
    ws5 = wb2["harvest tracker"]
    h_rows = list(ws5.iter_rows(values_only=True))
    h_records = []
    for r in h_rows[1:]:
        r = list(r) + [None] * 40
        h_records.append({
            "strain":           r[0],
            "harvest_date":     r[2] if hasattr(r[2], "year") else None,
            "status":           r[3],
            "cycle":            r[4],
            "room":             r[5],
            "num_plants":       float(r[6])  if r[6]  else None,
            "fresh_frozen_g":   float(r[7])  if r[7]  else None,
            "wet_weight_g":     float(r[9])  if r[9]  else None,
            "dry_weight_g":     float(r[12]) if r[12] else None,
            "sf_canopy":        float(r[13]) if r[13] else None,
            "grams_per_sf":     float(r[14]) if r[14] else None,
            "trim_weight_g":    float(r[15]) if r[15] else None,
            "trimmed_flower_g": float(r[16]) if r[16] else None,
            "est_35_units":     float(r[18]) if r[18] else None,
            "est_1g_units":     float(r[19]) if r[19] else None,
            "designation":      r[21],
            "thc_pct":          float(r[24]) if r[24] else None,
            "terps_pct":        float(r[25]) if r[25] else None,
            "test_date":        r[26] if hasattr(r[26], "year") else None,
            "grade":            r[27],
        })
    harvest_df = pd.DataFrame(h_records)

    return inv_df, bio_df, sw_df, al_df, harvest_df


inv_df, bio_df, sw_df, al_df, harvest_df = load_data()


# ══════════════════════════════════════════════════════════════════════════════
# PRE-COMPUTE AGGREGATES (supply & demand)
# ══════════════════════════════════════════════════════════════════════════════
CAT_MAP = {
    "Flower": "Flower", "Flower ": "Flower",
    "Live Rosin": "Concentrates", "Live Rosin (Disposable)": "Carts",
    "Edibles": "Edibles", "Tincture": "Tinctures", "Topical": "Topical",
}
al_df["cat_norm"] = al_df["category"].map(CAT_MAP).fillna(al_df["category"])
LOC_MAP = {
    "EDEN Tampa": "Eden - Tampa", "EDEN Sarasota": "Eden - Sarasota",
    "Tavares Delivery": "Eden - Delivery Hub",
}
al_df["loc_norm"] = al_df["location"].map(LOC_MAP).fillna(al_df["location"])

ytd_gross = al_df["gross"].sum() + sw_df["gross"].sum()
ytd_net   = al_df["net"].sum()   + sw_df["net"].sum()
ytd_qty   = al_df["qty"].sum()   + sw_df["qty"].sum()
sw_total_gross = sw_df["gross"].sum()

EXCLUDE_LOCS = {"Quarantine", "Quarantine 1283", "temp hoLDING"}
fin_df = inv_df[~inv_df["location"].isin(EXCLUDE_LOCS)]
total_finished = fin_df["qty"].sum()

# Biotrack pipeline
PIPELINE_ROOMS = {
    "02.1 - Hash Processing Room":             "Concentrate Processing",
    "02.2 - Hash Washing Room (Freezer)":      "Concentrate Processing",
    "02.3 - Hash Processing Room (WIP)":       "Concentrate Processing",
    "02.4 - Hash Processing Room Storage Cabinet": "Concentrate Processing",
    "03 - Trimming (WIP)":                     "Trimming",
    "03.1 - Trim Room":                        "Trimming",
    "TAV-PROD-002 - Kitchen (WIP)":            "Edible Production",
    "TAV-PROD-004-Kitchen Storage":            "Edible Production",
    "04 - Ready for Packaging":                "Ready for Packaging",
    "05 - Packaging (WIP)":                    "Packaging WIP",
    "12 - CGMP Packaging Room":                "CGMP Packaging",
    "12.1 CGMP Finished Product Fridge":       "CGMP Finished",
    "12.2 CGMP WIP Fridge":                    "CGMP WIP",
    "07 - Awaiting COA":                       "Awaiting Test Results",
    "09 - Approved for Labeling":              "Approved for Labeling ✓",
    "13 - Processing Vault":                   "Processing Vault",
}
bio_df["stage"] = bio_df["room"].map(PIPELINE_ROOMS)

BIO_CAT_MAP = {
    "Whole Flower": "Flower", "Trim/Shake": "Flower",
    "Pre-roll Material": "Flower", "Pre-Rolls": "Pre-rolls",
    "Hash": "Concentrates", "Fresh Frozen Material": "Concentrates",
    "Rosin": "Concentrates", "Live Rosin": "Concentrates", "Kief": "Concentrates",
    "510 Thread": "Carts", "510-Threads": "Carts", "Rosin (Cart)": "Carts",
    "Edible": "Edibles", "Tincture": "Tinctures", "Syringe": "Concentrates",
}
bio_df["broad_cat"] = bio_df["category"].map(BIO_CAT_MAP).fillna("Other")

FLOWER_CATS_BIO = {"Whole Flower", "Trim/Shake", "Pre-roll Material", "Fresh Frozen Material"}
MFG_CATS_BIO    = {"Hash", "Live Rosin", "Rosin", "Kief"}
G_PER_L = 1000.0

def _bio_convert(row):
    cat = row["category"] or ""
    g = row["count"]
    if cat in FLOWER_CATS_BIO: return g / G_PER_LB, "lbs"
    if cat in MFG_CATS_BIO:    return g / G_PER_L,  "L"
    return g, "units"

bio_df[["display_count", "unit"]] = bio_df.apply(_bio_convert, axis=1, result_type="expand")

STAGE_ORDER = [
    "Concentrate Processing", "Edible Production", "Trimming",
    "Ready for Packaging", "Packaging WIP", "CGMP WIP", "CGMP Packaging",
    "CGMP Finished", "Awaiting Test Results", "Approved for Labeling ✓", "Processing Vault",
]

VAULT_LOCS = {"TAV-Stock Room"}
DISP_ORGS  = {
    "Eden - Tampa": "Tampa", "Eden - Sarasota": "Sarasota",
    "Eden - Cocoa Beach": "Cocoa Beach", "Eden - Orlando": "Orlando",
}
vault_df = fin_df[fin_df["location"].isin(VAULT_LOCS)]
disp_df  = fin_df[fin_df["org"].isin(DISP_ORGS.keys()) & ~fin_df["location"].isin(VAULT_LOCS)]

# Harvest pre-compute
h = harvest_df.copy()
h["harvest_date"] = pd.to_datetime(h["harvest_date"], errors="coerce")
h["dry_lbs"]  = h["dry_weight_g"]  / G_PER_LB
h["wet_lbs"]  = h["wet_weight_g"]  / G_PER_LB
h["ff_lbs"]   = h["fresh_frozen_g"] / G_PER_LB
h["trim_lbs"] = h["trim_weight_g"] / G_PER_LB
h["flower_lbs"] = h["trimmed_flower_g"] / G_PER_LB
h["month"] = h["harvest_date"].dt.to_period("M").apply(
    lambda x: x.start_time if pd.notna(x) else pd.NaT
)

harvested = h[h["status"] == "Harvested"].copy()
active     = h[h["status"].isin(["Flowering", "Moved to Trim Room"])].copy()
upcoming   = h[h["status"] == "Upcoming Cycle"].copy()


# ══════════════════════════════════════════════════════════════════════════════
# HEADER  (always visible above tabs)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="eden-header">
  <h1>🌿 Eden Cannabis — Operations Dashboard</h1>
  <p>CEO Overview &nbsp;|&nbsp; Vertical Operations &nbsp;|&nbsp;
     Data as of <strong>June 26, 2026</strong></p>
</div>
""", unsafe_allow_html=True)

approved_df = bio_df[bio_df["room"] == "09 - Approved for Labeling"].copy()
if not approved_df.empty:
    products = approved_df.groupby("product")["count"].sum()
    prod_list = ", ".join(f"<strong>{p}</strong> ({int(q)} units)" for p, q in products.items())
    st.markdown(f"""
    <div class="alert-box">
      <div class="alert-title">🟢 NEW — Approved for Labeling Today</div>
      <div class="alert-body">
        Passing COA results received today. Ready to ship to dispensaries.<br>{prod_list}
      </div>
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab1, tab2 = st.tabs(["📦  Supply & Demand", "🌱  Harvest"])


# ─────────────────────────────────────────────────────────────────────────────
with tab1:
# ─────────────────────────────────────────────────────────────────────────────

    # ── KPIs ─────────────────────────────────────────────────────────────────
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f"""<div class="kpi-card">
          <div class="kpi-label">YTD Gross Revenue</div>
          <div class="kpi-value">${ytd_gross:,.0f}</div>
          <div class="kpi-sub">Jan 1 – Jun 26, 2026</div>
        </div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""<div class="kpi-card">
          <div class="kpi-label">YTD Net Revenue</div>
          <div class="kpi-value">${ytd_net:,.0f}</div>
          <div class="kpi-sub">After discounts &amp; adjustments</div>
        </div>""", unsafe_allow_html=True)
    with k3:
        st.markdown(f"""<div class="kpi-card">
          <div class="kpi-label">YTD Units Sold</div>
          <div class="kpi-value">{ytd_qty:,.0f}</div>
          <div class="kpi-sub">All locations combined</div>
        </div>""", unsafe_allow_html=True)
    with k4:
        st.markdown(f"""<div class="kpi-card">
          <div class="kpi-label">Finished Units On-Hand</div>
          <div class="kpi-value">{total_finished:,}</div>
          <div class="kpi-sub">All dispensaries + delivery hub</div>
        </div>""", unsafe_allow_html=True)
    with k5:
        pipeline_bio = bio_df[
            bio_df["stage"].notna() &
            ~bio_df["stage"].str.contains("Approved|Sample", na=False)
        ]
        flower_lbs = pipeline_bio[pipeline_bio["unit"] == "lbs"]["display_count"].sum()
        mfg_L      = pipeline_bio[pipeline_bio["unit"] == "L"]["display_count"].sum()
        st.markdown(f"""<div class="kpi-card">
          <div class="kpi-label">Bulk in Pipeline</div>
          <div class="kpi-value" style="font-size:1.3rem">{flower_lbs:,.0f} lbs</div>
          <div class="kpi-sub">flower &amp; fresh frozen</div>
          <div class="kpi-value" style="font-size:1.3rem; margin-top:6px">{mfg_L:,.1f} L</div>
          <div class="kpi-sub">manufactured / concentrate</div>
        </div>""", unsafe_allow_html=True)

    # ── Production pipeline ───────────────────────────────────────────────────
    st.markdown('<div class="section-header">🏭 Supply — Production Pipeline (Tavares Farm)</div>',
                unsafe_allow_html=True)

    pipe_df = bio_df[bio_df["stage"].notna()].copy()

    def _stage_bar(df, unit_label, title, height=340):
        agg = df.groupby(["stage", "category"])["display_count"].sum().reset_index()
        agg["order"] = agg["stage"].map({s: i for i, s in enumerate(STAGE_ORDER)}).fillna(99)
        agg = agg.sort_values("order")
        fig = go.Figure()
        for cat in agg["category"].unique():
            d = agg[agg["category"] == cat]
            fig.add_trace(go.Bar(
                name=cat, x=d["stage"], y=d["display_count"],
                marker_color=COLORS.get(BIO_CAT_MAP.get(cat, "Other"), "#888"),
                hovertemplate=f"%{{x}}<br>%{{y:,.2f}} {unit_label}<extra>{cat}</extra>",
            ))
        fig.update_layout(barmode="stack", xaxis_tickangle=-35, yaxis_title=unit_label,
                          legend=dict(orientation="h", y=1.08, font=dict(size=10)))
        return chart_layout(fig, title, height=height)

    flower_pipe = pipe_df[pipe_df["unit"] == "lbs"]
    mfg_pipe    = pipe_df[pipe_df["unit"] == "L"]
    unit_pipe   = pipe_df[pipe_df["unit"] == "units"]

    col_fl, col_mfg = st.columns(2)
    with col_fl:
        st.plotly_chart(
            _stage_bar(flower_pipe, "lbs",
                       f"Flower Pipeline — {flower_pipe['display_count'].sum():,.0f} lbs total"),
            use_container_width=True)
    with col_mfg:
        st.plotly_chart(
            _stage_bar(mfg_pipe, "L",
                       f"Manufactured Pipeline — {mfg_pipe['display_count'].sum():,.2f} L total"),
            use_container_width=True)

    if not unit_pipe.empty:
        unit_agg = unit_pipe.groupby(["stage", "category"])["display_count"].sum().reset_index()
        unit_agg["order"] = unit_agg["stage"].map(
            {s: i for i, s in enumerate(STAGE_ORDER)}).fillna(99)
        unit_agg = unit_agg.sort_values("order")
        fig_u = go.Figure()
        for cat in unit_agg["category"].unique():
            d = unit_agg[unit_agg["category"] == cat]
            fig_u.add_trace(go.Bar(
                name=cat, x=d["stage"], y=d["display_count"],
                marker_color=COLORS.get(BIO_CAT_MAP.get(cat, "Other"), "#888"),
                hovertemplate="%{x}<br>%{y:,.0f} units<extra>" + cat + "</extra>",
            ))
        fig_u.update_layout(barmode="stack", xaxis_tickangle=-35, yaxis_title="Units",
                             legend=dict(orientation="h", y=1.08, font=dict(size=10)))
        chart_layout(fig_u, f"Packaged / Counted Items — {unit_pipe['display_count'].sum():,.0f} units total",
                     height=300)
        st.plotly_chart(fig_u, use_container_width=True)

    with st.expander("Pipeline Stage Detail"):
        def _fmt(row):
            if row["unit"] == "lbs": return f"{row['display_count']:,.2f} lbs"
            if row["unit"] == "L":   return f"{row['display_count']:,.3f} L"
            return f"{row['display_count']:,.0f} units"
        sd = pipe_df.groupby(["stage", "category", "unit"]).agg(
            display_count=("display_count", "sum")).reset_index()
        sd["Amount"] = sd.apply(_fmt, axis=1)
        sd["order"]  = sd["stage"].map({s: i for i, s in enumerate(STAGE_ORDER)}).fillna(99)
        sd = sd.sort_values(["order", "display_count"], ascending=[True, False])
        st.dataframe(sd[["stage", "category", "Amount"]].rename(
            columns={"stage": "Stage", "category": "Category"}),
            use_container_width=True, hide_index=True)

    # ── Finished goods inventory ──────────────────────────────────────────────
    st.markdown('<div class="section-header">🏪 Supply — Finished Goods Inventory (Sweed)</div>',
                unsafe_allow_html=True)

    col_disp, col_vault = st.columns([3, 2])
    with col_disp:
        disp_agg = disp_df.copy()
        disp_agg["store"] = disp_agg["org"].map(DISP_ORGS)
        disp_agg = disp_agg.groupby(["store", "category"])["qty"].sum().reset_index()
        store_order = ["Tampa", "Sarasota", "Cocoa Beach", "Orlando"]
        cat_order_d = ["Flower", "Concentrates", "Edibles", "Carts", "Pre-rolls", "Tinctures", "Topical"]
        fig3 = go.Figure()
        for cat in cat_order_d:
            d = disp_agg[disp_agg["category"] == cat]
            fig3.add_trace(go.Bar(
                name=cat, x=d["store"], y=d["qty"],
                marker_color=COLORS.get(cat, "#888"),
                hovertemplate="%{x} — " + cat + "<br>%{y:,} units<extra></extra>",
            ))
        fig3.update_layout(barmode="stack", yaxis_title="Units",
                           legend=dict(orientation="h", y=1.08),
                           xaxis=dict(categoryorder="array", categoryarray=store_order))
        chart_layout(fig3, "Dispensary Inventory by Location & Category", height=340)
        st.plotly_chart(fig3, use_container_width=True)

    with col_vault:
        vault_agg   = vault_df.groupby("category")["qty"].sum().reset_index().sort_values("qty", ascending=False)
        vault_total = vault_agg["qty"].sum()
        fig4 = go.Figure(go.Pie(
            labels=vault_agg["category"], values=vault_agg["qty"], hole=0.52,
            marker=dict(colors=[COLORS.get(c, "#888") for c in vault_agg["category"]]),
            textinfo="label+percent", textfont=dict(color=TEXT_COLOR, size=11),
            hovertemplate="%{label}<br>%{value:,} units<br>%{percent}<extra></extra>",
        ))
        fig4.add_annotation(text=f"{vault_total:,}<br>units", x=0.5, y=0.5,
                            showarrow=False, font=dict(color="#e8f4e8", size=13))
        chart_layout(fig4, f"Processing Vault Back-Stock ({vault_total:,} units)", height=340)
        st.plotly_chart(fig4, use_container_width=True)

    with st.expander("Full Inventory Detail by Location & Category"):
        inv_sum = fin_df.groupby(["org", "category"])["qty"].sum().reset_index()
        inv_sum.columns = ["Location", "Category", "Units"]
        inv_sum["Units"] = inv_sum["Units"].apply(lambda x: f"{x:,}")
        st.dataframe(inv_sum, use_container_width=True, hide_index=True)

    # ── YTD Sales ─────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">📈 Demand — YTD Sales (Jan 1 – Jun 26, 2026)</div>',
                unsafe_allow_html=True)
    st.markdown("""<p style="font-size:0.8rem; color:#5a8fa0; margin-bottom:12px;">
      <strong>Data sources:</strong> Alleaves POS (Jan 1 – Apr 2) covers Tampa, Sarasota &amp; Delivery.
      Sweed POS (Mar 16 – Jun 26) reflects Tampa only in this export.</p>""",
      unsafe_allow_html=True)

    col_cat, col_loc = st.columns(2)
    with col_cat:
        sw_cat = sw_df.groupby("category").agg(gross=("gross","sum"), qty=("qty","sum")).reset_index()
        al_cat = al_df.groupby("cat_norm").agg(gross=("gross","sum"), qty=("qty","sum")).reset_index()
        al_cat.rename(columns={"cat_norm": "category"}, inplace=True)
        cat_total = pd.concat([sw_cat, al_cat]).groupby("category").agg(
            gross=("gross","sum"), qty=("qty","sum")).reset_index().sort_values("gross", ascending=True)
        fig5 = go.Figure(go.Bar(
            x=cat_total["gross"], y=cat_total["category"], orientation="h",
            marker_color=[COLORS.get(c, "#2e7bc4") for c in cat_total["category"]],
            text=cat_total["gross"].apply(lambda x: f"${x:,.0f}"),
            textposition="outside", textfont=dict(color=TEXT_COLOR, size=11),
            hovertemplate="%{y}<br>Gross: $%{x:,.0f}<extra></extra>",
        ))
        chart_layout(fig5, "YTD Gross Sales by Category", height=320)
        fig5.update_layout(xaxis_title="Gross Sales ($)", yaxis=dict(gridcolor="rgba(0,0,0,0)"))
        st.plotly_chart(fig5, use_container_width=True)

    with col_loc:
        al_loc = al_df.groupby("loc_norm").agg(gross=("gross","sum"), qty=("qty","sum")).reset_index()
        sw_loc = sw_df.groupby("store").agg(gross=("gross","sum"), qty=("qty","sum")).reset_index()
        sw_loc.rename(columns={"store": "loc_norm"}, inplace=True)
        al_loc["period"] = "Jan–Apr (Alleaves)"
        sw_loc["period"] = "Mar–Jun (Sweed, Tampa)"
        loc_all = pd.concat([al_loc, sw_loc])
        fig6 = go.Figure()
        for period, color in [("Jan–Apr (Alleaves)", "#2e7bc4"), ("Mar–Jun (Sweed, Tampa)", "#4a9e6b")]:
            d = loc_all[loc_all["period"] == period]
            fig6.add_trace(go.Bar(name=period, x=d["loc_norm"], y=d["gross"],
                                  marker_color=color,
                                  hovertemplate="%{x}<br>$%{y:,.0f}<extra>" + period + "</extra>"))
        fig6.update_layout(barmode="group", yaxis_title="Gross Sales ($)",
                           legend=dict(orientation="h", y=1.08))
        chart_layout(fig6, "YTD Sales by Location", height=320)
        st.plotly_chart(fig6, use_container_width=True)

    st.markdown('<div class="section-header">📅 Tampa Sales Trend (Sweed, Mar 16 – Jun 26)</div>',
                unsafe_allow_html=True)
    col_trend, col_pie = st.columns([3, 2])
    with col_trend:
        weekly = sw_df.groupby(["week","category"]).agg(gross=("gross","sum")).reset_index()
        fig7 = go.Figure()
        for cat in ["Flower","Concentrates","Edibles","Carts","Pre-rolls","Tinctures","Topical"]:
            d = weekly[weekly["category"] == cat].sort_values("week")
            fig7.add_trace(go.Scatter(
                x=d["week"], y=d["gross"], name=cat, mode="lines",
                line=dict(color=COLORS.get(cat,"#888"), width=2), stackgroup="one",
                hovertemplate="%{x|%b %d}<br>$%{y:,.0f}<extra>" + cat + "</extra>",
            ))
        fig7.update_layout(yaxis_title="Gross Sales ($)", legend=dict(orientation="h", y=1.08))
        chart_layout(fig7, "Weekly Gross Sales by Category (Tampa)", height=340)
        st.plotly_chart(fig7, use_container_width=True)

    with col_pie:
        sw_ct = sw_df.groupby("category")["gross"].sum().reset_index().sort_values("gross", ascending=False)
        fig8 = go.Figure(go.Pie(
            labels=sw_ct["category"], values=sw_ct["gross"], hole=0.52,
            marker=dict(colors=[COLORS.get(c,"#888") for c in sw_ct["category"]]),
            textinfo="label+percent", textfont=dict(color=TEXT_COLOR, size=11),
            hovertemplate="%{label}<br>$%{value:,.0f}<br>%{percent}<extra></extra>",
        ))
        fig8.add_annotation(text=f"${sw_total_gross:,.0f}<br>Tampa", x=0.5, y=0.5,
                            showarrow=False, font=dict(color="#e8f4e8", size=12))
        chart_layout(fig8, "Tampa Sales Mix (Sweed)", height=340)
        st.plotly_chart(fig8, use_container_width=True)

    # ── Supply vs Demand ──────────────────────────────────────────────────────
    st.markdown('<div class="section-header">⚖️ Supply vs. Demand Snapshot</div>',
                unsafe_allow_html=True)
    floor_df = disp_df.groupby("category")["qty"].sum().reset_index()
    floor_df.columns = ["category", "on_hand"]
    cutoff = sw_df["date"].max() - pd.Timedelta(weeks=4)
    velocity_4w = sw_df[sw_df["date"] >= cutoff].groupby("category")["qty"].sum().reset_index()
    velocity_4w.columns = ["category", "sold_4w"]
    velocity_4w["weekly_velocity"] = velocity_4w["sold_4w"] / 4
    svd = floor_df.merge(velocity_4w, on="category", how="outer").fillna(0)
    svd["weeks_of_stock"] = svd.apply(
        lambda r: round(r["on_hand"] / r["weekly_velocity"], 1) if r["weekly_velocity"] > 0 else None,
        axis=1)

    col_svd, col_svd_tbl = st.columns([3, 2])
    with col_svd:
        fig9 = make_subplots(specs=[[{"secondary_y": True}]])
        fig9.add_trace(go.Bar(
            name="Dispensary Units On-Hand", x=svd["category"], y=svd["on_hand"],
            marker_color=[COLORS.get(c, "#888") for c in svd["category"]], opacity=0.85,
            hovertemplate="%{x}<br>%{y:,} units on hand<extra></extra>",
        ), secondary_y=False)
        fig9.add_trace(go.Scatter(
            name="Wkly Velocity (Tampa, last 4w)", x=svd["category"], y=svd["weekly_velocity"],
            mode="markers+lines", marker=dict(size=10, color="#f5c842", symbol="diamond"),
            line=dict(color="#f5c842", width=2, dash="dot"),
            hovertemplate="%{x}<br>~%{y:.0f} units/wk<extra></extra>",
        ), secondary_y=True)
        fig9.update_yaxes(title_text="Units On-Hand", secondary_y=False,
                          gridcolor=GRID_COLOR, tickfont=dict(color=TEXT_COLOR))
        fig9.update_yaxes(title_text="Units/Week (Tampa)", secondary_y=True,
                          gridcolor="rgba(0,0,0,0)", tickfont=dict(color="#f5c842"))
        fig9.update_layout(
            paper_bgcolor=CHART_BG, plot_bgcolor=CHART_BG, font=dict(color=TEXT_COLOR),
            height=340, margin=dict(l=10, r=10, t=40, b=10),
            legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=TEXT_COLOR, size=11)),
            title=dict(text="Dispensary Stock vs. Sales Velocity (Tampa proxy)",
                       font=dict(color="#e8f4e8", size=14), x=0.01),
            xaxis=dict(gridcolor=GRID_COLOR, tickfont=dict(color=TEXT_COLOR)),
        )
        st.plotly_chart(fig9, use_container_width=True)

    with col_svd_tbl:
        svd_d = svd[["category","on_hand","weekly_velocity","weeks_of_stock"]].copy()
        svd_d.columns = ["Category","Floor Units","Wkly Velocity*","Weeks of Stock*"]
        svd_d["Floor Units"]      = svd_d["Floor Units"].apply(lambda x: f"{int(x):,}")
        svd_d["Wkly Velocity*"]   = svd_d["Wkly Velocity*"].apply(lambda x: f"~{x:.0f}/wk" if x else "–")
        svd_d["Weeks of Stock*"]  = svd_d["Weeks of Stock*"].apply(lambda x: f"{x:.1f} wks" if x else "–")
        st.markdown('<p style="font-size:0.75rem;color:#5a8fa0;margin-bottom:8px;">'
                    '* Velocity = Tampa Sweed sales, last 4 weeks.</p>', unsafe_allow_html=True)
        st.dataframe(svd_d, use_container_width=True, hide_index=True)
        st.markdown('<p style="font-size:0.75rem;color:#5a8fa0;margin-top:12px;">'
                    '<strong>Processing vault</strong> holds an additional</p>', unsafe_allow_html=True)
        vt = vault_df.groupby("category")["qty"].sum().reset_index()
        vt.columns = ["Category", "Vault Units"]
        vt["Vault Units"] = vt["Vault Units"].apply(lambda x: f"{int(x):,}")
        st.dataframe(vt, use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────────────────────────────────────
with tab2:
# ─────────────────────────────────────────────────────────────────────────────

    # ── Harvest KPIs ─────────────────────────────────────────────────────────
    h1, h2, h3, h4, h5 = st.columns(5)
    total_cycles    = len(harvested)
    total_dry_lbs   = harvested["dry_lbs"].sum()
    total_ff_lbs    = harvested["ff_lbs"].sum()
    total_wet_lbs   = harvested["wet_lbs"].sum()
    avg_dry_ratio   = (total_dry_lbs / total_wet_lbs * 100) if total_wet_lbs > 0 else 0
    avg_per_plant   = (harvested["dry_weight_g"] / harvested["num_plants"]).dropna().mean() / G_PER_LB

    with h1:
        st.markdown(f"""<div class="kpi-card">
          <div class="kpi-label">Cycles Harvested</div>
          <div class="kpi-value">{total_cycles}</div>
          <div class="kpi-sub">Nov 2024 – Jun 2026</div>
        </div>""", unsafe_allow_html=True)
    with h2:
        st.markdown(f"""<div class="kpi-card">
          <div class="kpi-label">Total Dry Yield</div>
          <div class="kpi-value">{total_dry_lbs:,.0f} lbs</div>
          <div class="kpi-sub">All harvested cycles</div>
        </div>""", unsafe_allow_html=True)
    with h3:
        st.markdown(f"""<div class="kpi-card">
          <div class="kpi-label">Total Fresh Frozen</div>
          <div class="kpi-value">{total_ff_lbs:,.0f} lbs</div>
          <div class="kpi-sub">Rosin / live extract feedstock</div>
        </div>""", unsafe_allow_html=True)
    with h4:
        st.markdown(f"""<div class="kpi-card">
          <div class="kpi-label">Wet → Dry Ratio</div>
          <div class="kpi-value">{avg_dry_ratio:.1f}%</div>
          <div class="kpi-sub">Avg across all harvests</div>
        </div>""", unsafe_allow_html=True)
    with h5:
        st.markdown(f"""<div class="kpi-card">
          <div class="kpi-label">Avg Yield / Plant</div>
          <div class="kpi-value">{avg_per_plant:.2f} lbs</div>
          <div class="kpi-sub">Dry weight per plant</div>
        </div>""", unsafe_allow_html=True)

    # ── Monthly harvest output ────────────────────────────────────────────────
    st.markdown('<div class="section-header">📅 Monthly Harvest Output</div>',
                unsafe_allow_html=True)

    monthly_h = harvested.dropna(subset=["month"]).groupby("month").agg(
        dry_lbs=("dry_lbs", "sum"),
        ff_lbs=("ff_lbs", "sum"),
        trim_lbs=("trim_lbs", "sum"),
        cycles=("strain", "count"),
    ).reset_index().sort_values("month")

    col_m1, col_m2 = st.columns([3, 2])
    with col_m1:
        fig_mo = go.Figure()
        fig_mo.add_trace(go.Bar(
            name="Dry Flower (lbs)", x=monthly_h["month"], y=monthly_h["dry_lbs"],
            marker_color="#4a9e6b",
            hovertemplate="%{x|%b %Y}<br>%{y:,.0f} lbs dry<extra></extra>",
        ))
        fig_mo.add_trace(go.Bar(
            name="Fresh Frozen (lbs)", x=monthly_h["month"], y=monthly_h["ff_lbs"],
            marker_color="#2e7bc4",
            hovertemplate="%{x|%b %Y}<br>%{y:,.0f} lbs fresh frozen<extra></extra>",
        ))
        fig_mo.add_trace(go.Bar(
            name="Trim (lbs)", x=monthly_h["month"], y=monthly_h["trim_lbs"],
            marker_color="#c4a244",
            hovertemplate="%{x|%b %Y}<br>%{y:,.0f} lbs trim<extra></extra>",
        ))
        fig_mo.add_trace(go.Scatter(
            name="Cycles", x=monthly_h["month"], y=monthly_h["cycles"],
            mode="markers+lines", yaxis="y2",
            marker=dict(size=8, color="#f5c842"), line=dict(color="#f5c842", width=2, dash="dot"),
            hovertemplate="%{x|%b %Y}<br>%{y} cycles<extra></extra>",
        ))
        fig_mo.update_layout(
            barmode="stack", yaxis_title="lbs",
            yaxis2=dict(title="Cycles", overlaying="y", side="right",
                        gridcolor="rgba(0,0,0,0)", tickfont=dict(color="#f5c842")),
            legend=dict(orientation="h", y=1.08),
        )
        chart_layout(fig_mo, "Monthly Harvest — Dry, Fresh Frozen & Trim (lbs)", height=360)
        st.plotly_chart(fig_mo, use_container_width=True)

    with col_m2:
        # Designation breakdown of harvested cycles
        des_agg = harvested.groupby("designation").agg(
            dry_lbs=("dry_lbs","sum"), cycles=("strain","count")).reset_index()
        des_agg = des_agg[des_agg["designation"].notna() & (des_agg["designation"] != "")]
        des_colors = {"Flower": "#4a9e6b", "Wash": "#2e7bc4", "Pre-rolls": "#c4a244", "TBD": "#5a8fa0"}
        fig_des = go.Figure(go.Pie(
            labels=des_agg["designation"], values=des_agg["dry_lbs"], hole=0.52,
            marker=dict(colors=[des_colors.get(d, "#888") for d in des_agg["designation"]]),
            textinfo="label+percent", textfont=dict(color=TEXT_COLOR, size=11),
            hovertemplate="%{label}<br>%{value:,.0f} lbs<br>%{percent}<extra></extra>",
        ))
        fig_des.add_annotation(
            text=f"{des_agg['dry_lbs'].sum():,.0f}<br>lbs total",
            x=0.5, y=0.5, showarrow=False, font=dict(color="#e8f4e8", size=12))
        chart_layout(fig_des, "Harvest Designation Mix (by Dry lbs)", height=360)
        st.plotly_chart(fig_des, use_container_width=True)

    # ── Room performance ──────────────────────────────────────────────────────
    st.markdown('<div class="section-header">🏠 Flower Room Performance</div>',
                unsafe_allow_html=True)

    col_r1, col_r2 = st.columns(2)
    with col_r1:
        room_agg = harvested.dropna(subset=["room"]).groupby("room").agg(
            dry_lbs=("dry_lbs","sum"), ff_lbs=("ff_lbs","sum"),
            cycles=("strain","count"), plants=("num_plants","sum"),
        ).reset_index()
        room_agg["dry_per_plant"] = (room_agg["dry_lbs"] / room_agg["plants"]).round(3)

        fig_room = go.Figure()
        fig_room.add_trace(go.Bar(
            name="Dry Flower", x=room_agg["room"], y=room_agg["dry_lbs"],
            marker_color=[ROOM_COLORS.get(r, "#888") for r in room_agg["room"]],
            hovertemplate="%{x}<br>%{y:,.0f} lbs dry<extra></extra>",
        ))
        fig_room.add_trace(go.Bar(
            name="Fresh Frozen", x=room_agg["room"], y=room_agg["ff_lbs"],
            marker_color=[ROOM_COLORS.get(r, "#888") + "88" for r in room_agg["room"]],
            hovertemplate="%{x}<br>%{y:,.0f} lbs fresh frozen<extra></extra>",
        ))
        fig_room.update_layout(barmode="group", yaxis_title="lbs",
                                legend=dict(orientation="h", y=1.08))
        chart_layout(fig_room, "Total Output by Flower Room (all cycles)", height=340)
        st.plotly_chart(fig_room, use_container_width=True)

    with col_r2:
        # Dry lbs per plant by room — efficiency metric
        fig_eff = go.Figure(go.Bar(
            x=room_agg["room"], y=room_agg["dry_per_plant"],
            marker_color=[ROOM_COLORS.get(r, "#888") for r in room_agg["room"]],
            text=room_agg["dry_per_plant"].apply(lambda x: f"{x:.3f} lbs"),
            textposition="outside", textfont=dict(color=TEXT_COLOR),
            hovertemplate="%{x}<br>%{y:.3f} lbs/plant<extra></extra>",
        ))
        chart_layout(fig_eff, "Avg Dry Yield per Plant by Room (lbs)", height=340)
        fig_eff.update_layout(yaxis_title="lbs / plant")
        st.plotly_chart(fig_eff, use_container_width=True)

    # ── Top strains ───────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">🌿 Strain Performance</div>',
                unsafe_allow_html=True)

    col_s1, col_s2 = st.columns([3, 2])
    with col_s1:
        strain_agg = harvested.groupby("strain").agg(
            dry_lbs=("dry_lbs","sum"), wet_lbs=("wet_lbs","sum"),
            cycles=("cycle","count"), plants=("num_plants","sum"),
        ).reset_index()
        strain_agg["dry_ratio"] = (strain_agg["dry_lbs"] / strain_agg["wet_lbs"] * 100).round(1)
        strain_agg["lbs_per_plant"] = (strain_agg["dry_lbs"] / strain_agg["plants"]).round(3)
        strain_top = strain_agg.sort_values("dry_lbs", ascending=True).tail(20)

        fig_st = go.Figure()
        fig_st.add_trace(go.Bar(
            name="Dry lbs", y=strain_top["strain"], x=strain_top["dry_lbs"],
            orientation="h", marker_color="#4a9e6b",
            hovertemplate="%{y}<br>%{x:,.0f} lbs dry (%{customdata:.1f}% of wet)"
                          "<extra></extra>",
            customdata=strain_top["dry_ratio"],
        ))
        chart_layout(fig_st, "Top 20 Strains by Total Dry Yield (lbs)", height=480)
        fig_st.update_layout(xaxis_title="Dry lbs", yaxis=dict(gridcolor="rgba(0,0,0,0)"))
        st.plotly_chart(fig_st, use_container_width=True)

    with col_s2:
        # THC % for the strains that have test data
        thc_data = harvested.dropna(subset=["thc_pct"]).copy()
        thc_data = thc_data.sort_values("thc_pct", ascending=True)
        if not thc_data.empty:
            fig_thc = go.Figure()
            fig_thc.add_trace(go.Bar(
                y=thc_data["strain"], x=thc_data["thc_pct"], orientation="h",
                marker_color="#7b5ea7",
                text=thc_data.apply(
                    lambda r: f"{r['thc_pct']:.1f}% THC  {r['terps_pct']:.2f}% terps"
                              if r['terps_pct'] else f"{r['thc_pct']:.1f}% THC",
                    axis=1),
                textposition="outside", textfont=dict(color=TEXT_COLOR, size=11),
                hovertemplate="%{y}<br>THC: %{x:.1f}%<extra></extra>",
            ))
            chart_layout(fig_thc, "COA Results — THC & Terpenes %", height=280)
            fig_thc.update_layout(xaxis_title="THC %", yaxis=dict(gridcolor="rgba(0,0,0,0)"),
                                  xaxis=dict(range=[0, max(thc_data["thc_pct"]) * 1.35]))
            st.plotly_chart(fig_thc, use_container_width=True)

        # Wet-to-dry scatter for individual cycles
        scatter_data = harvested.dropna(subset=["wet_lbs", "dry_lbs"]).copy()
        scatter_data = scatter_data[scatter_data["wet_lbs"] > 0]
        fig_scat = go.Figure()
        for room in scatter_data["room"].dropna().unique():
            d = scatter_data[scatter_data["room"] == room]
            fig_scat.add_trace(go.Scatter(
                x=d["wet_lbs"], y=d["dry_lbs"], name=room, mode="markers",
                marker=dict(size=8, color=ROOM_COLORS.get(room, "#888"), opacity=0.8),
                hovertemplate="%{text}<br>Wet: %{x:,.0f} lbs<br>Dry: %{y:,.0f} lbs<extra></extra>",
                text=d["strain"],
            ))
        chart_layout(fig_scat, "Wet vs. Dry Weight per Cycle (by Room)", height=280)
        fig_scat.update_layout(xaxis_title="Wet lbs", yaxis_title="Dry lbs",
                               legend=dict(orientation="h", y=1.12))
        st.plotly_chart(fig_scat, use_container_width=True)

    # ── Active & upcoming cycles ──────────────────────────────────────────────
    st.markdown('<div class="section-header">🗓️ Active & Upcoming Cycles</div>',
                unsafe_allow_html=True)

    col_a1, col_a2 = st.columns(2)
    with col_a1:
        st.markdown("**Currently Flowering / In Trim**")
        active_display = active[["strain","status","harvest_date","room","cycle"]].copy()
        active_display["harvest_date"] = active_display["harvest_date"].dt.strftime("%b %d, %Y")
        active_display.columns = ["Strain","Status","Est. Harvest","Room","Cycle"]
        active_display = active_display.sort_values("Est. Harvest").fillna("–")
        st.dataframe(active_display, use_container_width=True, hide_index=True)

    with col_a2:
        st.markdown("**Upcoming Planned Cycles**")
        upcoming_display = upcoming[["strain","harvest_date"]].copy()
        upcoming_display["harvest_date"] = upcoming_display["harvest_date"].dt.strftime("%b %d, %Y")
        upcoming_display.columns = ["Strain","Planned Harvest"]
        upcoming_display = upcoming_display.sort_values("Planned Harvest").fillna("–")
        st.dataframe(upcoming_display, use_container_width=True, hide_index=True)

    # ── Full harvest log ──────────────────────────────────────────────────────
    with st.expander("Full Harvest Log"):
        log = harvested[[
            "strain","harvest_date","room","cycle","designation",
            "dry_lbs","ff_lbs","trim_lbs","thc_pct","terps_pct","num_plants",
        ]].copy()
        log["harvest_date"] = log["harvest_date"].dt.strftime("%b %d, %Y")
        log["dry_lbs"]  = log["dry_lbs"].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "–")
        log["ff_lbs"]   = log["ff_lbs"].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "–")
        log["trim_lbs"] = log["trim_lbs"].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "–")
        log["thc_pct"]  = log["thc_pct"].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "–")
        log["terps_pct"]= log["terps_pct"].apply(lambda x: f"{x:.3f}%" if pd.notna(x) else "–")
        log["num_plants"]= log["num_plants"].apply(lambda x: f"{int(x)}" if pd.notna(x) else "–")
        log.columns = ["Strain","Harvest Date","Room","Cycle","Designation",
                       "Dry (lbs)","Fresh Frozen (lbs)","Trim (lbs)","THC %","Terps %","Plants"]
        log = log.sort_values("Harvest Date", ascending=False)
        st.dataframe(log, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# FOOTER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("""
<p style="font-size:0.75rem; color:#3a6070; text-align:center;">
  Eden Cannabis · Operations Dashboard · Confidential · June 26 2026<br>
  Sources: Biotrack (farm bulk inventory), Sweed POS (dispensary inventory + Tampa sales),
  Alleaves POS (all-location sales Jan–Apr 2026), Harvest Tracker (internal)
</p>
""", unsafe_allow_html=True)
