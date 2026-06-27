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
  [data-testid="stAppViewContainer"] { background: #f0f6f0; }
  [data-testid="stHeader"] { background: rgba(240,246,240,0.9); }
  .main .block-container { padding-top: 1.5rem; max-width: 1400px; margin: 0 auto; }
  #MainMenu { display: none; }
  [data-testid="stFooter"] { display: none; }

  .kpi-card {
    background: #ffffff;
    border-radius: 14px;
    padding: 20px 16px;
    box-shadow: 0 1px 5px rgba(0,0,0,0.08);
    text-align: center;
  }
  .kpi-hero {
    background: #1a4731;
    border-radius: 14px;
    padding: 20px 16px;
    box-shadow: 0 4px 14px rgba(26,71,49,0.22);
    text-align: center;
  }
  .kpi-label      { color: #6b7280; font-size: 0.69rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.09em; margin-bottom: 10px; }
  .kpi-label-hero { color: rgba(255,255,255,0.58); font-size: 0.69rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.09em; margin-bottom: 10px; }
  .kpi-value      { color: #111827; font-size: 1.8rem; font-weight: 700; line-height: 1; }
  .kpi-value-hero { color: #ffffff; font-size: 1.8rem; font-weight: 700; line-height: 1; }
  .kpi-sub        { color: #9ca3af; font-size: 0.69rem; margin-top: 7px; }
  .kpi-sub-hero   { color: rgba(255,255,255,0.46); font-size: 0.69rem; margin-top: 7px; }

  .alert-box {
    background: #d1fae5;
    border-left: 3px solid #059669;
    border-radius: 10px;
    padding: 12px 16px;
    margin-bottom: 18px;
  }
  .alert-title { color: #064e3b; font-weight: 700; font-size: 0.82rem; }
  .alert-body  { color: #065f46; font-size: 0.79rem; margin-top: 3px; }

  .story-banner {
    background: #ffffff;
    border-radius: 14px;
    padding: 16px 22px;
    box-shadow: 0 1px 5px rgba(0,0,0,0.08);
    margin-bottom: 16px;
    border-left: 4px solid #1a4731;
  }
  .story-head { color: #111827; font-size: 1.05rem; font-weight: 700; margin: 0 0 4px 0; }
  .story-sub  { color: #6b7280; font-size: 0.80rem; margin: 0; }

  div[data-testid="stRadio"] { border-bottom: 1px solid #e5e7eb; margin-bottom: 20px; }
  div[data-testid="stRadio"] > div { gap: 0 !important; }
  div[data-testid="stRadio"] label {
    display: inline-flex; align-items: center; gap: 0;
    padding: 8px 24px 10px 24px;
    border-bottom: 2px solid transparent;
    margin-bottom: -1px;
    cursor: pointer;
    font-weight: 500; font-size: 0.92rem; color: #6b7280;
  }
  div[data-testid="stRadio"] label:has(input:checked) {
    color: #1a4731;
    border-bottom: 2px solid #1a4731;
  }
  div[data-testid="stRadio"] label span { margin-left: 0 !important; }
  div[data-testid="stRadio"] label input[type="radio"] { display: none; }

  h1 { color: #111827 !important; }
  h2, h3 { color: #1f2937 !important; }
  hr { border-color: #e5e7eb !important; }
  .data-card {
    background: #ffffff;
    border-radius: 14px;
    padding: 20px;
    box-shadow: 0 1px 5px rgba(0,0,0,0.08);
    height: 318px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
  }
  .data-table { width: 100%; border-collapse: collapse; font-size: 0.83rem; }
  .data-table th {
    color: #6b7280; font-size: 0.68rem; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.06em;
    padding: 0 10px 10px 0; text-align: left;
    border-bottom: 1px solid #e5e7eb;
  }
  .data-table td {
    color: #374151; padding: 8px 10px 8px 0;
    border-bottom: 1px solid #f3f4f6;
  }
  .data-table tr:last-child td { border-bottom: none; }
  .data-note { color: #9ca3af; font-size: 0.68rem; margin: 10px 0 0 0; }
</style>
""", unsafe_allow_html=True)

# ── Color palette ─────────────────────────────────────────────────────────────
COLORS = {
    "Flower":        "#2d6a4f",
    "Concentrates":  "#2563eb",
    "Carts":         "#7c3aed",
    "Edibles":       "#d97706",
    "Pre-rolls":     "#059669",
    "Tinctures":     "#0891b2",
    "Topical":       "#db2777",
}
ROOM_COLORS = {
    "F1-Trailer": "#2d6a4f",
    "F2-Trailer": "#2563eb",
    "F3-Trailer": "#d97706",
    "F4-Trailer": "#7c3aed",
}
CHART_BG   = "#ffffff"
GRID_COLOR = "#f3f4f6"
TEXT_COLOR = "#6b7280"

G_PER_LB = 453.592


def chart_layout(fig, title="", height=340):
    fig.update_layout(
        title=dict(text=title, font=dict(color="#111827", size=13), x=0.01, xanchor="left"),
        paper_bgcolor=CHART_BG, plot_bgcolor=CHART_BG,
        font=dict(color=TEXT_COLOR, family="sans-serif"),
        height=height, margin=dict(l=10, r=10, t=40 if title else 10, b=10),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#374151", size=11)),
        xaxis=dict(gridcolor=GRID_COLOR, zeroline=False, tickfont=dict(color=TEXT_COLOR), linecolor="#e5e7eb"),
        yaxis=dict(gridcolor=GRID_COLOR, zeroline=False, tickfont=dict(color=TEXT_COLOR), linecolor="#e5e7eb"),
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

    return inv_df, bio_df, sw_df, al_df


inv_df, bio_df, sw_df, al_df = load_data()


@st.cache_data(ttl=300)
def load_harvest():
    wb2 = openpyxl.load_workbook("harvest_tracker.xlsx", read_only=True, data_only=True)
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
    return pd.DataFrame(h_records)


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

# Harvest pre-compute — always produces DataFrames with correct columns
_harvest_error = None
_EMPTY_H = pd.DataFrame(columns=[
    "strain","harvest_date","status","cycle","room","num_plants",
    "fresh_frozen_g","wet_weight_g","dry_weight_g","sf_canopy","grams_per_sf",
    "trim_weight_g","trimmed_flower_g","est_35_units","est_1g_units","designation",
    "thc_pct","terps_pct","test_date","grade",
    "dry_lbs","wet_lbs","ff_lbs","trim_lbs","flower_lbs","month",
])
try:
    harvest_df = load_harvest()
    h = harvest_df.copy()
    h["harvest_date"] = pd.to_datetime(h["harvest_date"], errors="coerce")
    h["dry_lbs"]    = h["dry_weight_g"]  / G_PER_LB
    h["wet_lbs"]    = h["wet_weight_g"]  / G_PER_LB
    h["ff_lbs"]     = h["fresh_frozen_g"] / G_PER_LB
    h["trim_lbs"]   = h["trim_weight_g"] / G_PER_LB
    h["flower_lbs"] = h["trimmed_flower_g"] / G_PER_LB
    h["month"] = h["harvest_date"].dt.to_period("M").apply(
        lambda x: x.start_time if pd.notna(x) else pd.NaT
    )
    harvested = h[h["status"] == "Harvested"].copy()
    active    = h[h["status"].isin(["Flowering", "Moved to Trim Room"])].copy()
    upcoming  = h[h["status"] == "Upcoming Cycle"].copy()
except Exception as e:
    _harvest_error = str(e)
    harvested = active = upcoming = _EMPTY_H.copy()

# Pipeline bulk totals (needed in narrative banner)
pipeline_bio = bio_df[
    bio_df["stage"].notna() &
    ~bio_df["stage"].str.contains("Approved|Sample", na=False)
]
flower_lbs = pipeline_bio[pipeline_bio["unit"] == "lbs"]["display_count"].sum()
mfg_L      = pipeline_bio[pipeline_bio["unit"] == "L"]["display_count"].sum()

# ── Header ───────────────────────────────────────────────────────────────────
st.title("Eden Cannabis")

# Narrative banner
sw_cat_top = sw_df.groupby("category")["gross"].sum().idxmax() if not sw_df.empty else "Flower"
sw_cat_pct = (sw_df.groupby("category")["gross"].sum().max() / sw_df["gross"].sum() * 100) if not sw_df.empty else 0
st.markdown(f"""<div class="story-banner">
  <p class="story-head">YTD: ${ytd_gross:,.0f} gross revenue &nbsp;·&nbsp; {int(ytd_qty):,} units sold across all locations</p>
  <p class="story-sub">{sw_cat_top} leads Tampa sales at {sw_cat_pct:.0f}% of revenue &nbsp;·&nbsp; {total_finished:,} finished units on-hand &nbsp;·&nbsp; {flower_lbs:,.0f} lbs flower + {mfg_L:.1f} L concentrate in production pipeline</p>
</div>""", unsafe_allow_html=True)


view = st.radio("", ["Supply & Demand", "Harvest"], horizontal=True, label_visibility="collapsed")

if view == "Supply & Demand":

    # ── KPIs ──────────────────────────────────────────────────────────────────
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f"""<div class="kpi-hero">
          <div class="kpi-label-hero">YTD Gross Revenue</div>
          <div class="kpi-value-hero">${ytd_gross:,.0f}</div>
          <div class="kpi-sub-hero">Jan 1 – Jun 26, 2026</div>
        </div>""", unsafe_allow_html=True)
    for col, label, value, sub in [
        (k2, "YTD Net Revenue",  f"${ytd_net:,.0f}",         "After discounts"),
        (k3, "Units Sold YTD",   f"{int(ytd_qty):,}",        "All locations"),
        (k4, "On-Hand Units",    f"{total_finished:,}",       "Dispensaries + hub"),
        (k5, "Bulk Pipeline",    f"{flower_lbs:,.0f} lbs / {mfg_L:.1f} L", "Flower / concentrate"),
    ]:
        with col:
            st.markdown(f"""<div class="kpi-card">
              <div class="kpi-label">{label}</div>
              <div class="kpi-value">{value}</div>
              <div class="kpi-sub">{sub}</div>
            </div>""", unsafe_allow_html=True)

    st.divider()

    # ── Pipeline ──────────────────────────────────────────────────────────────
    st.subheader("Production Pipeline")

    pipe_df = bio_df[bio_df["stage"].notna()].copy()

    def _stage_bar(df, unit_label, title, height=320):
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
        fig.update_layout(barmode="stack", xaxis_tickangle=-30, yaxis_title=unit_label,
                          legend=dict(orientation="h", y=1.1, font=dict(size=10)))
        return chart_layout(fig, title, height=height)

    col_fl, col_mfg = st.columns(2)
    with col_fl:
        flower_pipe = pipe_df[pipe_df["unit"] == "lbs"]
        st.plotly_chart(
            _stage_bar(flower_pipe, "lbs", f"Flower  ·  {flower_pipe['display_count'].sum():,.0f} lbs total"),
            use_container_width=True)
    with col_mfg:
        mfg_pipe = pipe_df[pipe_df["unit"] == "L"]
        st.plotly_chart(
            _stage_bar(mfg_pipe, "L", f"Manufactured  ·  {mfg_pipe['display_count'].sum():,.2f} L total"),
            use_container_width=True)

    st.divider()

    # ── Inventory ─────────────────────────────────────────────────────────────
    st.subheader("Finished Goods Inventory")

    col_disp, col_vault = st.columns([3, 2])
    with col_disp:
        disp_agg = disp_df.copy()
        disp_agg["store"] = disp_agg["org"].map(DISP_ORGS)
        disp_agg = disp_agg.groupby(["store", "category"])["qty"].sum().reset_index()
        store_order = ["Tampa", "Sarasota", "Cocoa Beach", "Orlando"]
        fig3 = go.Figure()
        for cat in ["Flower","Concentrates","Edibles","Carts","Pre-rolls","Tinctures","Topical"]:
            d = disp_agg[disp_agg["category"] == cat]
            fig3.add_trace(go.Bar(
                name=cat, x=d["store"], y=d["qty"],
                marker_color=COLORS.get(cat, "#888"),
                hovertemplate="%{x} — " + cat + "<br>%{y:,} units<extra></extra>",
            ))
        fig3.update_layout(barmode="stack", yaxis_title="Units",
                           legend=dict(orientation="h", y=1.1),
                           xaxis=dict(categoryorder="array", categoryarray=store_order))
        chart_layout(fig3, "Dispensaries", height=320)
        st.plotly_chart(fig3, use_container_width=True)

    with col_vault:
        vault_agg   = vault_df.groupby("category")["qty"].sum().reset_index().sort_values("qty", ascending=False)
        vault_total = vault_agg["qty"].sum()
        fig4 = go.Figure(go.Pie(
            labels=vault_agg["category"], values=vault_agg["qty"], hole=0.55,
            marker=dict(colors=[COLORS.get(c, "#888") for c in vault_agg["category"]]),
            textinfo="label+percent", textfont=dict(color="#374151", size=11),
            hovertemplate="%{label}<br>%{value:,} units<extra></extra>",
        ))
        fig4.add_annotation(text=f"{vault_total:,}<br>units", x=0.5, y=0.5,
                            showarrow=False, font=dict(color="#111827", size=13))
        chart_layout(fig4, f"Processing Vault  ·  {vault_total:,} units", height=320)
        st.plotly_chart(fig4, use_container_width=True)

    st.divider()

    # ── Sales ─────────────────────────────────────────────────────────────────
    st.subheader("YTD Sales")

    col_cat, col_loc = st.columns(2)
    with col_cat:
        sw_cat = sw_df.groupby("category").agg(gross=("gross","sum")).reset_index()
        al_cat = al_df.groupby("cat_norm").agg(gross=("gross","sum")).reset_index()
        al_cat.rename(columns={"cat_norm": "category"}, inplace=True)
        cat_total = pd.concat([sw_cat, al_cat]).groupby("category").agg(
            gross=("gross","sum")).reset_index().sort_values("gross", ascending=True)
        fig5 = go.Figure(go.Bar(
            x=cat_total["gross"], y=cat_total["category"], orientation="h",
            marker_color=[COLORS.get(c, "#2e7bc4") for c in cat_total["category"]],
            text=cat_total["gross"].apply(lambda x: f"${x:,.0f}"),
            textposition="outside", textfont=dict(color=TEXT_COLOR, size=11),
            hovertemplate="%{y}  $%{x:,.0f}<extra></extra>",
        ))
        chart_layout(fig5, "Gross Sales by Category", height=300)
        fig5.update_layout(xaxis_title="", yaxis=dict(gridcolor="rgba(0,0,0,0)"))
        st.plotly_chart(fig5, use_container_width=True)

    with col_loc:
        al_loc = al_df.groupby("loc_norm").agg(gross=("gross","sum")).reset_index()
        sw_loc = sw_df.groupby("store").agg(gross=("gross","sum")).reset_index()
        sw_loc.rename(columns={"store": "loc_norm"}, inplace=True)
        al_loc["period"] = "Jan–Apr"
        sw_loc["period"] = "Mar–Jun (Tampa)"
        loc_all = pd.concat([al_loc, sw_loc])
        fig6 = go.Figure()
        for period, color in [("Jan–Apr", "#2563eb"), ("Mar–Jun (Tampa)", "#2d6a4f")]:
            d = loc_all[loc_all["period"] == period]
            fig6.add_trace(go.Bar(name=period, x=d["loc_norm"], y=d["gross"],
                                  marker_color=color,
                                  hovertemplate="%{x}  $%{y:,.0f}<extra>" + period + "</extra>"))
        fig6.update_layout(barmode="group", yaxis_title="",
                           legend=dict(orientation="h", y=1.1))
        chart_layout(fig6, "Gross Sales by Location", height=300)
        st.plotly_chart(fig6, use_container_width=True)

    col_trend, col_pie = st.columns([3, 2])
    with col_trend:
        weekly = sw_df.groupby(["week","category"]).agg(gross=("gross","sum")).reset_index()
        fig7 = go.Figure()
        for cat in ["Flower","Concentrates","Edibles","Carts","Pre-rolls","Tinctures","Topical"]:
            d = weekly[weekly["category"] == cat].sort_values("week")
            fig7.add_trace(go.Scatter(
                x=d["week"], y=d["gross"], name=cat, mode="lines",
                line=dict(color=COLORS.get(cat,"#888"), width=2), stackgroup="one",
                hovertemplate="%{x|%b %d}  $%{y:,.0f}<extra>" + cat + "</extra>",
            ))
        fig7.update_layout(yaxis_title="", legend=dict(orientation="h", y=1.1))
        chart_layout(fig7, "Weekly Sales — Tampa (Sweed)", height=300)
        st.plotly_chart(fig7, use_container_width=True)

    with col_pie:
        sw_ct = sw_df.groupby("category")["gross"].sum().reset_index().sort_values("gross", ascending=False)
        fig8 = go.Figure(go.Pie(
            labels=sw_ct["category"], values=sw_ct["gross"], hole=0.55,
            marker=dict(colors=[COLORS.get(c,"#888") for c in sw_ct["category"]]),
            textinfo="label+percent", textfont=dict(color="#374151", size=11),
            hovertemplate="%{label}  $%{value:,.0f}<extra></extra>",
        ))
        fig8.add_annotation(text=f"${sw_total_gross:,.0f}", x=0.5, y=0.5,
                            showarrow=False, font=dict(color="#111827", size=13))
        chart_layout(fig8, "Tampa Sales Mix", height=300)
        st.plotly_chart(fig8, use_container_width=True)

    st.divider()

    # ── Supply vs Demand ──────────────────────────────────────────────────────
    st.subheader("Weeks of Stock")

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
            name="Units On-Hand", x=svd["category"], y=svd["on_hand"],
            marker_color=[COLORS.get(c, "#888") for c in svd["category"]], opacity=0.8,
            hovertemplate="%{x}  %{y:,} units<extra></extra>",
        ), secondary_y=False)
        fig9.add_trace(go.Scatter(
            name="Weekly Velocity (Tampa)", x=svd["category"], y=svd["weekly_velocity"],
            mode="markers+lines", marker=dict(size=9, color="#d97706", symbol="diamond"),
            line=dict(color="#d97706", width=2, dash="dot"),
            hovertemplate="%{x}  ~%{y:.0f}/wk<extra></extra>",
        ), secondary_y=True)
        fig9.update_yaxes(title_text="Units", secondary_y=False,
                          gridcolor=GRID_COLOR, tickfont=dict(color=TEXT_COLOR))
        fig9.update_yaxes(title_text="Units/Week", secondary_y=True,
                          gridcolor="rgba(0,0,0,0)", tickfont=dict(color="#d97706"))
        fig9.update_layout(
            paper_bgcolor=CHART_BG, plot_bgcolor=CHART_BG, font=dict(color=TEXT_COLOR),
            height=300, margin=dict(l=10, r=10, t=40, b=10),
            legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#374151", size=11)),
            title=dict(text="Dispensary Stock vs. Weekly Sales Velocity",
                       font=dict(color="#111827", size=13), x=0.01),
            xaxis=dict(gridcolor=GRID_COLOR, tickfont=dict(color=TEXT_COLOR)),
        )
        st.plotly_chart(fig9, use_container_width=True)

    with col_svd_tbl:
        svd_d = svd[["category","on_hand","weekly_velocity","weeks_of_stock"]].copy()
        svd_d["on_hand_fmt"]   = svd_d["on_hand"].apply(lambda x: f"{int(x):,}")
        svd_d["velocity_fmt"]  = svd_d["weekly_velocity"].apply(lambda x: f"~{x:.0f}/wk" if pd.notna(x) and x > 0 else "–")
        svd_d["weeks_fmt"]     = svd_d["weeks_of_stock"].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "–")
        rows_html = "".join(
            f"<tr><td>{r['category']}</td><td>{r['on_hand_fmt']}</td>"
            f"<td>{r['velocity_fmt']}</td><td>{r['weeks_fmt']}</td></tr>"
            for _, r in svd_d.sort_values("category").iterrows()
        )
        st.markdown(f"""<div class="data-card">
          <table class="data-table">
            <thead><tr><th>Category</th><th>On-Hand</th><th>Wkly Vel.</th><th>Wks Stock</th></tr></thead>
            <tbody>{rows_html}</tbody>
          </table>
          <p class="data-note">Tampa velocity · last 4 wks &nbsp;·&nbsp; +{vault_df['qty'].sum():,} units in processing vault</p>
        </div>""", unsafe_allow_html=True)


if view == "Harvest":

    # ── KPIs ──────────────────────────────────────────────────────────────────
    if _harvest_error:
        st.error(f"Could not load harvest data: {_harvest_error}")

    total_cycles  = len(harvested)
    total_dry_lbs = harvested["dry_lbs"].sum()
    total_ff_lbs  = harvested["ff_lbs"].sum()
    total_wet_lbs = harvested["wet_lbs"].sum()
    avg_dry_ratio = (total_dry_lbs / total_wet_lbs * 100) if total_wet_lbs > 0 else 0
    avg_per_plant = (harvested["dry_weight_g"] / harvested["num_plants"]).dropna().mean() / G_PER_LB

    h1, h2, h3, h4, h5 = st.columns(5)
    with h1:
        st.markdown(f"""<div class="kpi-hero">
          <div class="kpi-label-hero">Cycles Harvested</div>
          <div class="kpi-value-hero">{total_cycles}</div>
          <div class="kpi-sub-hero">Nov 2024 – Jun 2026</div>
        </div>""", unsafe_allow_html=True)
    for col, label, value, sub in [
        (h2, "Total Dry Yield",    f"{total_dry_lbs:,.0f} lbs", "All cycles"),
        (h3, "Total Fresh Frozen", f"{total_ff_lbs:,.0f} lbs",  "Extract feedstock"),
        (h4, "Wet → Dry Ratio",    f"{avg_dry_ratio:.1f}%",     "Avg across all harvests"),
        (h5, "Yield / Plant",      f"{avg_per_plant:.2f} lbs",  "Avg dry weight"),
    ]:
        with col:
            st.markdown(f"""<div class="kpi-card">
              <div class="kpi-label">{label}</div>
              <div class="kpi-value">{value}</div>
              <div class="kpi-sub">{sub}</div>
            </div>""", unsafe_allow_html=True)

    st.divider()

    # ── Monthly output ────────────────────────────────────────────────────────
    st.subheader("Monthly Harvest Output")

    monthly_h = harvested.dropna(subset=["month"]).groupby("month").agg(
        dry_lbs=("dry_lbs","sum"), ff_lbs=("ff_lbs","sum"),
        trim_lbs=("trim_lbs","sum"), cycles=("strain","count"),
    ).reset_index().sort_values("month")

    col_m1, col_m2 = st.columns([3, 2])
    with col_m1:
        fig_mo = go.Figure()
        for name, col_data, color in [
            ("Dry Flower", monthly_h["dry_lbs"],  "#2d6a4f"),
            ("Fresh Frozen", monthly_h["ff_lbs"], "#2563eb"),
            ("Trim",        monthly_h["trim_lbs"],"#d97706"),
        ]:
            fig_mo.add_trace(go.Bar(name=name, x=monthly_h["month"], y=col_data,
                                    marker_color=color,
                                    hovertemplate="%{x|%b %Y}  %{y:,.0f} lbs<extra>" + name + "</extra>"))
        fig_mo.add_trace(go.Scatter(
            name="Cycles", x=monthly_h["month"], y=monthly_h["cycles"],
            mode="markers+lines", yaxis="y2",
            marker=dict(size=7, color="#d97706"), line=dict(color="#d97706", width=2, dash="dot"),
            hovertemplate="%{x|%b %Y}  %{y} cycles<extra></extra>",
        ))
        fig_mo.update_layout(
            barmode="stack", yaxis_title="lbs",
            yaxis2=dict(title="Cycles", overlaying="y", side="right",
                        gridcolor="rgba(0,0,0,0)", tickfont=dict(color="#d97706")),
            legend=dict(orientation="h", y=1.1),
        )
        chart_layout(fig_mo, "Dry Flower, Fresh Frozen & Trim (lbs)", height=340)
        st.plotly_chart(fig_mo, use_container_width=True)

    with col_m2:
        des_agg = harvested.groupby("designation").agg(dry_lbs=("dry_lbs","sum")).reset_index()
        des_agg = des_agg[des_agg["designation"].notna() & (des_agg["designation"] != "")]
        des_colors = {"Flower":"#2d6a4f","Wash":"#2563eb","Pre-rolls":"#d97706","TBD":"#9ca3af"}
        fig_des = go.Figure(go.Pie(
            labels=des_agg["designation"], values=des_agg["dry_lbs"], hole=0.55,
            marker=dict(colors=[des_colors.get(d,"#888") for d in des_agg["designation"]]),
            textinfo="label+percent", textfont=dict(color=TEXT_COLOR, size=11),
            hovertemplate="%{label}  %{value:,.0f} lbs<extra></extra>",
        ))
        fig_des.add_annotation(text=f"{des_agg['dry_lbs'].sum():,.0f}<br>lbs",
                               x=0.5, y=0.5, showarrow=False, font=dict(color="#111827", size=12))
        chart_layout(fig_des, "Designation Mix", height=340)
        st.plotly_chart(fig_des, use_container_width=True)

    st.divider()

    # ── Room performance ──────────────────────────────────────────────────────
    st.subheader("Room Performance")

    room_agg = harvested.dropna(subset=["room"]).groupby("room").agg(
        dry_lbs=("dry_lbs","sum"), ff_lbs=("ff_lbs","sum"),
        cycles=("strain","count"), plants=("num_plants","sum"),
    ).reset_index()
    room_agg["dry_per_plant"] = (room_agg["dry_lbs"] / room_agg["plants"]).round(3)

    col_r1, col_r2 = st.columns(2)
    with col_r1:
        fig_room = go.Figure()
        fig_room.add_trace(go.Bar(name="Dry Flower", x=room_agg["room"], y=room_agg["dry_lbs"],
            marker_color=[ROOM_COLORS.get(r,"#888") for r in room_agg["room"]],
            hovertemplate="%{x}  %{y:,.0f} lbs<extra></extra>"))
        fig_room.add_trace(go.Bar(name="Fresh Frozen", x=room_agg["room"], y=room_agg["ff_lbs"],
            marker_color=[ROOM_COLORS.get(r,"#888") for r in room_agg["room"]],
            opacity=0.5,
            hovertemplate="%{x}  %{y:,.0f} lbs FF<extra></extra>"))
        fig_room.update_layout(barmode="group", yaxis_title="lbs", legend=dict(orientation="h", y=1.1))
        chart_layout(fig_room, "Output by Room", height=300)
        st.plotly_chart(fig_room, use_container_width=True)

    with col_r2:
        fig_eff = go.Figure(go.Bar(
            x=room_agg["room"], y=room_agg["dry_per_plant"],
            marker_color=[ROOM_COLORS.get(r,"#888") for r in room_agg["room"]],
            text=room_agg["dry_per_plant"].apply(lambda x: f"{x:.3f}"),
            textposition="outside", textfont=dict(color=TEXT_COLOR),
            hovertemplate="%{x}  %{y:.3f} lbs/plant<extra></extra>",
        ))
        chart_layout(fig_eff, "Avg Dry Yield per Plant (lbs)", height=300)
        fig_eff.update_layout(yaxis_title="lbs / plant")
        st.plotly_chart(fig_eff, use_container_width=True)

    st.divider()

    # ── Strains ───────────────────────────────────────────────────────────────
    st.subheader("Strain Performance")

    col_s1, col_s2 = st.columns([3, 2])
    with col_s1:
        strain_agg = harvested.groupby("strain").agg(
            dry_lbs=("dry_lbs","sum"), wet_lbs=("wet_lbs","sum"),
            cycles=("cycle","count"), plants=("num_plants","sum"),
        ).reset_index()
        strain_agg["dry_ratio"] = (strain_agg["dry_lbs"] / strain_agg["wet_lbs"] * 100).round(1)
        strain_top = strain_agg.sort_values("dry_lbs", ascending=True).tail(20)
        fig_st = go.Figure(go.Bar(
            name="Dry lbs", y=strain_top["strain"], x=strain_top["dry_lbs"],
            orientation="h", marker_color="#2d6a4f",
            hovertemplate="%{y}  %{x:,.0f} lbs (%{customdata:.1f}% of wet)<extra></extra>",
            customdata=strain_top["dry_ratio"],
        ))
        chart_layout(fig_st, "Top 20 Strains — Dry Yield (lbs)", height=460)
        fig_st.update_layout(xaxis_title="", yaxis=dict(gridcolor="rgba(0,0,0,0)"))
        st.plotly_chart(fig_st, use_container_width=True)

    with col_s2:
        thc_data = harvested.dropna(subset=["thc_pct"]).sort_values("thc_pct", ascending=True)
        if not thc_data.empty:
            fig_thc = go.Figure(go.Bar(
                y=thc_data["strain"], x=thc_data["thc_pct"], orientation="h",
                marker_color="#7c3aed",
                text=thc_data.apply(
                    lambda r: f"{r['thc_pct']:.1f}%  {r['terps_pct']:.2f}% terps"
                              if r["terps_pct"] else f"{r['thc_pct']:.1f}%", axis=1),
                textposition="outside", textfont=dict(color=TEXT_COLOR, size=11),
                hovertemplate="%{y}  THC %{x:.1f}%<extra></extra>",
            ))
            chart_layout(fig_thc, "COA Results — THC %", height=240)
            fig_thc.update_layout(xaxis=dict(range=[0, max(thc_data["thc_pct"]) * 1.4]),
                                  yaxis=dict(gridcolor="rgba(0,0,0,0)"))
            st.plotly_chart(fig_thc, use_container_width=True)

        scatter_data = harvested.dropna(subset=["wet_lbs","dry_lbs"])
        scatter_data = scatter_data[scatter_data["wet_lbs"] > 0]
        fig_scat = go.Figure()
        for room in scatter_data["room"].dropna().unique():
            d = scatter_data[scatter_data["room"] == room]
            fig_scat.add_trace(go.Scatter(
                x=d["wet_lbs"], y=d["dry_lbs"], name=room, mode="markers",
                marker=dict(size=7, color=ROOM_COLORS.get(room,"#888"), opacity=0.75),
                hovertemplate="%{text}<br>Wet %{x:,.0f} · Dry %{y:,.0f} lbs<extra></extra>",
                text=d["strain"],
            ))
        chart_layout(fig_scat, "Wet vs. Dry per Cycle", height=240)
        fig_scat.update_layout(xaxis_title="Wet lbs", yaxis_title="Dry lbs",
                               legend=dict(orientation="h", y=1.12))
        st.plotly_chart(fig_scat, use_container_width=True)

    st.divider()

    # ── Active & upcoming ─────────────────────────────────────────────────────
    st.subheader("Active & Upcoming Cycles")

    def _html_table(df, note=None):
        cols = list(df.columns)
        header = "".join(f"<th>{c}</th>" for c in cols)
        body = "".join(
            "<tr>" + "".join(f"<td>{row[c] if pd.notna(row[c]) else '–'}</td>" for c in cols) + "</tr>"
            for _, row in df.iterrows()
        )
        note_html = f'<p class="data-note">{note}</p>' if note else ""
        return f"""<div class="data-card" style="height:auto;max-height:360px;overflow-y:auto;">
          <table class="data-table">
            <thead><tr>{header}</tr></thead>
            <tbody>{body}</tbody>
          </table>{note_html}
        </div>"""

    col_a1, col_a2 = st.columns(2)
    with col_a1:
        active_display = active[["strain","status","harvest_date","room","cycle"]].copy()
        active_display["harvest_date"] = active_display["harvest_date"].dt.strftime("%b %d, %Y").fillna("–")
        active_display.columns = ["Strain","Status","Est. Harvest","Room","Cycle"]
        st.markdown(_html_table(active_display.sort_values("Est. Harvest").fillna("–"),
                                note="Currently flowering / in trim"), unsafe_allow_html=True)

    with col_a2:
        upcoming_display = upcoming[["strain","harvest_date"]].copy()
        upcoming_display["harvest_date"] = upcoming_display["harvest_date"].dt.strftime("%b %d, %Y").fillna("–")
        upcoming_display.columns = ["Strain","Planned Harvest"]
        st.markdown(_html_table(upcoming_display.sort_values("Planned Harvest").fillna("–"),
                                note="Upcoming planned cycles"), unsafe_allow_html=True)

    st.divider()
    st.subheader("Full Harvest Log")
    log = harvested[["strain","harvest_date","room","cycle","designation",
                      "dry_lbs","ff_lbs","trim_lbs","thc_pct","terps_pct","num_plants"]].copy()
    log["harvest_date"] = log["harvest_date"].dt.strftime("%b %d, %Y")
    log["dry_lbs"]   = log["dry_lbs"].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "–")
    log["ff_lbs"]    = log["ff_lbs"].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "–")
    log["trim_lbs"]  = log["trim_lbs"].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "–")
    log["thc_pct"]   = log["thc_pct"].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "–")
    log["terps_pct"] = log["terps_pct"].apply(lambda x: f"{x:.3f}%" if pd.notna(x) else "–")
    log["num_plants"]= log["num_plants"].apply(lambda x: f"{int(x)}" if pd.notna(x) else "–")
    log.columns = ["Strain","Date","Room","Cycle","Designation",
                   "Dry (lbs)","FF (lbs)","Trim (lbs)","THC%","Terps%","Plants"]
    st.markdown(
        f"""<div class="data-card" style="height:auto;max-height:500px;overflow-y:auto;overflow-x:auto;">
          <table class="data-table">
            <thead><tr>{"".join(f"<th>{c}</th>" for c in log.columns)}</tr></thead>
            <tbody>{"".join(
              "<tr>" + "".join(f"<td>{row[c]}</td>" for c in log.columns) + "</tr>"
              for _, row in log.sort_values("Date", ascending=False).iterrows()
            )}</tbody>
          </table>
        </div>""", unsafe_allow_html=True)


st.caption("Eden Cannabis · Confidential · June 2026")
