"""
PPC Campaign Dashboard
======================
Upload your Amazon PPC bulk Excel file and explore campaign-level
performance across all 4 ad type tabs with decoded campaign name filters.

HOW TO RUN:
    pip install streamlit pandas openpyxl
    streamlit run ppc_dashboard.py
"""

import streamlit as st
import pandas as pd
import io

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="PPC Dashboard", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .metric-card {
        background: #f8f9fa; border-radius: 8px;
        padding: 16px 20px; border-left: 4px solid #FF9900;
        margin-bottom: 8px;
    }
    .metric-label { font-size: 11px; color: #888; font-weight: 700;
                    text-transform: uppercase; letter-spacing: 0.5px; }
    .metric-value { font-size: 22px; font-weight: 800; color: #1a1a1a; margin-top: 4px; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
CAMPAIGN_NAME_COL = "Campaign Name (Informational only)"
PORTFOLIO_COL     = "Portfolio Name (Informational only)"

DECODED_FIELDS = [
    "BRAND", "SKU", "TYPE", "TARGET_MODE", "MATCH_TYPE",
    "KW_BUCKET", "GOAL", "KW_ROOT", "BID_STRAT",
    "ASIN_CHILD", "ASIN_PARENT", "SB_FORMAT", "VIDEO_TYPE", "THEME",
]

NULL_VALUES = {"_", "-", "", "nan", "None"}

PERF_COLS = ["Impressions", "Clicks", "Spend", "Sales", "Orders", "Units"]

TAB_SHEETS = [
    "Sponsored Products Campaigns",
    "Sponsored Brands Campaigns",
    "SB Multi Ad Group Campaigns",
    "Sponsored Display Campaigns",
]

TAB_LABELS = {
    "Sponsored Products Campaigns": "🛒 Sponsored Products",
    "Sponsored Brands Campaigns":   "🏷️ Sponsored Brands",
    "SB Multi Ad Group Campaigns":  "🗂️ SB Multi Ad Group",
    "Sponsored Display Campaigns":  "🖥️ Sponsored Display",
}

FILTER_FIELDS = [
    ("BRAND",       "Brand"),
    ("SKU",         "SKU"),
    ("TYPE",        "Ad Type"),
    ("TARGET_MODE", "Target Mode"),
    ("MATCH_TYPE",  "Match Type"),
    ("KW_BUCKET",   "KW Bucket"),
    ("GOAL",        "Goal"),
    ("KW_ROOT",     "KW Root"),
    ("BID_STRAT",   "Bid Strategy"),
    ("ASIN_CHILD",  "ASIN Child"),
    ("ASIN_PARENT", "ASIN Parent"),
    ("SB_FORMAT",   "SB Format"),
    ("VIDEO_TYPE",  "Video Type"),
    ("THEME",       "Theme"),
]


# ── Helpers ───────────────────────────────────────────────────────────────────

def decode_campaign_name(name: str) -> dict:
    if not isinstance(name, str) or name.strip() in NULL_VALUES:
        return {f: None for f in DECODED_FIELDS}
    parts = [p.strip() for p in name.split("~")]
    parts += [""] * (14 - len(parts))
    parts = parts[:14]
    return {
        field: (None if value in NULL_VALUES else value)
        for field, value in zip(DECODED_FIELDS, parts)
    }


def safe_float(val):
    try:
        return float(str(val).replace(",", "").replace("%", "").strip())
    except Exception:
        return 0.0


@st.cache_data(show_spinner=False)
def load_and_process(file_bytes: bytes) -> dict:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    result = {}

    for sheet in TAB_SHEETS:
        if sheet not in xls.sheet_names:
            result[sheet] = None
            continue

        df = xls.parse(sheet)
        df.columns = [str(c).strip() for c in df.columns]

        if CAMPAIGN_NAME_COL not in df.columns:
            result[sheet] = None
            continue

        # Keep only rows with a valid campaign name
        df = df[df[CAMPAIGN_NAME_COL].notna()].copy()
        df[CAMPAIGN_NAME_COL] = df[CAMPAIGN_NAME_COL].astype(str).str.strip()
        df = df[~df[CAMPAIGN_NAME_COL].isin(NULL_VALUES)]

        # Ensure performance columns exist as numeric
        for col in PERF_COLS:
            if col in df.columns:
                df[col] = df[col].apply(safe_float)
            else:
                df[col] = 0.0

        if PORTFOLIO_COL not in df.columns:
            df[PORTFOLIO_COL] = None

        # Aggregate to campaign level (sum metrics, first portfolio)
        agg_dict = {col: "sum" for col in PERF_COLS}
        agg_dict[PORTFOLIO_COL] = "first"

        campaign_df = df.groupby(CAMPAIGN_NAME_COL, as_index=False).agg(agg_dict)

        # Compute KPIs
        spend  = campaign_df["Spend"]
        sales  = campaign_df["Sales"]
        clicks = campaign_df["Clicks"]
        orders = campaign_df["Orders"]

        campaign_df["Conv. Rate"] = (orders / clicks.replace(0, float("nan"))).fillna(0)
        campaign_df["ACOS"]       = (spend  / sales.replace(0, float("nan"))).fillna(0)
        campaign_df["CPC"]        = (spend  / clicks.replace(0, float("nan"))).fillna(0)
        campaign_df["ROAS"]       = (sales  / spend.replace(0, float("nan"))).fillna(0)

        # Decode campaign name
        decoded = campaign_df[CAMPAIGN_NAME_COL].apply(decode_campaign_name).apply(pd.Series)
        campaign_df = pd.concat([campaign_df.reset_index(drop=True),
                                  decoded.reset_index(drop=True)], axis=1)

        result[sheet] = campaign_df

    return result


def render_metrics(df):
    impr   = int(df["Impressions"].sum())
    clicks = int(df["Clicks"].sum())
    spend  = df["Spend"].sum()
    sales  = df["Sales"].sum()
    orders = int(df["Orders"].sum())
    acos   = spend / sales  if sales  > 0 else 0
    roas   = sales / spend  if spend  > 0 else 0
    ctr    = clicks / impr  if impr   > 0 else 0
    cpc    = spend / clicks if clicks > 0 else 0
    cvr    = orders / clicks if clicks > 0 else 0

    row1 = st.columns(5)
    for col, (lbl, val) in zip(row1, [
        ("💰 Spend",   f"${spend:,.2f}"),
        ("📈 Sales",   f"${sales:,.2f}"),
        ("📦 Orders",  f"{orders:,}"),
        ("🎯 ACOS",    f"{acos*100:.2f}%"),
        ("🔄 ROAS",    f"{roas:.2f}"),
    ]):
        col.markdown(f'<div class="metric-card"><div class="metric-label">{lbl}</div>'
                     f'<div class="metric-value">{val}</div></div>', unsafe_allow_html=True)

    st.markdown("")
    row2 = st.columns(5)
    for col, (lbl, val) in zip(row2, [
        ("👁️ Impressions", f"{impr:,}"),
        ("🖱️ Clicks",      f"{clicks:,}"),
        ("📊 CTR",         f"{ctr*100:.3f}%"),
        ("💵 CPC",         f"${cpc:.2f}"),
        ("✅ Conv. Rate",  f"{cvr*100:.2f}%"),
    ]):
        col.markdown(f'<div class="metric-card"><div class="metric-label">{lbl}</div>'
                     f'<div class="metric-value">{val}</div></div>', unsafe_allow_html=True)

    st.markdown("")


def render_table(df, sheet_name):
    display_cols = (
        [CAMPAIGN_NAME_COL, PORTFOLIO_COL]
        + DECODED_FIELDS
        + ["Impressions", "Clicks", "Spend", "Sales", "Orders", "Units",
           "Conv. Rate", "ACOS", "CPC", "ROAS"]
    )
    display_cols = [c for c in display_cols if c in df.columns]
    show = df[display_cols].copy()

    # Format
    for c in ["Impressions", "Clicks", "Orders", "Units"]:
        if c in show.columns:
            show[c] = show[c].apply(lambda x: f"{int(x):,}")
    for c in ["Spend", "Sales", "CPC"]:
        if c in show.columns:
            show[c] = show[c].apply(lambda x: f"${x:,.2f}")
    for c in ["Conv. Rate", "ACOS"]:
        if c in show.columns:
            show[c] = show[c].apply(lambda x: f"{x*100:.2f}%")
    if "ROAS" in show.columns:
        show["ROAS"] = show["ROAS"].apply(lambda x: f"{x:.2f}")

    show = show.rename(columns={CAMPAIGN_NAME_COL: "Campaign Name", PORTFOLIO_COL: "Portfolio"})

    st.dataframe(show, use_container_width=True, height=480)
    st.caption(f"**{len(show):,}** campaigns shown")

    csv = df[display_cols].to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download filtered CSV",
        csv,
        file_name=f"{sheet_name.replace(' ', '_')}_filtered.csv",
        mime="text/csv",
        key=f"dl_{sheet_name}",
    )


# ── App ───────────────────────────────────────────────────────────────────────

st.title("📊 PPC Campaign Dashboard")

uploaded = st.file_uploader(
    "Drop your Amazon PPC bulk Excel file (.xlsx)",
    type=["xlsx"],
    help="The file should contain tabs: Sponsored Products Campaigns, Sponsored Brands Campaigns, SB Multi Ad Group Campaigns, Sponsored Display Campaigns"
)

if not uploaded:
    st.info("👆 Upload your weekly PPC bulk file to get started. All processing happens locally — nothing is stored.")
    st.stop()

with st.spinner("Loading data..."):
    data = load_and_process(uploaded.read())

available = {k: v for k, v in data.items() if v is not None and len(v) > 0}

if not available:
    st.error("No valid PPC campaign data found in the uploaded file.")
    st.stop()

st.success(f"✅ Loaded {sum(len(v) for v in available.values()):,} campaigns across {len(available)} tabs")

tab_objects = st.tabs([TAB_LABELS.get(s, s) for s in TAB_SHEETS if s in available])

for tab_obj, sheet_name in zip(tab_objects, [s for s in TAB_SHEETS if s in available]):
    with tab_obj:
        df = available[sheet_name].copy()
        total = len(df)

        with st.expander(f"🔍 Filters  —  {total:,} campaigns total", expanded=False):
            fcols = st.columns(4)
            col_idx = 0

            # Portfolio filter first
            if PORTFOLIO_COL in df.columns:
                port_opts = sorted([
                    v for v in df[PORTFOLIO_COL].dropna().unique()
                    if str(v).strip() not in NULL_VALUES
                ])
                if port_opts:
                    sel = fcols[col_idx % 4].multiselect(
                        "Portfolio", port_opts, key=f"{sheet_name}_portfolio"
                    )
                    if sel:
                        df = df[df[PORTFOLIO_COL].isin(sel)]
                    col_idx += 1

            # Decoded field filters
            for field, label in FILTER_FIELDS:
                if field not in df.columns:
                    continue
                opts = sorted([v for v in df[field].dropna().unique() if v is not None])
                if not opts:
                    continue
                sel = fcols[col_idx % 4].multiselect(label, opts, key=f"{sheet_name}_{field}")
                if sel:
                    df = df[df[field].isin(sel)]
                col_idx += 1

        render_metrics(df)
        render_table(df, sheet_name)
