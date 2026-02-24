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
    .filter-header { font-size: 13px; font-weight: 700; color: #444;
                     margin-bottom: 8px; margin-top: 4px; }
    .section-header { font-size: 15px; font-weight: 700; color: #333;
                      margin-top: 24px; margin-bottom: 4px; }

    /* Pivot subtotal rows styling */
    .pivot-total   { background: #fff3e0 !important; font-weight: 700; }
    .pivot-grand   { background: #FF9900 !important; color: white !important; font-weight: 800; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
CAMPAIGN_NAME_COL  = "Campaign Name (Informational only)"
PORTFOLIO_COL      = "Portfolio Name (Informational only)"
CAMPAIGN_STATE_COL = "Campaign State (Informational only)"
SERVING_STATUS_COL = "Campaign Serving Status (Informational only)"

DECODED_FIELDS = [
    "BRAND", "SKU", "TYPE", "TARGET_MODE", "MATCH_TYPE",
    "KW_BUCKET", "GOAL", "KW_ROOT", "BID_STRAT",
    "ASIN_CHILD", "ASIN_PARENT", "SB_FORMAT", "VIDEO_TYPE", "THEME",
]

NULL_VALUES = {"_", "-", "", "nan", "None"}

PERF_COLS = ["Impressions", "Clicks", "Spend", "Sales", "Orders", "Units"]

PIVOT_DIMS = ["SKU", "TYPE", "TARGET_MODE", "MATCH_TYPE", "KW_BUCKET", "GOAL", "KW_ROOT"]
PIVOT_METRICS = ["Clicks", "Spend", "Sales"]

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

PLACEHOLDER = "—"


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

        df = df[df[CAMPAIGN_NAME_COL].notna()].copy()
        df[CAMPAIGN_NAME_COL] = df[CAMPAIGN_NAME_COL].astype(str).str.strip()
        df = df[~df[CAMPAIGN_NAME_COL].isin(NULL_VALUES)]

        for col in PERF_COLS:
            if col in df.columns:
                df[col] = df[col].apply(safe_float)
            else:
                df[col] = 0.0

        for col in [PORTFOLIO_COL, CAMPAIGN_STATE_COL, SERVING_STATUS_COL]:
            if col not in df.columns:
                df[col] = None

        agg_dict = {col: "sum" for col in PERF_COLS}
        agg_dict[PORTFOLIO_COL]      = "first"
        agg_dict[CAMPAIGN_STATE_COL] = "first"
        agg_dict[SERVING_STATUS_COL] = "first"

        campaign_df = df.groupby(CAMPAIGN_NAME_COL, as_index=False).agg(agg_dict)

        spend  = campaign_df["Spend"]
        sales  = campaign_df["Sales"]
        clicks = campaign_df["Clicks"]
        orders = campaign_df["Orders"]

        campaign_df["Conv. Rate"] = (orders / clicks.replace(0, float("nan"))).fillna(0)
        campaign_df["ACOS"]       = (spend  / sales.replace(0, float("nan"))).fillna(0)
        campaign_df["CPC"]        = (spend  / clicks.replace(0, float("nan"))).fillna(0)
        campaign_df["ROAS"]       = (sales  / spend.replace(0, float("nan"))).fillna(0)

        decoded = campaign_df[CAMPAIGN_NAME_COL].apply(decode_campaign_name).apply(pd.Series)
        campaign_df = pd.concat([campaign_df.reset_index(drop=True),
                                  decoded.reset_index(drop=True)], axis=1)

        result[sheet] = campaign_df

    return result


def build_pivot(df: pd.DataFrame) -> pd.DataFrame:
    """
    Build a hierarchical pivot with:
    - Detail rows: one per unique combination of all 7 dims
    - Subtotals per SKU
    - Subtotals per SKU + TYPE
    - Grand total row at the bottom
    Returned as a flat DataFrame with a _row_type column for styling.
    """
    dims = [d for d in PIVOT_DIMS if d in df.columns]
    metrics = [m for m in PIVOT_METRICS if m in df.columns]

    # Fill nulls with placeholder so groupby works
    work = df.copy()
    for d in dims:
        work[d] = work[d].fillna(PLACEHOLDER)

    rows = []

    # Group by SKU
    for sku, sku_df in work.groupby("SKU", sort=True):

        # Group by SKU + TYPE
        for type_val, type_df in sku_df.groupby("TYPE", sort=True):

            # Detail rows for this SKU+TYPE combination
            detail_dims = [d for d in dims if d not in ("SKU", "TYPE")]
            detail_group = type_df.groupby(detail_dims, sort=True)[metrics].sum().reset_index()

            for _, row in detail_group.iterrows():
                r = {d: PLACEHOLDER for d in dims}
                r["SKU"]  = sku
                r["TYPE"] = type_val
                for d in detail_dims:
                    r[d] = row[d]
                for m in metrics:
                    r[m] = row[m]
                r["_row_type"] = "detail"
                rows.append(r)

            # Subtotal: SKU + TYPE
            type_totals = type_df[metrics].sum()
            r = {d: PLACEHOLDER for d in dims}
            r["SKU"]       = sku
            r["TYPE"]      = f"∑ {type_val}"
            r["_row_type"] = "type_subtotal"
            for m in metrics:
                r[m] = type_totals[m]
            rows.append(r)

        # Subtotal: SKU
        sku_totals = sku_df[metrics].sum()
        r = {d: PLACEHOLDER for d in dims}
        r["SKU"]       = f"∑ {sku}"
        r["_row_type"] = "sku_subtotal"
        for m in metrics:
            r[m] = sku_totals[m]
        rows.append(r)

    # Grand total
    grand = work[metrics].sum()
    r = {d: PLACEHOLDER for d in dims}
    r["SKU"]       = "GRAND TOTAL"
    r["_row_type"] = "grand_total"
    for m in metrics:
        r[m] = grand[m]
    rows.append(r)

    result = pd.DataFrame(rows, columns=dims + metrics + ["_row_type"])
    return result


def format_pivot(df: pd.DataFrame) -> pd.DataFrame:
    """Format metric columns for display."""
    out = df.copy()
    if "Clicks" in out.columns:
        out["Clicks"] = out["Clicks"].apply(
            lambda x: f"{int(x):,}" if isinstance(x, (int, float)) else x
        )
    for col in ["Spend", "Sales"]:
        if col in out.columns:
            out[col] = out[col].apply(
                lambda x: f"${x:,.2f}" if isinstance(x, (int, float)) else x
            )
    return out


def render_metrics(df):
    impr   = int(df["Impressions"].sum())
    clicks = int(df["Clicks"].sum())
    spend  = df["Spend"].sum()
    sales  = df["Sales"].sum()
    orders = int(df["Orders"].sum())
    acos   = spend / sales   if sales  > 0 else 0
    roas   = sales / spend   if spend  > 0 else 0
    ctr    = clicks / impr   if impr   > 0 else 0
    cpc    = spend / clicks  if clicks > 0 else 0
    cvr    = orders / clicks if clicks > 0 else 0

    row1 = st.columns(5)
    for col, (lbl, val) in zip(row1, [
        ("💰 Spend",  f"${spend:,.2f}"),
        ("📈 Sales",  f"${sales:,.2f}"),
        ("📦 Orders", f"{orders:,}"),
        ("🎯 ACOS",   f"{acos*100:.2f}%"),
        ("🔄 ROAS",   f"{roas:.2f}"),
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


def render_filters(df, sheet_name):
    st.markdown('<div class="filter-header">🔍 Filters</div>', unsafe_allow_html=True)

    row1 = st.columns(5)

    port_opts = sorted([
        v for v in df[PORTFOLIO_COL].dropna().unique()
        if str(v).strip() not in NULL_VALUES
    ])
    if port_opts:
        sel = row1[0].multiselect("Portfolio", port_opts, key=f"{sheet_name}_portfolio")
        if sel:
            df = df[df[PORTFOLIO_COL].isin(sel)]

    state_opts = sorted([
        v for v in df[CAMPAIGN_STATE_COL].dropna().unique()
        if str(v).strip() not in NULL_VALUES
    ])
    if state_opts:
        sel = row1[1].multiselect("Campaign State", state_opts, key=f"{sheet_name}_state")
        if sel:
            df = df[df[CAMPAIGN_STATE_COL].isin(sel)]

    status_opts = sorted([
        v for v in df[SERVING_STATUS_COL].dropna().unique()
        if str(v).strip() not in NULL_VALUES
    ])
    if status_opts:
        sel = row1[2].multiselect("Serving Status", status_opts, key=f"{sheet_name}_serving")
        if sel:
            df = df[df[SERVING_STATUS_COL].isin(sel)]

    for idx, (field, label) in enumerate([("BRAND", "Brand"), ("SKU", "SKU")]):
        if field in df.columns:
            opts = sorted([v for v in df[field].dropna().unique() if v is not None])
            if opts:
                sel = row1[3 + idx].multiselect(label, opts, key=f"{sheet_name}_{field}")
                if sel:
                    df = df[df[field].isin(sel)]

    remaining = [f for f in FILTER_FIELDS if f[0] not in ("BRAND", "SKU")]
    row2 = st.columns(6)
    for idx, (field, label) in enumerate(remaining):
        if field not in df.columns:
            continue
        opts = sorted([v for v in df[field].dropna().unique() if v is not None])
        if not opts:
            continue
        sel = row2[idx % 6].multiselect(label, opts, key=f"{sheet_name}_{field}")
        if sel:
            df = df[df[field].isin(sel)]

    st.markdown("---")
    return df


def render_table(df, sheet_name):
    display_cols = [
        CAMPAIGN_NAME_COL, PORTFOLIO_COL,
        CAMPAIGN_STATE_COL, SERVING_STATUS_COL,
        "Impressions", "Clicks", "Spend", "Sales", "Orders", "Units",
        "Conv. Rate", "ACOS", "CPC", "ROAS",
    ]
    display_cols = [c for c in display_cols if c in df.columns]
    show = df[display_cols].copy()

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

    show = show.rename(columns={
        CAMPAIGN_NAME_COL:  "Campaign Name",
        PORTFOLIO_COL:      "Portfolio",
        CAMPAIGN_STATE_COL: "Campaign State",
        SERVING_STATUS_COL: "Serving Status",
    })

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


def render_pivot(df, sheet_name):
    st.markdown('<div class="section-header">📊 Pivot by SKU / Type / Targeting</div>',
                unsafe_allow_html=True)
    st.caption("Hierarchical breakdown with subtotals per SKU and Type. "
               "SKU subtotals (🟠) and Grand Total (🔶) are highlighted.")

    pivot_raw = build_pivot(df)

    if pivot_raw.empty:
        st.info("Not enough data to build a pivot for the current filter selection.")
        return

    row_types = pivot_raw["_row_type"].tolist()
    display   = format_pivot(pivot_raw.drop(columns=["_row_type"]))

    # Style rows by type
    def style_row(row):
        rt = row_types[row.name]
        if rt == "grand_total":
            return ["background-color: #FF9900; color: white; font-weight: 800;"] * len(row)
        elif rt == "sku_subtotal":
            return ["background-color: #FFE0B2; font-weight: 700;"] * len(row)
        elif rt == "type_subtotal":
            return ["background-color: #FFF8F0; font-weight: 600; color: #555;"] * len(row)
        else:
            return [""] * len(row)

    styled = display.style.apply(style_row, axis=1)
    st.dataframe(styled, use_container_width=True,
                 height=min(60 + len(display) * 35, 700))

    # Download
    csv = pivot_raw.drop(columns=["_row_type"]).to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download pivot CSV",
        csv,
        file_name=f"{sheet_name.replace(' ', '_')}_pivot.csv",
        mime="text/csv",
        key=f"pivot_dl_{sheet_name}",
    )


# ── App ───────────────────────────────────────────────────────────────────────

st.title("📊 PPC Campaign Dashboard")

uploaded = st.file_uploader(
    "Drop your Amazon PPC bulk Excel file (.xlsx)",
    type=["xlsx"],
    help="File should contain: Sponsored Products Campaigns, Sponsored Brands Campaigns, "
         "SB Multi Ad Group Campaigns, Sponsored Display Campaigns"
)

if not uploaded:
    st.info("👆 Upload your weekly PPC bulk file to get started. "
            "Files are processed in memory and never stored.")
    st.stop()

with st.spinner("Loading data..."):
    data = load_and_process(uploaded.read())

available = {k: v for k, v in data.items() if v is not None and len(v) > 0}

if not available:
    st.error("No valid PPC campaign data found in the uploaded file.")
    st.stop()

st.success(f"✅ Loaded {sum(len(v) for v in available.values()):,} campaigns "
           f"across {len(available)} tabs")

tab_objects = st.tabs([TAB_LABELS.get(s, s) for s in TAB_SHEETS if s in available])

for tab_obj, sheet_name in zip(tab_objects, [s for s in TAB_SHEETS if s in available]):
    with tab_obj:
        df = available[sheet_name].copy()

        # 1. KPI metrics
        render_metrics(df)

        # 2. Filters (visible row between metrics and table)
        df = render_filters(df, sheet_name)

        # 3. Campaign table
        st.markdown('<div class="section-header">📋 Campaign Table</div>',
                    unsafe_allow_html=True)
        render_table(df, sheet_name)

        st.markdown("")

        # 4. Pivot table
        render_pivot(df, sheet_name)
