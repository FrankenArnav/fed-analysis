import streamlit as st
import subprocess
import json
import os
import pandas as pd
import re
import sys
import time
from pathlib import Path

try:
    import plotly.express as px
    PLOTLY_OK = True
except ImportError:
    PLOTLY_OK = False

# Auto-load variables from a local .env file (e.g. COMTRADE_API_KEY) into the
# process environment, so the API Key field below pre-fills without the user
# having to paste it in manually every session. Safe no-op if python-dotenv
# isn't installed — falls back to whatever's already in the OS environment.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Both helper scripts (create_sector_config_from_qe.py, sector_comtrade_pipeline.py)
# print Unicode symbols (✓ ✅ ✗ ⚠ → etc.) to stdout. On Windows, a Python child
# process whose stdout is redirected to a pipe (as it is here, via subprocess)
# defaults to the system's legacy codepage (e.g. cp1252) instead of UTF-8 unless
# told otherwise — encoding any of those symbols then raises UnicodeEncodeError
# and crashes the subprocess before it finishes its work. Forcing UTF-8 via the
# child's environment fixes this on every platform without touching the scripts.
SUBPROCESS_ENV = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}

# Set page config
st.set_page_config(page_title="Comtrade Auto-Pipeline", page_icon="⚡", layout="wide")

# Pre-mapped DGCIS QE Groups for common sectors
SECTOR_PRESETS = {
    # NOTE: these strings are matched against the "Major Commodity Groups"
    # column of QE_PC_HS_Mapping_2025-26_Final_Sent.xlsx (case-insensitive,
    # exact match for single groups; combine-mode also requires an exact
    # match per group). They MUST exactly match the mapping file's wording
    # or create_sector_config_from_qe.py silently skips the group (combine
    # mode) or hard-fails (single-group mode), producing an empty/zero-code
    # config with no useful error surfaced to the user.
    #
    # Verified against the mapping file's actual 31 Major Commodity Groups:
    # several presets previously used "&"/wrong wording where the file uses
    # "AND"/different phrasing, silently producing empty or partial sectors
    # (Leather Products, Gems & Jewellery, Chemicals all hard-failed or
    # produced 0 matched codes; Textiles & Apparel silently dropped its
    # "MAN-MADE YARN/FABS./MADEUPS ETC." group). Fixed below.
    "Textiles & Apparel": "RMG OF ALL TEXTILES;CARPET;JUTE MFG. INCLUDING FLOOR COVERING;MAN-MADE YARN/FABS./MADEUPS ETC.;COTTON YARN/FABS./MADEUPS, HANDLOOM PRODUCTS ETC.",
    "Electronic Goods": "ELECTRONIC GOODS",
    "Drugs & Pharmaceuticals": "DRUGS AND PHARMACEUTICALS",
    "Engineering Goods": "ENGINEERING GOODS",
    "Leather Products": "LEATHER AND LEATHER MANUFACTURES",
    "Gems & Jewellery": "GEMS AND JEWELLERY",
    # "AGRO CHEMICALS" is a Principal Commodity sub-group already contained
    # within "ORGANIC AND INORGANIC CHEMICALS" — it isn't a separate Major
    # Commodity Group, so combining it as a second group is both wrong-named
    # and redundant. A single group captures everything.
    "Chemicals": "ORGANIC AND INORGANIC CHEMICALS",
    "Marine Products": "MARINE PRODUCTS",
    "Custom / Manual Entry": ""
}

# ---------------------------------------------------------
# CHART METADATA — sector-agnostic display labels for each
# sheet in the *_context_graph_data.xlsx workbook produced
# by `--mode visuals_data`. The underlying data itself is
# always computed dynamically from whatever sector/config
# the user is running (buckets, competitors and importing
# markets are all derived at run time) — nothing here is
# hardcoded to a specific sector. "{sector}" is filled in
# with whatever sector the user picked in Step 1.
# ---------------------------------------------------------
CHART_META = {
    "Chart1_GlobalExports_Bn": {
        "title": "Global {sector} Exports by Product Bucket",
        "unit": "Bn USD",
        "kind": "fixed",
    },
    "Chart2_IndiaExports_Bn": {
        "title": "India's {sector} Exports by Product Bucket",
        "unit": "Bn USD",
        "kind": "fixed",
    },
    "Chart3_TopBucketCountries": {
        "title": "Top Product Bucket — India vs Key Competitors",
        "unit": "Bn USD",
        "kind": "fixed",
    },
    "Chart4_TotalSectorCountries": {
        "title": "Total {sector} Trade — India vs Key Competitors",
        "unit": "Bn USD",
        "kind": "fixed",
    },
    "Chart5_MarketShares_Pct": {
        "title": "Export Market Shares in Top Importing Markets",
        "unit": "%",
        "kind": "fixed",
    },
    "Chart6_ImporterTotals_Bn": {
        "title": "Total Imports — Top Importing Markets",
        "unit": "Bn USD",
        "kind": "fixed",
    },
    "Chart7_ChinaShares_Pct": {
        "title": "China's Export Share in Top Importing Markets",
        "unit": "%",
        "kind": "fixed",
    },
    "Chart8_MarketShareTrends": {
        "title": "Market Share Trends Over Time",
        "unit": "%",
        "kind": "fixed",
    },
    # ── Ranking-style sheets — single metric, sorted descending, support
    # a "Top N" selector instead of a fixed series count. ──
    "Ranking_Exporters_Bn": {
        "title": "Top Exporting Countries — {sector}",
        "unit": "Bn USD",
        "kind": "ranking",
    },
    "Ranking_Importers_Bn": {
        "title": "Top Importing Countries — {sector}",
        "unit": "Bn USD",
        "kind": "ranking",
    },
    "Ranking_SectorBuckets_Bn": {
        "title": "Largest Segments in {sector}",
        "unit": "Bn USD",
        "kind": "ranking",
    },
    # ── Trend sheet — wide Country × Year table, supports a free pick of
    # which countries to compare. ──
    "CountryTrend_AllSector_Bn": {
        "title": "Competitor Comparison Over Time — {sector}",
        "unit": "Bn USD",
        "kind": "trend",
    },
}

CHART_TYPES = ["Bar (Grouped)", "Bar (Stacked)", "Line", "Area", "Pie"]

# ---------------------------------------------------------
# COLOR PALETTES — "FED Classic" is extracted directly from the theme
# named "FED Final" inside the reference deck "Sector's context setting
# slides.pptx" (its accent1-6 + hyperlink + dk2 colors). The other presets
# are derived variations so users can switch the overall look from a
# slider-style control without typing hex codes.
# ---------------------------------------------------------
PALETTE_PRESETS = {
    "FED Classic":     ["#133E68", "#009F75", "#FEB95F", "#9F4A54", "#7DE2D1", "#0097A7", "#C2C1C2", "#595959"],
    "FED Bold":        ["#0A2C4D", "#00B386", "#FF9F1C", "#7A1F2B", "#37CFB5", "#005F73", "#8C8C8C", "#262626"],
    "FED Pastel":      ["#6E8FB8", "#5FC9A8", "#FFD08A", "#C98A92", "#B7ECE0", "#5FB8C9", "#D9D9D9", "#8C8C8C"],
    "Monochrome Navy": ["#133E68", "#2C5683", "#46709E", "#608ABA", "#7AA4D6", "#94BEF1", "#0A2640", "#C2C1C2"],
}
DEFAULT_PALETTE = "FED Classic"


def get_config_field(config_path, sheet, field_name):
    """Read a Field/Value-style sheet from the sector config workbook and
    return the Value for a given Field name. Returns None if not found."""
    try:
        xls = pd.ExcelFile(config_path)
        if sheet not in xls.sheet_names:
            return None
        df = pd.read_excel(xls, sheet, header=1)
        df.columns = [str(c).strip() for c in df.columns]
        if "Field" not in df.columns or "Value" not in df.columns:
            return None
        match = df[df["Field"].astype(str).str.strip().str.lower() == field_name.lower()]
        if not match.empty:
            val = match.iloc[0]["Value"]
            if pd.notna(val):
                return str(val).strip()
    except Exception:
        pass
    return None


def locate_chart_data_workbook(config_path):
    """Find the *_context_graph_data.xlsx produced for the active config.
    Works for any sector — derives the expected filename from the config's
    own 'Sector slug' field (and an optional visual_settings override)
    rather than assuming any particular sector."""
    candidates = []
    if config_path and os.path.exists(config_path):
        override = get_config_field(config_path, "visual_settings", "Output chart data workbook name")
        slug = get_config_field(config_path, "sector_details", "Sector slug")
        if override:
            candidates.append(override)
        if slug:
            candidates.append(f"{slug}_context_graph_data.xlsx")
    for c in candidates:
        if c and os.path.exists(c):
            return c
    # Fallback: newest matching file sitting in the working directory
    matches = [f for f in os.listdir(".") if f.endswith("_context_graph_data.xlsx")]
    if matches:
        matches.sort(key=lambda f: os.path.getmtime(f), reverse=True)
        return matches[0]
    return None


def locate_cleaned_workbook(config_path):
    """Find the *_comtrade_cleaned.xlsx produced for the active config.
    Same sector-agnostic lookup strategy as locate_chart_data_workbook."""
    candidates = []
    if config_path and os.path.exists(config_path):
        slug = get_config_field(config_path, "sector_details", "Sector slug")
        if slug:
            candidates.append(f"{slug}_comtrade_cleaned.xlsx")
    for c in candidates:
        if c and os.path.exists(c):
            return c
    matches = [f for f in os.listdir(".") if f.endswith("_comtrade_cleaned.xlsx")]
    if matches:
        matches.sort(key=lambda f: os.path.getmtime(f), reverse=True)
        return matches[0]
    return None


def diagnose_cleaned_workbook(cleaned_path):
    """Run sanity checks on a cleaned workbook so failures are surfaced as
    clear diagnostic messages instead of a blank/incorrect-looking output.
    Returns a list of (level, message) tuples; level is error/warning/success."""
    findings = []
    if not cleaned_path or not os.path.exists(cleaned_path):
        return [("error", "No cleaned workbook found yet for this sector. Run an output type "
                           "that includes the cleaning step first.")]
    try:
        xls = pd.ExcelFile(cleaned_path)
    except Exception as e:
        return [("error", f"Could not open the cleaned workbook: {e}")]

    sheets = xls.sheet_names
    required = ["Master_Table", "World_Exports", "World_Imports", "India_Exports",
                "India_Imports", "Summary_Tables"]
    missing = [s for s in required if s not in sheets]
    if missing:
        findings.append(("warning", f"Cleaned workbook is missing expected sheet(s): {', '.join(missing)}."))

    competitor_sheets = ["Competitor_Exporters", "Competitor_Importers",
                         "Competitor_Market_Share", "Competitor_Trends", "Competitor_Summary"]
    if not any(s in sheets for s in competitor_sheets):
        findings.append(("warning", "No competitor analysis sheets found yet — choose the "
                                     "'Cleaned Excel Workbook (incl. Competitor Analysis)' output type "
                                     "(or run the Full Auto-Sequence) to add the 5 Competitor_* sheets."))

    try:
        master = pd.read_excel(xls, "Master_Table", header=1, dtype={"HS Code": str})
        master.columns = [str(c).strip() for c in master.columns]
        if master.empty:
            findings.append(("error", "Master_Table has no rows — the Comtrade API likely returned no "
                                       "data for this sector's HS codes/years/flows. Check the Errors_Log "
                                       "sheet in the downloaded workbook for the specific cause."))
        else:
            if "Sector Bucket" in master.columns:
                other_share = (master["Sector Bucket"].astype(str).str.strip().str.lower() == "other").mean()
                if other_share > 0.5:
                    findings.append(("warning",
                        f"{other_share:.0%} of rows fall into the 'Other' sector bucket — the HS-code-to-"
                        "bucket mapping (Step 2 'Buckets' tab) may not cover the HS codes that were actually "
                        "fetched. Double-check the Buckets tab before re-running."))
            findings.append(("success", f"Master_Table has {len(master):,} rows"
                              + (f" covering years {int(master['Year'].min())}–{int(master['Year'].max())}."
                                 if "Year" in master.columns and master["Year"].notna().any() else ".")))
    except Exception as e:
        findings.append(("error", f"Could not read Master_Table: {e}"))

    if "Errors_Log" in sheets:
        try:
            err_df = pd.read_excel(xls, "Errors_Log", header=1)
            if not err_df.empty:
                findings.append(("warning", f"Errors_Log contains {len(err_df)} logged issue(s) from this "
                                             "run — open the downloaded workbook's Errors_Log sheet for details."))
        except Exception:
            pass

    if not findings:
        findings.append(("success", "Cleaned workbook looks healthy."))
    return findings


def render_master_table_explorer(cleaned_path, sector_name, palette=None):
    """Ad-hoc filter/explore panel over the cleaned workbook's Master_Table —
    works for any sector since it reads whatever Reporters/Years/Buckets are
    actually present rather than assuming fixed values."""
    palette = palette or PALETTE_PRESETS[DEFAULT_PALETTE]
    st.subheader("🔎 Explore Cleaned Trade Data")

    if not PLOTLY_OK:
        st.error("The `plotly` package is required for charts. Run `pip install plotly`, then restart the app.")
        return

    try:
        master = pd.read_excel(cleaned_path, sheet_name="Master_Table", header=1, dtype={"HS Code": str})
        master.columns = [str(c).strip() for c in master.columns]
    except Exception as e:
        st.error(f"Could not load Master_Table from the cleaned workbook: {e}")
        return
    if master.empty:
        st.info("Master_Table has no rows yet — run the pipeline first.")
        return

    def _flow_label(src):
        s = str(src).lower()
        if "world_exports" in s: return "World Exports"
        if "world_imports" in s: return "World Imports"
        if "india_exports" in s: return "India Exports"
        if "india_imports" in s: return "India Imports"
        return "Other"

    master["Flow Type"] = master["Source File"].apply(_flow_label) if "Source File" in master.columns else "Other"
    flow_options = [f for f in ["World Exports", "World Imports", "India Exports", "India Imports"]
                    if f in master["Flow Type"].unique()]

    years = sorted(master["Year"].dropna().astype(int).unique().tolist()) if "Year" in master.columns else []

    # Reporters ranked by total trade value so the dropdown is sorted by
    # relevance (largest traders first) rather than an arbitrary/alphabetical list.
    if "Reporter" in master.columns and "Trade Value USD Mn" in master.columns:
        reporter_rank = (master.groupby("Reporter")["Trade Value USD Mn"].sum()
                          .sort_values(ascending=False).index.tolist())
    else:
        reporter_rank = []

    bucket_options = sorted(master["Sector Bucket"].dropna().unique().tolist()) if "Sector Bucket" in master.columns else []

    c1, c2 = st.columns(2)
    with c1:
        yr_range = (st.slider("Year range", min(years), max(years), (min(years), max(years)), key="explore_years")
                    if years else None)
        flow_sel = st.multiselect("Trade flow", flow_options, default=flow_options, key="explore_flow")
    with c2:
        bucket_sel = st.multiselect("Sector bucket / HS product group", bucket_options, default=[],
                                     key="explore_bucket", help="Leave empty to include all buckets.")
        default_countries = [c for c in reporter_rank if c.lower() == "india"] + \
                             [c for c in reporter_rank if c.lower() != "india"][:5]
        country_sel = st.multiselect("Reporter / countries to compare (sorted by trade value)",
                                      reporter_rank, default=default_countries[:6], key="explore_countries")

    filtered = master.copy()
    if yr_range:
        filtered = filtered[(filtered["Year"] >= yr_range[0]) & (filtered["Year"] <= yr_range[1])]
    if flow_sel:
        filtered = filtered[filtered["Flow Type"].isin(flow_sel)]
    if bucket_sel:
        filtered = filtered[filtered["Sector Bucket"].isin(bucket_sel)]
    if country_sel:
        filtered = filtered[filtered["Reporter"].isin(country_sel)]

    if filtered.empty:
        st.info("No rows match the selected filters — widen the year range or flow/country/bucket selection.")
        return

    metric = st.radio("Metric to chart", ["Trade Value (USD Mn)", "Unit Price (USD per quantity unit)"],
                       horizontal=True, key="explore_metric",
                       help="Unit Price = export/import value divided by quantity, as reported by Comtrade.")

    if metric == "Unit Price (USD per quantity unit)":
        # Unit price only makes sense within ONE quantity unit at a time —
        # USD/kg and USD/item aren't comparable. Comtrade reports whichever
        # unit is natural for each product, and rows with no usable quantity
        # (qty<=0/blank — value-only or weight-only records) are already
        # excluded upstream in the pipeline's build_master_table, so "Quantity"
        # here is blank for those.
        has_qty = "Quantity" in filtered.columns and "Qty Unit" in filtered.columns
        priced = filtered[filtered["Quantity"].notna() & (filtered["Quantity"] > 0)] if has_qty else filtered.iloc[0:0]
        unit_options = sorted({u for u in priced.get("Qty Unit", pd.Series(dtype=str)).dropna().tolist() if u})

        if not unit_options:
            st.info("No quantity data reported for the current filter selection — unit price can't be "
                    "computed. Try a different sector bucket/HS chapter (some products are value-only in Comtrade).")
            return

        qty_unit_sel = st.selectbox(
            "Quantity unit (unit price is only comparable within a single unit)",
            unit_options, key="explore_qty_unit")
        scoped = priced[priced["Qty Unit"] == qty_unit_sel]

        # Quantity-weighted average (sum value / sum quantity) rather than the
        # average of each row's own unit price — keeps a few small, oddly
        # priced shipments from skewing the chart. Matches how the cleaned
        # workbook's Unit_Prices sheet computes the same figure.
        agg = (scoped.groupby(["Year", "Reporter"], as_index=False)
               .agg(**{"Trade Value USD": ("Trade Value USD", "sum"), "Quantity": ("Quantity", "sum")}))
        agg["Unit Price (USD)"] = (agg["Trade Value USD"] / agg["Quantity"]).round(4)
        agg = agg.rename(columns={"Reporter": "Country"})
        value_col   = "Unit Price (USD)"
        unit_label  = f"USD per {qty_unit_sel}"
        chart_title = f"{sector_name} — unit price (USD per {qty_unit_sel})"
        pie_title   = f"{sector_name} unit price (USD/{qty_unit_sel})"
    else:
        agg = filtered.groupby(["Year", "Reporter"], as_index=False)["Trade Value USD Mn"].sum()
        agg = agg.rename(columns={"Reporter": "Country"})
        value_col   = "Trade Value USD Mn"
        unit_label  = "USD Mn"
        chart_title = f"{sector_name} — filtered trade value (USD Mn)"
        pie_title   = f"{sector_name} trade value"

    if agg.empty:
        st.info("No data to chart for this filter combination.")
        return

    chart_type = st.selectbox("Chart type", CHART_TYPES, key="explore_chart_type")
    if chart_type == "Bar (Stacked)":
        fig = px.bar(agg, x="Year", y=value_col, color="Country", barmode="stack", color_discrete_sequence=palette)
    elif chart_type == "Bar (Grouped)":
        fig = px.bar(agg, x="Year", y=value_col, color="Country", barmode="group", color_discrete_sequence=palette)
    elif chart_type == "Area":
        fig = px.area(agg, x="Year", y=value_col, color="Country", color_discrete_sequence=palette)
    elif chart_type == "Pie":
        last_yr = agg["Year"].max()
        fig = px.pie(agg[agg["Year"] == last_yr], names="Country", values=value_col, color_discrete_sequence=palette)
        fig.update_layout(title=f"{pie_title} — {last_yr}", margin=dict(t=70))
        st.plotly_chart(fig, use_container_width=True)
        with st.expander("View filtered rows"):
            st.dataframe(filtered, use_container_width=True)
        return
    else:
        fig = px.line(agg, x="Year", y=value_col, color="Country", markers=True, color_discrete_sequence=palette)

    fig.update_layout(title=chart_title, yaxis_title=unit_label,
                       legend_title_text="", margin=dict(t=70))
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("View filtered rows"):
        st.dataframe(filtered, use_container_width=True)


def render_geography_drilldown(config_file_path, cleaned_path, sector_name, api_key, palette=None):
    """Geography Drill-Down: on-demand bilateral reporter<->partner trade
    lookup (e.g. "China imports from the US"), distinct from the Reporter-
    vs-World comparisons in render_master_table_explorer above. Two view
    modes:
      - Single pair: one reporter, one partner (original behaviour).
      - Compare multiple partners: one reporter, several partner countries,
        each fetched on demand (one API call per partner — never a full
        country×country matrix), plus an optional extra World-total call so
        the chart can show an accurate % share of the reporter's real total
        trade (with a "Rest of World" remainder) instead of just a share
        among the countries picked.
    Both modes trigger sector_comtrade_pipeline.py via the same subprocess
    pattern as the main Run Pipeline button, then read back whatever rows
    are saved in the Geography_Bilateral sheet (append/dedupe, so earlier
    lookups stay available too)."""
    palette = palette or PALETTE_PRESETS[DEFAULT_PALETTE]
    st.subheader("🌍 Geography Drill-Down — Country-to-Country Trade")
    st.caption("Look up trade between any two specific countries — e.g. “how much does China import from "
               "the US” — or compare several partner countries' share of one country's trade in an item.")

    if not PLOTLY_OK:
        st.error("The `plotly` package is required for charts. Run `pip install plotly`, then restart the app.")
        return

    def _run_geo_fetch(cmd, spinner_text):
        """Shared subprocess-streaming runner for both the single-pair and
        compare-multiple-partners fetch buttons — same pattern as the main
        Run Pipeline button, with accurate success/no-data/error messaging."""
        log_box = st.empty()
        log_text = ""
        with st.spinner(spinner_text):
            try:
                process = subprocess.Popen(
                    cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                    env=SUBPROCESS_ENV, encoding="utf-8", errors="replace",
                )
                for line in iter(process.stdout.readline, ''):
                    log_text += line
                    log_box.code(log_text[-2000:], language="text")
                process.wait()
                if process.returncode != 0:
                    st.error("❌ Fetch failed. Check the log above.")
                elif ("No bilateral data returned for this pair/flow/years" in log_text
                      or "No data returned for any selected partner" in log_text):
                    st.warning(
                        "⚠️ Comtrade returned no data for this exact combination of country/countries, "
                        "flow, HS code(s), and year range — nothing was added to the Geography_Bilateral "
                        "sheet. This usually means the trade doesn't exist (or is too small to report) "
                        "for this specific combination, not a fetch error. Try different countries, "
                        "widen the year range, or check the flow direction (export vs. import)."
                    )
                else:
                    st.success("✅ Saved to the Geography_Bilateral sheet in the cleaned workbook.")
            except Exception as e:
                st.error(f"An unexpected error occurred: {e}")

    def _run_top_partners_fetch(cmd, spinner_text):
        """Subprocess-streaming runner for the 'Find top partners' button —
        same pattern as _run_geo_fetch above, but watches for --mode
        top_partners' own success/no-data log lines. This mode never
        touches the Geography_Bilateral sheet — it writes a separate small
        JSON snapshot file that gets read back just below."""
        log_box = st.empty()
        log_text = ""
        with st.spinner(spinner_text):
            try:
                process = subprocess.Popen(
                    cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                    env=SUBPROCESS_ENV, encoding="utf-8", errors="replace",
                )
                for line in iter(process.stdout.readline, ''):
                    log_text += line
                    log_box.code(log_text[-2000:], language="text")
                process.wait()
                if process.returncode != 0:
                    st.error("❌ Could not find top partners. Check the log above.")
                elif "TOP PARTNERS COMPLETE" in log_text:
                    st.success("✅ Top partners found — see the tables below.")
                else:
                    st.warning("⚠️ Comtrade returned no partner-breakdown data for this reporter/HS/year "
                               "combination — try widening the HS codes or year range above.")
            except Exception as e:
                st.error(f"An unexpected error occurred: {e}")

    @st.cache_data(ttl=86400, show_spinner="Loading country list (first time only)...")
    def _country_ref():
        import sector_comtrade_pipeline as scp
        return scp.get_country_reference()

    try:
        ref = _country_ref()
    except Exception as e:
        st.warning(f"Could not load the country list (needs an internet connection the first time this "
                   f"runs): {e}")
        return

    names = ref["name"].tolist()
    name_to_code = dict(zip(ref["name"], ref["code"]))

    def _default_name(preferred, fallback_idx):
        for p in preferred:
            if p in names:
                return p
        return names[fallback_idx] if names else ""

    view_mode = st.radio("View", ["Single pair", "Compare multiple partners"],
                         horizontal=True, key="geo_view_mode")

    if "geo_reporter_country" not in st.session_state:
        st.session_state["geo_reporter_country"] = _default_name(["India"], 0)
    if "geo_partner_country" not in st.session_state:
        st.session_state["geo_partner_country"] = _default_name(
            ["United States of America", "USA", "China"], min(1, len(names) - 1))

    def _swap_countries():
        r = st.session_state["geo_reporter_country"]
        p = st.session_state["geo_partner_country"]
        st.session_state["geo_reporter_country"], st.session_state["geo_partner_country"] = p, r

    partner_names = []
    include_world = False
    if view_mode == "Single pair":
        c1, c2, c3 = st.columns([5, 1, 5])
        with c1:
            st.selectbox("Reporter country", names, key="geo_reporter_country")
        with c2:
            st.write("")
            st.button("🔄", help="Swap reporter and partner", on_click=_swap_countries, key="geo_swap_btn")
        with c3:
            st.selectbox("Partner country", names, key="geo_partner_country")
    else:
        st.selectbox("Reporter country", names, key="geo_reporter_country")
        default_partners = [n for n in [_default_name(["United States of America"], 0),
                                        _default_name(["China"], 0)]
                            if n and n != st.session_state["geo_reporter_country"]]
        # Streamlit forbids writing to a widget's session_state key once that
        # widget has been instantiated in the run. The "Add these N to partner
        # countries" button below can't update geo_partners_multi directly, so
        # it stashes the merged list in geo_partners_pending and reruns; this
        # block applies that pending value BEFORE the multiselect widget is
        # created, which is the one place it's still legal to do so.
        if "geo_partners_pending" in st.session_state:
            st.session_state["geo_partners_multi"] = st.session_state.pop("geo_partners_pending")
        partner_names = st.multiselect(
            "Partner countries to compare", names,
            default=default_partners[:2], key="geo_partners_multi")
        include_world = st.checkbox(
            "Include World total (accurate % share + \"Rest of World\" slice — 1 extra fetch)",
            value=True, key="geo_include_world")

    reporter_name = st.session_state["geo_reporter_country"]
    reporter_code = name_to_code.get(reporter_name)
    if view_mode == "Single pair":
        partner_name = st.session_state["geo_partner_country"]
        partner_code = name_to_code.get(partner_name)

    # Year range — read straight from the active config's "years" sheet
    # (same auto-detect-End-Year rule as the pipeline's read_config) so the
    # slider matches whatever this sector was actually set up with.
    import datetime as _dt
    _auto_latest = _dt.date.today().year - 1
    try:
        dfy = pd.read_excel(config_file_path, "years", header=1)
        dfy.columns = [str(c).strip() for c in dfy.columns]
        cfg_start = int(float(dfy.iloc[0]["Start Year"]))
        _ey = dfy.iloc[0].get("End Year", None)
        cfg_end = (int(float(_ey)) if pd.notna(_ey) and str(_ey).strip().lower() not in ("", "auto", "latest", "nan")
                  else _auto_latest)
    except Exception:
        cfg_start, cfg_end = _auto_latest - 10, _auto_latest
    # The slider's upper bound always reaches at least last year, even if the
    # config file's own End Year is older — lets you pull newer years on
    # demand here without ever having to edit that file.
    slider_max = max(cfg_end, _auto_latest)

    # HS codes currently included for this sector — default fetch scope.
    try:
        dfh = pd.read_excel(config_file_path, "hs_codes", header=1, dtype=str)
        dfh.columns = [str(c).strip() for c in dfh.columns]
        incl_col = next((c for c in dfh.columns if "include" in c.lower()), None)
        sector_hs = (dfh[dfh[incl_col].str.strip().str.lower() == "include"]["HS Code"]
                    .dropna().str.strip().tolist()) if incl_col else []
    except Exception:
        sector_hs = []

    cc1, cc2 = st.columns(2)
    with cc1:
        flow_choice = st.radio("Flow", ["Exports (Reporter → Partner)", "Imports (Reporter ← Partner)"],
                               horizontal=True, key="geo_flow")
        flow_code = "X" if flow_choice.startswith("Exports") else "M"
    with cc2:
        yr_range = st.slider("Year range", cfg_start, slider_max, (max(cfg_start, cfg_end - 9), cfg_end),
                             key="geo_years")

    hs_sel = st.multiselect(
        f"HS codes to fetch (leave empty to use this sector's full list — {len(sector_hs)} codes)",
        sector_hs, default=[], key="geo_hs_filter")
    active_hs = [str(h).strip() for h in (hs_sel if hs_sel else sector_hs)]
    if hs_sel:
        st.caption(f"📊 Chart below will sum only the {len(active_hs)} HS code(s) selected above "
                   f"({', '.join(active_hs[:8])}{'…' if len(active_hs) > 8 else ''}) — fetch this "
                   f"exact combination first if you haven't already; the chart only shows rows "
                   f"already saved for it, not rows from a previous broader fetch.")
    else:
        st.caption(f"📊 Chart below sums all {len(active_hs)} codes in this sector's full HS list.")

    with st.expander("🔎 Find top partners (which countries should I compare?)", expanded=False):
        st.caption("Adding partner countries above is mostly guesswork without this. Fetch the "
                   f"{reporter_name or 'reporter'}'s full partner breakdown for the HS code(s)/years "
                   "currently selected, and see its top trading partners ranked by trade value — "
                   "both top export destinations and top import sources, in one click.")
        tp_top_n = st.slider("How many top partners per direction", 5, 10, 10, key="geo_top_n")
        tp_clicked = st.button("🔎 Find top partners", key="geo_top_partners_btn")

        if tp_clicked:
            if not reporter_code:
                st.error("Pick a reporter country.")
            elif not api_key:
                st.error("API Key is required (set it in the field above, in Step 3).")
            else:
                os.environ["COMTRADE_API_KEY"] = api_key
                tp_years = [str(y) for y in range(yr_range[0], yr_range[1] + 1)]
                cmd = [sys.executable, "sector_comtrade_pipeline.py", "--config", config_file_path,
                      "--mode", "top_partners", "--reporter", reporter_code, "--top-n", str(tp_top_n)]
                if active_hs:
                    cmd += ["--hs", ",".join(active_hs)]
                if tp_years:
                    cmd += ["--years", ",".join(tp_years)]
                _run_top_partners_fetch(cmd, f"Finding {reporter_name}'s top trading partners...")

        tp_slug = get_config_field(config_file_path, "sector_details", "Sector slug")
        tp_json_path = f"{tp_slug}_top_partners.json" if tp_slug else None
        if tp_json_path and os.path.exists(tp_json_path):
            try:
                tp_data = json.loads(Path(tp_json_path).read_text(encoding="utf-8"))
            except Exception:
                tp_data = None
            if tp_data and tp_data.get("reporter_code") == reporter_code:
                st.caption(f"Last found for **{tp_data.get('reporter', '')}** — "
                          f"{len(tp_data.get('hs_codes', []))} HS code(s), "
                          f"{tp_data.get('year_start', '')}–{tp_data.get('year_end', '')} "
                          f"(generated {tp_data.get('generated_at', '')})")
                tcol1, tcol2 = st.columns(2)
                for col, dkey, label in (
                    (tcol1, "exports", f"Top export destinations for {tp_data.get('reporter', '')}"),
                    (tcol2, "imports", f"Top import sources for {tp_data.get('reporter', '')}"),
                ):
                    with col:
                        st.markdown(f"**{label}**")
                        rows = tp_data.get(dkey, [])
                        if not rows:
                            st.caption("No data returned for this direction.")
                            continue
                        tdf = pd.DataFrame(rows)
                        disp = tdf[["Rank", "Partner", "Trade Value (USD)", "Share %"]].copy()
                        disp["Trade Value (USD)"] = disp["Trade Value (USD)"].map(lambda v: f"${v:,.0f}")
                        disp["Share %"] = disp["Share %"].map(lambda v: f"{v:.2f}%")
                        st.dataframe(disp, hide_index=True, use_container_width=True)
                        add_names = [r["Partner"] for r in rows
                                    if r["Partner"] in names and r["Partner"] != reporter_name]
                        if add_names and view_mode == "Compare multiple partners":
                            if st.button(f"➕ Add these {len(add_names)} to partner countries to compare",
                                        key=f"geo_add_{dkey}"):
                                current = list(st.session_state.get("geo_partners_multi", []))
                                merged = current + [n for n in add_names if n not in current]
                                # Can't write geo_partners_multi directly here — the
                                # multiselect widget with that key already rendered
                                # earlier in this run. Stash it and rerun; the pending
                                # value is applied just before the widget is created
                                # on the next run (see geo_partners_pending above).
                                st.session_state["geo_partners_pending"] = merged
                                st.rerun()
            elif tp_data:
                st.caption("Top-partners results on file are for a different reporter country — "
                          "click “Find top partners” above to refresh for the current selection.")

    fetch_label = ("\U0001F4E1 Fetch bilateral trade data" if view_mode == "Single pair"
                   else "\U0001F4E1 Fetch comparison data")
    fetch_clicked = st.button(fetch_label, type="primary", key="geo_fetch_btn")

    if fetch_clicked:
        hs_arg    = active_hs
        years_arg = [str(y) for y in range(yr_range[0], yr_range[1] + 1)]

        if view_mode == "Single pair":
            if not reporter_code or not partner_code:
                st.error("Pick a reporter and partner country.")
            elif reporter_code == partner_code:
                st.error("Reporter and partner can't be the same country.")
            elif not api_key:
                st.error("API Key is required (set it in the field above, in Step 3).")
            else:
                os.environ["COMTRADE_API_KEY"] = api_key
                cmd = [sys.executable, "sector_comtrade_pipeline.py", "--config", config_file_path,
                      "--mode", "bilateral", "--reporter", reporter_code, "--partner", partner_code,
                      "--flow", flow_code]
                if hs_arg:
                    cmd += ["--hs", ",".join(hs_arg)]
                if years_arg:
                    cmd += ["--years", ",".join(years_arg)]
                _run_geo_fetch(cmd, f"Fetching {reporter_name} "
                               f"{'exports to' if flow_code == 'X' else 'imports from'} {partner_name}...")
        else:
            if not reporter_code:
                st.error("Pick a reporter country.")
            elif not partner_names:
                st.error("Pick at least one partner country to compare.")
            elif reporter_name in partner_names:
                st.error("Reporter can't also be one of the partners being compared.")
            elif not api_key:
                st.error("API Key is required (set it in the field above, in Step 3).")
            else:
                os.environ["COMTRADE_API_KEY"] = api_key
                partner_codes = [name_to_code[n] for n in partner_names]
                cmd = [sys.executable, "sector_comtrade_pipeline.py", "--config", config_file_path,
                      "--mode", "bilateral_compare", "--reporter", reporter_code,
                      "--partners", ",".join(partner_codes), "--flow", flow_code]
                if hs_arg:
                    cmd += ["--hs", ",".join(hs_arg)]
                if years_arg:
                    cmd += ["--years", ",".join(years_arg)]
                if not include_world:
                    cmd += ["--no-world-total"]
                _run_geo_fetch(cmd, f"Fetching {reporter_name}'s trade with {len(partner_names)} partner(s)...")

    # Show whatever has been saved to Geography_Bilateral so far (this run's
    # lookup plus any earlier ones — the sheet is append/dedupe, never
    # overwritten wholesale per lookup).
    geo_cleaned = cleaned_path or locate_cleaned_workbook(config_file_path)
    if not geo_cleaned or not os.path.exists(geo_cleaned):
        st.info("Run a lookup above to create the Geography_Bilateral sheet.")
        return
    try:
        xls = pd.ExcelFile(geo_cleaned)
        if "Geography_Bilateral" not in xls.sheet_names:
            st.info("No Geography Drill-Down lookups saved yet for this sector — run one above.")
            return
        geo_df = pd.read_excel(xls, "Geography_Bilateral", header=1, dtype={"HS Code": str})
        geo_df.columns = [str(c).strip() for c in geo_df.columns]
    except Exception as e:
        st.warning(f"Could not read the Geography_Bilateral sheet: {e}")
        return

    if geo_df.empty:
        st.info("Geography_Bilateral sheet is empty — run a lookup above.")
        return

    chart_type = st.selectbox("Chart type", CHART_TYPES, key="geo_chart_type")

    if view_mode == "Single pair":
        pair_df = geo_df[(geo_df["Reporter"] == reporter_name) & (geo_df["Partner"] == partner_name)
                         & (geo_df["HS Code"].astype(str).str.strip().isin(active_hs))]
        if pair_df.empty:
            st.info(f"No saved data yet for {reporter_name} → {partner_name} with the HS code(s) "
                    f"currently selected above. Run a lookup above (or widen the HS filter).")
            with st.expander("View all saved Geography Drill-Down lookups"):
                st.dataframe(geo_df, use_container_width=True)
            return

        metric = st.radio("Metric to chart", ["Trade Value (USD)", "Unit Price"],
                          horizontal=True, key="geo_metric")

        if metric == "Unit Price":
            priced = pair_df[pair_df["Quantity"].notna() & (pair_df["Quantity"] > 0)]
            if priced.empty:
                st.info("No quantity data reported for this pair — unit price can't be computed.")
            else:
                up = (priced.groupby(["Year", "Flow", "Qty Unit"], as_index=False)
                     .agg(**{"Trade Value (USD)": ("Trade Value (USD)", "sum"), "Quantity": ("Quantity", "sum")}))
                up["Unit Price"] = (up["Trade Value (USD)"] / up["Quantity"]).round(4)
                unit_sel = st.selectbox("Quantity unit", sorted(up["Qty Unit"].dropna().unique().tolist()),
                                        key="geo_qty_unit")
                wide = (up[up["Qty Unit"] == unit_sel]
                       .pivot(index="Flow", columns="Year", values="Unit Price").reset_index())
                fig = render_chart(wide, "Flow", chart_type, f"{reporter_name} ↔ {partner_name} — unit price",
                                   f"USD per {unit_sel}", kind="trend", palette=palette)
                st.plotly_chart(fig, use_container_width=True)
        else:
            agg = pair_df.groupby(["Year", "Flow"], as_index=False).agg(
                **{"Trade Value (USD)": ("Trade Value (USD)", "sum")})
            agg["Year"] = agg["Year"].astype(int)
            wide = agg.pivot(index="Flow", columns="Year", values="Trade Value (USD)").reset_index()
            fig = render_chart(wide, "Flow", chart_type, f"{reporter_name} ↔ {partner_name} — trade value (USD)",
                               "USD", kind="trend", palette=palette)
            st.plotly_chart(fig, use_container_width=True)

        with st.expander("View filtered rows for this pair"):
            st.dataframe(pair_df, use_container_width=True)
        with st.expander("View all saved Geography Drill-Down lookups"):
            st.dataframe(geo_df, use_container_width=True)

    else:
        flow_label_variants = {"X": ["Export", "Exports"], "M": ["Import", "Imports"]}
        wanted_partners = partner_names + (["World"] if include_world else [])
        cmp_df = geo_df[(geo_df["Reporter"] == reporter_name)
                       & (geo_df["Partner"].isin(wanted_partners))
                       & (geo_df["Flow"].isin(flow_label_variants[flow_code]))
                       & (geo_df["HS Code"].astype(str).str.strip().isin(active_hs))]
        if cmp_df.empty:
            st.info(f"No saved comparison data yet for {reporter_name} in this flow direction with the "
                    f"HS code(s) currently selected above. Run a lookup above (or widen the HS filter).")
            with st.expander("View all saved Geography Drill-Down lookups"):
                st.dataframe(geo_df, use_container_width=True)
            return

        pct_share = st.checkbox("Show as % share", value=True, key="geo_pct_share")

        agg = cmp_df.groupby(["Partner", "Year"], as_index=False).agg(
            **{"Trade Value (USD)": ("Trade Value (USD)", "sum")})
        agg["Year"] = agg["Year"].astype(int)
        wide = agg.pivot(index="Partner", columns="Year", values="Trade Value (USD)").fillna(0)

        has_world = "World" in wide.index
        if has_world:
            world_row = wide.loc["World"]
            others    = wide.drop(index="World")
            row_sel   = others.reindex([p for p in partner_names if p in others.index]).fillna(0)
            rest      = (world_row - row_sel.sum()).clip(lower=0)
            row_sel.loc["Rest of World"] = rest
            wide = row_sel
        else:
            wide = wide.reindex([p for p in partner_names if p in wide.index])

        unit_label = "USD"
        if pct_share:
            col_sums = wide.sum(axis=0).replace(0, pd.NA)
            wide = (wide.div(col_sums, axis=1) * 100).round(2)
            unit_label = "% share"

        plot_df = wide.reset_index().rename(columns={"index": "Partner"})

        title_suffix = "imports from" if flow_code == "M" else "exports to"
        share_label = ("% of true total trade" if (pct_share and has_world)
                      else "% among selected partners" if pct_share
                      else "trade value (USD)")
        fig = render_chart(plot_df, "Partner", chart_type,
                           f"{reporter_name} — {title_suffix} selected partners ({share_label})",
                           unit_label, kind="trend", palette=palette)
        st.plotly_chart(fig, use_container_width=True)

        if pct_share and not has_world:
            st.caption("ℹ️ \"Include World total\" wasn't used, so these percentages are shares among "
                      "the selected partners only — not the reporter's true total trade in this item.")

        with st.expander("View filtered rows for this comparison"):
            st.dataframe(cmp_df, use_container_width=True)
        with st.expander("View all saved Geography Drill-Down lookups"):
            st.dataframe(geo_df, use_container_width=True)


def render_chart(df, id_col, chart_type, title, unit, kind="fixed", palette=None):
    palette = palette or PALETTE_PRESETS[DEFAULT_PALETTE]
    value_cols = [c for c in df.columns if c != id_col]

    # ── Ranking sheets: one metric, many categories (countries/segments).
    # Color each bar/slice individually by category rather than by series,
    # since there's only one series — that reads much better for a ranking. ──
    if kind == "ranking":
        value_col = value_cols[-1]
        plot_df = df.copy()
        plot_df[id_col] = plot_df[id_col].astype(str)

        if chart_type == "Pie":
            fig = px.pie(plot_df, names=id_col, values=value_col, color_discrete_sequence=palette)
            fig.update_layout(title=title, margin=dict(t=70))
            return fig
        elif chart_type == "Line":
            fig = px.line(plot_df, x=id_col, y=value_col, markers=True, color_discrete_sequence=palette)
        elif chart_type == "Area":
            fig = px.area(plot_df, x=id_col, y=value_col, color_discrete_sequence=palette)
        else:  # Bar (Grouped/Stacked) — single series, color by category
            fig = px.bar(plot_df, x=id_col, y=value_col, color=id_col, color_discrete_sequence=palette)
            fig.update_layout(showlegend=False)

        fig.update_layout(title=title, yaxis_title=unit, xaxis_title="", margin=dict(t=70))
        return fig

    # ── Trend sheet: wide Country × Year table. We want Year on the x-axis
    # and each selected Country as its own colored line/series — the
    # opposite axis arrangement from the "fixed" sheets below. ──
    if kind == "trend":
        long_df = df.melt(id_vars=id_col, value_vars=value_cols, var_name="Year", value_name="Value")
        long_df["Year"] = long_df["Year"].astype(str)

        if chart_type == "Bar (Stacked)":
            fig = px.bar(long_df, x="Year", y="Value", color=id_col, barmode="stack", color_discrete_sequence=palette)
        elif chart_type == "Bar (Grouped)":
            fig = px.bar(long_df, x="Year", y="Value", color=id_col, barmode="group", color_discrete_sequence=palette)
        elif chart_type == "Area":
            fig = px.area(long_df, x="Year", y="Value", color=id_col, color_discrete_sequence=palette)
        elif chart_type == "Pie":
            last_year = value_cols[-1]
            fig = px.pie(df, names=id_col, values=last_year, color_discrete_sequence=palette)
            fig.update_layout(title=f"{title} — {last_year}", margin=dict(t=70))
            return fig
        else:  # Line — default and best fit for a multi-year trend
            fig = px.line(long_df, x="Year", y="Value", color=id_col, markers=True, color_discrete_sequence=palette)

        fig.update_layout(title=title, yaxis_title=unit, legend_title_text="", margin=dict(t=70))
        return fig

    # ── Fixed sheets: the original Chart1-8 multi-series tables ──
    long_df = df.melt(id_vars=id_col, value_vars=value_cols, var_name="Series", value_name="Value")
    long_df[id_col] = long_df[id_col].astype(str)

    if chart_type == "Bar (Grouped)":
        fig = px.bar(long_df, x=id_col, y="Value", color="Series", barmode="group", color_discrete_sequence=palette)
    elif chart_type == "Bar (Stacked)":
        fig = px.bar(long_df, x=id_col, y="Value", color="Series", barmode="stack", color_discrete_sequence=palette)
    elif chart_type == "Line":
        fig = px.line(long_df, x=id_col, y="Value", color="Series", markers=True, color_discrete_sequence=palette)
    elif chart_type == "Area":
        fig = px.area(long_df, x=id_col, y="Value", color="Series", color_discrete_sequence=palette)
    else:  # Pie — uses the most recent / last value column
        last_col = value_cols[-1]
        fig = px.pie(df, names=id_col, values=last_col, color_discrete_sequence=palette)
        fig.update_layout(title=f"{title} — {last_col}", margin=dict(t=70))
        return fig

    fig.update_layout(title=title, yaxis_title=unit, legend_title_text="", margin=dict(t=70))
    return fig


def show_interactive_visuals(chart_path, sector_name):
    st.subheader("📊 Interactive Visuals")
    st.caption(f"Source data: `{os.path.basename(chart_path)}` — regenerated automatically for whatever sector you're working with.")

    if not PLOTLY_OK:
        st.error("The `plotly` package is required for interactive charts. Run `pip install plotly` in your terminal, then restart the app.")
        return

    try:
        xls = pd.ExcelFile(chart_path)
    except Exception as e:
        st.error(f"Could not open chart data workbook: {e}")
        return

    available_sheets = [s for s in xls.sheet_names if s in CHART_META]
    if not available_sheets:
        st.info("No recognized chart datasets found in this workbook yet. Run 'Generate Visuals' first.")
        return

    label_map = {s: CHART_META[s]["title"].format(sector=sector_name) for s in available_sheets}

    # Color theme — a plain dropdown (not a slider — palette choices aren't an
    # ordered scale) with an optional "Custom" mode that exposes real color
    # pickers for each swatch. Defaults to the FED deck's own colors.
    palette_names = list(PALETTE_PRESETS.keys())
    palette_choice = st.selectbox(
        "🎨 Color theme",
        palette_names + ["Custom"],
        index=palette_names.index(DEFAULT_PALETTE),
        key="viz_palette",
    )
    if palette_choice == "Custom":
        base = PALETTE_PRESETS[DEFAULT_PALETTE]
        st.caption("Pick colors for the first few series — remaining series reuse the preset below them.")
        pick_cols = st.columns(4)
        custom_palette = []
        for i in range(4):
            with pick_cols[i]:
                custom_palette.append(
                    st.color_picker(f"Color {i + 1}", base[i % len(base)], key=f"custom_color_{i}")
                )
        active_palette = custom_palette + base[len(custom_palette):]
    else:
        active_palette = PALETTE_PRESETS[palette_choice]

    col1, col2 = st.columns([2, 1])
    with col1:
        chosen_label = st.selectbox("Choose a graph", list(label_map.values()), key="viz_choice")
    chosen_sheet = next(s for s, lbl in label_map.items() if lbl == chosen_label)

    df = pd.read_excel(xls, chosen_sheet)
    id_col = df.columns[0]
    meta = CHART_META[chosen_sheet]
    kind = meta.get("kind", "fixed")

    with col2:
        chart_type = st.selectbox("Chart type", CHART_TYPES, key="viz_chart_type")

    # Extra controls depending on the sheet's "kind":
    #   ranking → Top N selector (Top 5 / Top 10 / ...)
    #   trend   → which countries to compare, for competitor analysis
    plot_df = df
    if kind == "ranking":
        max_n = len(df)
        n_options = sorted({n for n in [5, 10, 15, 20] if n <= max_n})
        if not n_options or max_n not in n_options:
            n_options.append(max_n)
        n_options = sorted(set(n_options))
        default_n = 10 if 10 in n_options else n_options[-1]
        top_n = st.selectbox(
            "Show top",
            n_options,
            index=n_options.index(default_n),
            format_func=lambda n: f"Top {n}",
            key=f"topn_{chosen_sheet}",
        )
        plot_df = df.head(top_n)
    elif kind == "trend":
        all_countries = df[id_col].astype(str).tolist()
        default_sel = all_countries[:5] if len(all_countries) > 5 else all_countries
        selected_countries = st.multiselect(
            "Countries to compare (competitor analysis)",
            all_countries,
            default=default_sel,
            key=f"countries_{chosen_sheet}",
        )
        plot_df = df[df[id_col].astype(str).isin(selected_countries)] if selected_countries else df.iloc[0:0]

    if plot_df.empty:
        st.info("Select at least one item above to render this chart.")
        return

    fig = render_chart(plot_df, id_col, chart_type, label_map[chosen_sheet], meta["unit"], kind=kind, palette=active_palette)
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("View underlying data"):
        st.dataframe(plot_df, use_container_width=True)


st.title("⚡ Comtrade Sector Pipeline")
st.markdown("Select a sector, review the auto-generated configuration, run the pipeline, and explore the results as interactive charts.")

# Initialize session state variables
if "step" not in st.session_state:
    st.session_state.step = 1
if "config_file_path" not in st.session_state:
    st.session_state.config_file_path = None
if "run_timestamp" not in st.session_state:
    st.session_state.run_timestamp = None
if "sector_name" not in st.session_state:
    st.session_state.sector_name = None

def reset_app():
    st.session_state.step = 1
    st.session_state.config_file_path = None
    st.session_state.run_timestamp = None
    st.session_state.sector_name = None

st.sidebar.button("🔄 Reset Application", on_click=reset_app)

# ---------------------------------------------------------
# STEP 1: AUTO-GENERATE CONFIGURATION
# ---------------------------------------------------------
if st.session_state.step == 1:
    st.header("Step 1: Select Sector")

    col1, col2 = st.columns(2)
    with col1:
        sector_choice = st.selectbox("Choose a Sector Template", list(SECTOR_PRESETS.keys()))

        if sector_choice == "Custom / Manual Entry":
            sector_name = st.text_input("Sector Name", placeholder="e.g., Toys & Games")
            groups = st.text_input(
                "DGCIS QE Groups (Semicolon separated)",
                help="Optional. Only fill this in if you want to auto-pull HS codes from an "
                     "existing DGCIS Major Commodity Group (exact name required — see "
                     "create_sector_config_from_qe.py --list). Leave this BLANK for a fully "
                     "manual sector (e.g. one not covered by any DGCIS group, like Live Animals) "
                     "— just enter the HS code(s) directly in 'Additional Custom HS Codes' below."
            )
        else:
            sector_name = sector_choice
            groups = SECTOR_PRESETS[sector_choice]
            st.info(f"**Auto-mapped DGCIS Groups:**\n{groups.replace(';', ' | ')}")

    with col2:
        hs_level = st.selectbox(
            "HS Level Default", ["2", "4", "6"], index=0,
            help="HS chapter granularity used to build the sector_buckets mapping. "
                 "2-digit (chapter level) matches the reference cleaned-workbook format "
                 "and is recommended — 4 or 6 digit codes still work for HS-code filtering "
                 "but Sector Bucket grouping is always rolled up to 2-digit chapters downstream."
        )
        hs_filter_text = st.text_input(
            "HS Chapter Filter (optional, comma-separated 2-digit chapters)",
            placeholder="e.g., 64",
            help="Restricts the sector to only these HS chapters, dropping every other code the "
                 "QE group(s) would otherwise pull in. Use this when the product you want (e.g. "
                 "Footwear) is only part of a broader DGCIS group (e.g. 'LEATHER AND LEATHER "
                 "MANUFACTURES' also covers finished leather, leather garments, saddlery, etc.) "
                 "— leave blank to include every chapter the selected group(s) contain."
        )
        add_codes_text = st.text_area("Additional Custom HS Codes (comma separated, optional)", placeholder="e.g., 9031,9032")
        uploaded_hs_file = st.file_uploader("OR Upload CSV/Excel with Custom HS Codes (Codes in the first column)", type=["csv", "xlsx", "xls"])

    if st.button("Generate & Review Configuration", type="primary"):
        # Fully manual sector: "Custom / Manual Entry" chosen and the DGCIS QE
        # Groups field left blank. There's no DGCIS Major Commodity Group for
        # every sector (e.g. "Live Animals" / HS Chapter 01 isn't one of the
        # 31 groups in the QE/PC mapping file), so this path skips QE-group
        # matching entirely and builds the config from --add_codes only.
        is_manual = (sector_choice == "Custom / Manual Entry" and not groups)

        if not sector_name or (not groups and not is_manual):
            st.error("Please provide a sector name and at least one QE group.")
        elif is_manual and not add_codes_text and uploaded_hs_file is None:
            st.error(
                "Manual entry needs at least one HS code — type one into "
                "'Additional Custom HS Codes' or upload a file, since no DGCIS QE group was given."
            )
        else:
            # Process uploaded file
            file_hs_codes = []
            if uploaded_hs_file is not None:
                try:
                    if uploaded_hs_file.name.endswith('.csv'):
                        df_hs = pd.read_csv(uploaded_hs_file)
                    else:
                        df_hs = pd.read_excel(uploaded_hs_file)

                    # Assume HS codes are in the first column
                    if not df_hs.empty:
                        first_col = df_hs.columns[0]
                        # Convert to string, drop NAs, and clean up trailing .0 if read as floats
                        raw_codes = df_hs[first_col].dropna().astype(str).str.replace(r'\.0$', '', regex=True).tolist()
                        file_hs_codes.extend([c.strip() for c in raw_codes if c.strip()])
                except Exception as e:
                    st.error(f"Could not read uploaded file: {e}")
                    st.stop()

            # Combine manual text codes and file codes
            combined_codes = []
            if add_codes_text:
                combined_codes.extend([c.strip() for c in add_codes_text.split(',') if c.strip()])
            if file_hs_codes:
                combined_codes.extend(file_hs_codes)

            final_add_codes_str = ",".join(combined_codes)

            with st.spinner("Compiling sector mapping..."):
                cmd = [sys.executable, "create_sector_config_from_qe.py"]
                cmd.extend(["--sector", sector_name])
                if is_manual:
                    cmd.append("--manual")
                else:
                    cmd.extend(["--groups" if ";" in groups else "--group", groups])
                cmd.extend(["--hs_level", hs_level])

                # HS chapter filter only means something when codes are being
                # pulled from a QE group — manual/blank configs have no group
                # codes to filter, so skip it there even if the field was left
                # filled in from a previous selection.
                if hs_filter_text.strip() and not is_manual:
                    cmd.extend(["--hs_filter", hs_filter_text.strip()])

                # Pass the combined codes to the command line if any exist
                if final_add_codes_str:
                    cmd.extend(["--add_codes", final_add_codes_str])

                try:
                    result = subprocess.run(
                        cmd, capture_output=True, text=True, check=True,
                        env=SUBPROCESS_ENV, encoding="utf-8", errors="replace",
                    )
                    # IGNORECASE: the --manual CLI path prints "Blank config saved → ..."
                    # (lowercase "config"), while --group/--groups print "Config saved → ..."
                    match = re.search(r"config saved → (.*?\.xlsx)", result.stdout, re.IGNORECASE)
                    if match:
                        st.session_state.config_file_path = match.group(1).strip()
                        st.session_state.sector_name = sector_name
                        st.session_state.step = 2
                        st.rerun()
                    else:
                        st.error("Could not determine the generated filename. Check logs:")
                        st.code(result.stdout)
                except subprocess.CalledProcessError as e:
                    st.error("Error generating config:")
                    error_output = e.stderr if e.stderr and e.stderr.strip() else e.stdout
                    st.code(error_output if error_output else f"Exit code {e.returncode} (No terminal output captured).", language="text")
                except Exception as e:
                    st.error(f"An unexpected Python error occurred: {str(e)}")

# ---------------------------------------------------------
# STEP 2: REVIEW & EDIT
# ---------------------------------------------------------
elif st.session_state.step == 2:
    st.header("Step 2: Review Sector Parameters")
    st.markdown("We've automatically mapped the HS codes. Review them below and add any custom Buckets if needed.")

    xls = pd.ExcelFile(st.session_state.config_file_path)
    sheets = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

    tab1, tab2, tab3 = st.tabs(["📝 Details", "🪣 Buckets", "🔢 HS Codes"])

    with tab1:
        edited_details = st.data_editor(sheets.get("sector_details", pd.DataFrame()), use_container_width=True)
    with tab2:
        st.caption("Define custom analytic groups here. If left blank, the pipeline uses default HS chapters.")
        edited_buckets = st.data_editor(sheets.get("sector_buckets", pd.DataFrame()), num_rows="dynamic", use_container_width=True)
    with tab3:
        st.caption("Change 'Include' to 'Exclude' to drop irrelevant codes from the pipeline.")
        edited_hs_codes = st.data_editor(sheets.get("hs_codes", pd.DataFrame()), num_rows="dynamic", height=400, use_container_width=True)

    col1, col2 = st.columns([1, 5])
    with col1:
        if st.button("← Back"):
            st.session_state.step = 1
            st.rerun()
    with col2:
        if st.button("Save & Proceed to Run Pipeline", type="primary"):
            # Save edits silently
            with pd.ExcelWriter(st.session_state.config_file_path, engine="openpyxl") as writer:
                edited_details.to_excel(writer, sheet_name="sector_details", index=False)
                edited_buckets.to_excel(writer, sheet_name="sector_buckets", index=False)
                edited_hs_codes.to_excel(writer, sheet_name="hs_codes", index=False)
                for sheet in sheets:
                    if sheet not in ["sector_details", "sector_buckets", "hs_codes"]:
                        sheets[sheet].to_excel(writer, sheet_name=sheet, index=False)

            st.session_state.step = 3
            st.rerun()

# ---------------------------------------------------------
# STEP 3: RUN, EXPLORE VISUALS & DOWNLOAD OUTPUTS
# ---------------------------------------------------------
elif st.session_state.step == 3:
    st.header("Step 3: Execute & Explore")

    # Single point of API-key handling for the whole app: prefer a value
    # already configured via Streamlit Secrets (st.secrets["COMTRADE_API_KEY"],
    # set in .streamlit/secrets.toml or the Streamlit Cloud dashboard) or a
    # local .env-style environment variable, so a deployed app doesn't force
    # every user to paste the key in by hand. The text input still lets
    # anyone override it for a single run.
    _default_key = ""
    _secrets_paths = [
        Path.home() / ".streamlit" / "secrets.toml",
        Path(__file__).resolve().parent / ".streamlit" / "secrets.toml",
    ]
    if any(p.exists() for p in _secrets_paths):
        # Only touch st.secrets if a secrets.toml actually exists - st.secrets
        # itself renders a "No secrets files found" warning banner into the
        # app on first access when no file is present, even when the access
        # is wrapped in try/except (the banner isn't a Python exception, so
        # try/except can't suppress it).
        try:
            _default_key = st.secrets.get("COMTRADE_API_KEY", "")
        except Exception:
            pass
    if not _default_key:
        _default_key = os.environ.get("COMTRADE_API_KEY", "")

    api_key = st.text_input(
        "UN Comtrade API Key",
        value=_default_key,
        type="password",
        help="Pre-filled automatically if COMTRADE_API_KEY is set via Streamlit "
             "Secrets or a .env/environment variable. Paste a different key here "
             "to override it for this run.",
    )

    # Map human-readable dropdown options ("Output type") to lists of
    # execution modes. No PPTX/deck is ever produced by any of these —
    # "visuals_data" only writes the chart-data workbook used to render
    # live charts below, and "slide_ready" only writes an Excel summary-
    # tables workbook (the pipeline's run_pptx step is never invoked from
    # this app). "competitor_full" is chained right after "clean" so the
    # cleaned workbook always gets its 5 Competitor_* sheets — previously
    # this was missing because no dropdown option ever ran it.
    pipeline_modes_map = {
        "🚀 Full Auto-Sequence (Download → Clean → Competitor → Visuals → Summary)":
            ["full", "clean", "competitor_full", "visuals_data", "slide_ready"],
        "Setup Check Only": ["setup_check"],
        "Test Sample Only (1 HS code, 1 year)": ["test"],
        "Full Data Download Only": ["full"],
        "Cleaned Excel Workbook (incl. Competitor Analysis)": ["clean", "competitor_full"],
        "Generate Visuals Only": ["visuals_data"],
        "Generate Summary Tables Only": ["slide_ready"],
    }

    col1, col2 = st.columns(2)
    with col1:
        selected_label = st.selectbox("Select Output Type", list(pipeline_modes_map.keys()), index=0)
        modes_to_run = pipeline_modes_map[selected_label]
        st.caption("Pipeline modes that will run: " + " → ".join(f"`{m}`" for m in modes_to_run))

    # Year-range override — the config file's "years" sheet has its own
    # Start/End Year, but that's a static value baked in once at sector
    # setup time and easy to go stale. Rather than asking anyone to open
    # that file, expose it here: if the config's End Year is older than
    # "last year," default to overriding it so a run pulls the freshest
    # data automatically. Applies to every mode below via --start-year/
    # --end-year — the config file itself is never modified.
    import datetime as _dt3
    _auto_latest3 = _dt3.date.today().year - 1
    try:
        _dfy3 = pd.read_excel(st.session_state.config_file_path, "years", header=1)
        _dfy3.columns = [str(c).strip() for c in _dfy3.columns]
        _cfg_start3 = int(float(_dfy3.iloc[0]["Start Year"]))
        _ey3 = _dfy3.iloc[0].get("End Year", None)
        _cfg_end3 = (int(float(_ey3)) if pd.notna(_ey3) and str(_ey3).strip().lower() not in ("", "auto", "latest", "nan")
                    else _auto_latest3)
    except Exception:
        _cfg_start3, _cfg_end3 = _auto_latest3 - 10, _auto_latest3

    year_override = None
    with col2:
        override_on = st.checkbox(
            f"Override year range (config file currently says {_cfg_start3}–{_cfg_end3})",
            value=(_cfg_end3 < _auto_latest3), key="pipeline_year_override_on")
        if override_on:
            year_override = st.slider(
                "Years to fetch this run", min(_cfg_start3, _auto_latest3 - 10), _auto_latest3,
                (_cfg_start3, _auto_latest3), key="pipeline_year_override_range")
            st.caption(f"Will fetch {year_override[0]}–{year_override[1]} for this run only "
                       f"(config file's own Start/End Year is left untouched).")

    if st.button("🚀 Run Pipeline", type="primary"):
        # Check if an API key is required for any of the selected modes
        requires_api = any(m in ["test", "full"] for m in modes_to_run)

        if not api_key and requires_api:
            st.error("API Key is required to fetch data for the selected mode(s).")
        else:
            if api_key:
                os.environ["COMTRADE_API_KEY"] = api_key

            # Record the time before running to find freshly modified files later
            st.session_state.run_timestamp = time.time()

            log_container = st.empty()
            log_text = ""

            total_steps = len(modes_to_run)
            progress_bar = st.progress(0) if total_steps > 1 else None

            pipeline_failed = False

            # Execute each selected mode sequentially
            for idx, current_mode in enumerate(modes_to_run):
                step_indicator = f"Step {idx + 1}/{total_steps}: Running `{current_mode}`" if total_steps > 1 else f"Running `{current_mode}`"

                with st.spinner(f"{step_indicator}... Please wait."):
                    cmd = [sys.executable, "sector_comtrade_pipeline.py", "--config", st.session_state.config_file_path, "--mode", current_mode]
                    if year_override:
                        cmd += ["--start-year", str(year_override[0]), "--end-year", str(year_override[1])]

                    try:
                        process = subprocess.Popen(
                            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                            env=SUBPROCESS_ENV, encoding="utf-8", errors="replace",
                        )

                        for line in iter(process.stdout.readline, ''):
                            log_text += line
                            # Truncate to the last 2000 characters to prevent memory lag
                            log_container.code(log_text[-2000:], language="text")

                        process.wait()

                        if process.returncode != 0:
                            st.error(f"❌ Pipeline encountered an error during `{current_mode}`. Check logs above.")
                            pipeline_failed = True
                            break # Halt the sequence if a step fails

                    except Exception as e:
                        st.error(f"An unexpected error occurred: {e}")
                        pipeline_failed = True
                        break

                # Update progress bar if there are multiple steps
                if progress_bar:
                    progress_bar.progress((idx + 1) / total_steps)

            if not pipeline_failed:
                st.success("✅ Pipeline execution complete!")

                # Diagnostics: if this run touched the cleaned workbook, check it
                # for the common failure modes (empty Master_Table, missing
                # competitor sheets, mostly-"Other" bucket mapping) so a broken
                # sector shows a clear message instead of a blank/odd-looking output.
                if any(m in ("clean", "competitor_full", "competitor_test") for m in modes_to_run):
                    cleaned_check = locate_cleaned_workbook(st.session_state.config_file_path)
                    st.markdown("**Cleaned workbook diagnostics:**")
                    for level, msg in diagnose_cleaned_workbook(cleaned_check):
                        {"error": st.error, "warning": st.warning, "success": st.success}[level](msg)

    # Interactive Visuals Section: shows live, filterable charts built from
    # whatever sector the user is currently working with — no deck involved.
    chart_data_path = locate_chart_data_workbook(st.session_state.config_file_path)
    if chart_data_path:
        st.divider()
        show_interactive_visuals(chart_data_path, st.session_state.sector_name or "Sector")

    # Master_Table explorer: free-form filtering (year range, trade flow,
    # reporter/competitor countries, sector bucket) over the cleaned workbook,
    # independent of the fixed Chart1-8 set above.
    cleaned_path_for_explore = locate_cleaned_workbook(st.session_state.config_file_path)
    if cleaned_path_for_explore:
        st.divider()
        render_master_table_explorer(cleaned_path_for_explore, st.session_state.sector_name or "Sector")

    # Geography Drill-Down: bilateral reporter<->partner lookup. Shown even
    # before the cleaned workbook exists yet (a lookup here will create it),
    # unlike the explorer above which needs Master_Table to already exist.
    st.divider()
    render_geography_drilldown(st.session_state.config_file_path, cleaned_path_for_explore,
                               st.session_state.sector_name or "Sector", api_key)

    # File Download Section: Scans directory for newly updated Excel/CSV files
    if st.session_state.run_timestamp:
        st.divider()
        st.subheader("📥 Download Outputs")

        output_files_found = False
        current_dir = os.getcwd()

        for filename in os.listdir(current_dir):
            if filename.endswith((".xlsx", ".csv")):
                filepath = os.path.join(current_dir, filename)
                # If the file was modified AFTER the run button was clicked
                if os.path.getmtime(filepath) >= st.session_state.run_timestamp:
                    # Ignore the config file itself
                    if filename != st.session_state.config_file_path:
                        output_files_found = True
                        with open(filepath, "rb") as file_data:
                            file_bytes = file_data.read()

                            st.download_button(
                                label=f"Download 📊 {filename}",
                                data=file_bytes,
                                file_name=filename,
                                mime="application/octet-stream"
                            )

        if not output_files_found:
            st.info("No new output files generated yet.")
