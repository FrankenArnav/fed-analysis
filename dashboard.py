# ==============================================================================
# UNIFIED ANALYTICS HUB — ALL-IN-ONE STREAMLIT APPLICATION
# ==============================================================================
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import textwrap
import io
import time
import random
import os
import subprocess
import re
import sys
from pathlib import Path
import json

# Plotly Imports
try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_OK = True
except ImportError:
    PLOTLY_OK = False

# Optional dotenv for Comtrade
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ==============================================================================
# 0. MASTER PAGE CONFIGURATION (Must be the first Streamlit command)
# ==============================================================================
st.set_page_config(page_title="Unified Data Hub", layout="wide")

# ==============================================================================
# 1. GLOBAL CONSTANTS & CACHED FUNCTIONS
# ==============================================================================

# --- ASI CONSTANTS ---
ASI_NIC_NAMES = {
    "10": "Food Products",        "11": "Beverages",            "12": "Tobacco",              "13": "Textiles",
    "14": "Wearing Apparel",      "15": "Leather & Footwear",   "16": "Wood Products",        "17": "Paper",
    "18": "Printing",             "19": "Petroleum/Coke",       "20": "Chemicals",            "21": "Pharmaceuticals",
    "22": "Rubber & Plastics",    "23": "Non-metallic Minerals","24": "Basic Metals",         "25": "Fabricated Metals",
    "26": "Electronics",          "27": "Electrical Equipment", "28": "Machinery",            "29": "Automobiles & Auto",
    "30": "Other Transport",      "31": "Furniture",            "32": "Other Manufacturing",  "33": "Repair & Installation",
}

ASI_STATE_NAMES = {
    "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab", "04": "Chandigarh", "05": "Uttarakhand", 
    "06": "Haryana", "07": "Delhi", "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim", 
    "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur", "15": "Mizoram", "16": "Tripura", "17": "Meghalaya", 
    "18": "Assam", "19": "West Bengal", "20": "Jharkhand", "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh", 
    "24": "Gujarat", "25": "Daman & Diu", "26": "Dadra & Nagar Haveli", "27": "Maharashtra", "28": "Andhra Pradesh",
    "29": "Karnataka", "30": "Goa", "31": "Lakshadweep", "32": "Kerala", "33": "Tamil Nadu", "34": "Puducherry", 
    "35": "Andaman & Nicobar", "36": "Telangana", "37": "Ladakh"
}

PRESET_MAP = {
    "Food Products": ["10"], "Beverages": ["11"], "Tobacco": ["12"], "Textiles": ["13"],
    "Wearing Apparel": ["14"], "Leather & Footwear": ["15"], "Chemicals": ["20"], 
    "Pharmaceuticals": ["21"], "Electronics": ["26"], "Automobiles & Auto": ["29"]
}

# --- PLFS CONSTANTS ---
PLFS_NIC_DICTIONARY = {
    "10": "Food Products", "11": "Beverages", "12": "Tobacco Products", "13": "Textiles", 
    "14": "Wearing Apparel", "15": "Leather and Related Products", "16": "Wood and Wood Products", 
    "17": "Paper and Paper Products", "18": "Printing and Reproduction", "19": "Coke and Refined Petroleum", 
    "20": "Chemicals and Chemical Products", "21": "Pharmaceuticals & Botanicals", "22": "Rubber and Plastics Products", 
    "23": "Other Non-Metallic Minerals", "24": "Basic Metals", "25": "Fabricated Metal Products", 
    "26": "Computer, Electronic & Optical", "27": "Electrical Equipment", "28": "Machinery and Equipment n.e.c.", 
    "29": "Motor Vehicles and Trailers", "30": "Other Transport Equipment", "31": "Furniture", 
    "32": "Other Manufacturing", "33": "Repair & Installation"
}

PLFS_STATE_DICTIONARY = {
    "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab", "04": "Chandigarh", "05": "Uttarakhand", 
    "06": "Haryana", "07": "Delhi", "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim", 
    "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur", "15": "Mizoram", "16": "Tripura", "17": "Meghalaya", 
    "18": "Assam", "19": "West Bengal", "20": "Jharkhand", "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh", 
    "24": "Gujarat", "25": "Daman & Diu", "26": "Dadra & Nagar Haveli", "27": "Maharashtra", "28": "Andhra Pradesh",
    "29": "Karnataka", "30": "Goa", "31": "Lakshadweep", "32": "Kerala", "33": "Tamil Nadu", "34": "Puducherry", 
    "35": "Andaman & Nicobar Islands", "36": "Telangana", "37": "Ladakh", "38": "The Dadra & Nagar Haveli and Daman & Diu"
}

WORKER_CATEGORY_DICT = {
    "11": "Self-employed (Own account worker)", "12": "Self-employed (Employer)", "21": "Helper in household enterprise",
    "31": "Regular salaried/wage employee", "41": "Casual wage labour (Public works)", "51": "Casual wage labour (Other)",
    "81": "Unemployed (Seeking work)", "82": "Unemployed (Available, not seeking)", "91": "Student",
    "92": "Domestic duties", "93": "Domestic duties & free collection", "94": "Rentiers, pensioners, etc.",
    "95": "Not able to work due to disability", "97": "Others (begging, etc.)", "99": "Children (0-4 years)"
}

MARITAL_STATUS_DICT = {
    "1": "Never Married",
    "2": "Currently Married",
    "3": "Widowed",
    "4": "Divorced / Separated"
}

EDUCATION_DICT = {
    "01": "Not Literate",
    "02": "Literate (No Formal Schooling)",
    "03": "Below Primary",
    "04": "Primary",
    "05": "Middle",
    "06": "Secondary",
    "07": "Higher Secondary",
    "08": "Diploma / Certificate",
    "10": "Graduate",
    "11": "Postgraduate & Above"
}

# --- COMTRADE CONSTANTS & DIRECTORIES ---
SUBPROCESS_ENV = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}

COMTRADE_DIR = os.path.join(os.getcwd(), "comtrade")
COMTRADE_EXPORTS_DIR = os.path.join(COMTRADE_DIR, "exports")

if not os.path.exists(COMTRADE_EXPORTS_DIR):
    os.makedirs(COMTRADE_EXPORTS_DIR)

SECTOR_PRESETS_COMTRADE = {
    "Textiles & Apparel": "RMG OF ALL TEXTILES;CARPET;JUTE MFG. INCLUDING FLOOR COVERING;MAN-MADE YARN/FABS./MADEUPS ETC.;COTTON YARN/FABS./MADEUPS, HANDLOOM PRODUCTS ETC.",
    "Electronic Goods": "ELECTRONIC GOODS",
    "Drugs & Pharmaceuticals": "DRUGS AND PHARMACEUTICALS",
    "Engineering Goods": "ENGINEERING GOODS",
    "Leather Products": "LEATHER AND LEATHER MANUFACTURES",
    "Gems & Jewellery": "GEMS AND JEWELLERY",
    "Chemicals": "ORGANIC AND INORGANIC CHEMICALS",
    "Marine Products": "MARINE PRODUCTS",
    "Custom / Manual Entry": ""
}

CHART_META = {
    "Chart1_GlobalExports_Bn": {"title": "Global {sector} Exports by Product Bucket", "unit": "Bn USD", "kind": "fixed"},
    "Chart2_IndiaExports_Bn": {"title": "India's {sector} Exports by Product Bucket", "unit": "Bn USD", "kind": "fixed"},
    "Chart3_TopBucketCountries": {"title": "Top Product Bucket — India vs Key Competitors", "unit": "Bn USD", "kind": "fixed"},
    "Chart4_TotalSectorCountries": {"title": "Total {sector} Trade — India vs Key Competitors", "unit": "Bn USD", "kind": "fixed"},
    "Chart5_MarketShares_Pct": {"title": "Export Market Shares in Top Importing Markets", "unit": "%", "kind": "fixed"},
    "Chart6_ImporterTotals_Bn": {"title": "Total Imports — Top Importing Markets", "unit": "Bn USD", "kind": "fixed"},
    "Chart7_ChinaShares_Pct": {"title": "China's Export Share in Top Importing Markets", "unit": "%", "kind": "fixed"},
    "Chart8_MarketShareTrends": {"title": "Market Share Trends Over Time", "unit": "%", "kind": "fixed"},
    "Ranking_Exporters_Bn": {"title": "Top Exporting Countries — {sector}", "unit": "Bn USD", "kind": "ranking"},
    "Ranking_Importers_Bn": {"title": "Top Importing Countries — {sector}", "unit": "Bn USD", "kind": "ranking"},
    "Ranking_SectorBuckets_Bn": {"title": "Largest Segments in {sector}", "unit": "Bn USD", "kind": "ranking"},
    "CountryTrend_AllSector_Bn": {"title": "Competitor Comparison Over Time — {sector}", "unit": "Bn USD", "kind": "trend"},
}

CHART_TYPES = ["Bar (Grouped)", "Bar (Stacked)", "Line", "Area", "Pie"]

PALETTE_PRESETS = {
    "FED Classic":     ["#133E68", "#009F75", "#FEB95F", "#9F4A54", "#7DE2D1", "#0097A7", "#C2C1C2", "#595959"],
    "FED Bold":        ["#0A2C4D", "#00B386", "#FF9F1C", "#7A1F2B", "#37CFB5", "#005F73", "#8C8C8C", "#262626"],
    "FED Pastel":      ["#6E8FB8", "#5FC9A8", "#FFD08A", "#C98A92", "#B7ECE0", "#5FB8C9", "#D9D9D9", "#8C8C8C"],
    "Monochrome Navy": ["#133E68", "#2C5683", "#46709E", "#608ABA", "#7AA4D6", "#94BEF1", "#0A2640", "#C2C1C2"],
}
DEFAULT_PALETTE = "FED Classic"


# --- GLOBAL HELPER FUNCTIONS ---
@st.cache_data
def extract_section(raw_df, start_keyword, end_keyword=None):
    try:
        start_mask = raw_df[0].astype(str).str.strip().str.upper() == start_keyword.upper()
        if not start_mask.any(): return None
        start_idx = raw_df[start_mask].index[0]
        
        if end_keyword:
            end_mask = raw_df[0].astype(str).str.strip().str.upper() == end_keyword.upper()
            if end_mask.any():
                end_idx = raw_df[end_mask].index[0]
                section = raw_df.iloc[start_idx : end_idx].copy()
            else:
                section = raw_df.iloc[start_idx :].copy()
        else:
            section = raw_df.iloc[start_idx :].copy()
            
        header_mask = section[0].astype(str).str.contains('Report Date', na=False, case=False)
        if not header_mask.any(): return None
        header_idx = section[header_mask].index[0]
        
        clean_section = section.loc[header_idx:].copy()
        raw_columns = clean_section.iloc[0]
        clean_section = clean_section[1:]
        
        clean_section.set_index(clean_section.columns[0], inplace=True)
        clean_section.index.name = 'Line Item'
        
        normalized_cols = []
        for col in raw_columns[1:]: 
            col_str = str(col).strip()
            if '-' in col_str: normalized_cols.append(col_str.split('-')[0])
            else: normalized_cols.append(col_str)
        clean_section.columns = normalized_cols
        
        clean_section.dropna(how='all', inplace=True)
        clean_section.dropna(axis=1, how='all', inplace=True)
        clean_section = clean_section.apply(pd.to_numeric, errors='coerce')
        return clean_section
    except Exception as e:
        return None

def calculate_percentage_of_revenue(target_df, revenue_series):
    if target_df is None or revenue_series is None or revenue_series.empty: return target_df
    return (target_df.div(revenue_series, axis=1) * 100).round(2)

def calculate_aggregate_average(df_dict):
    if not df_dict: return None
    concat_df = pd.concat(df_dict.values())
    avg_df = concat_df.groupby(level=0).mean().round(2)
    sorted_cols = sorted([c for c in avg_df.columns if str(c).isnumeric()])
    other_cols = [c for c in avg_df.columns if not str(c).isnumeric()]
    return avg_df[sorted_cols + other_cols]

@st.cache_data
def process_plfs_base_data(file_buffer, layout_specs, col_names):
    df = pd.read_fwf(file_buffer, colspecs=layout_specs, names=col_names, dtype=str)
    
    for col in col_names:
        df[col] = df[col].str.strip().fillna("")
        
    df['mult'] = pd.to_numeric(df['mult'], errors='coerce')
    df['final_wt'] = df['mult'] / 100.0
    
    df['state_label'] = df['state_raw'].map(PLFS_STATE_DICTIONARY).fillna(df['state_raw'])
    df['region_label'] = df['sector_raw'].map({'1': 'Rural', '2': 'Urban'}).fillna("Unknown")
    df['gender_label'] = df['sex'].map({'1': 'Male', '2': 'Female', '3': 'Third Gender'}).fillna("Unknown")
    df['worker_category'] = df['pas'].map(WORKER_CATEGORY_DICT).fillna("Unclassified/Other")
    
    if 'marital_raw' in df.columns:
        df['marital_status'] = df['marital_raw'].map(MARITAL_STATUS_DICT).fillna("Unclassified")
    if 'edu_raw' in df.columns:
        df['education_level'] = df['edu_raw'].map(EDUCATION_DICT).fillna("Unclassified")
    
    df['age'] = pd.to_numeric(df['age_raw'], errors='coerce')
    df['age_group'] = pd.cut(df['age'], bins=[0, 14, 29, 45, 60, 120], labels=["0-14", "15-29", "30-45", "46-60", "60+"])
    df['pas_num'] = pd.to_numeric(df['pas'], errors='coerce')
    
    df['is_employed'] = df['pas_num'].between(11, 51)
    df['is_unemployed'] = df['pas_num'].between(81, 82)
    df['is_lf'] = df['is_employed'] | df['is_unemployed']
    df['nic_2d'] = df['ind_pas'].str[:2]
    return df

# --- COMTRADE HELPER FUNCTIONS ---
def get_config_field(config_path, sheet, field_name):
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
    candidates = []
    if config_path and os.path.exists(config_path):
        override = get_config_field(config_path, "visual_settings", "Output chart data workbook name")
        slug = get_config_field(config_path, "sector_details", "Sector slug")
        if override:
            candidates.append(os.path.join(COMTRADE_EXPORTS_DIR, override) if not os.path.isabs(override) else override)
        if slug:
            candidates.append(os.path.join(COMTRADE_EXPORTS_DIR, f"{slug}_context_graph_data.xlsx"))
    for c in candidates:
        if c and os.path.exists(c):
            return c
    if os.path.exists(COMTRADE_EXPORTS_DIR):
        matches = [f for f in os.listdir(COMTRADE_EXPORTS_DIR) if f.endswith("_context_graph_data.xlsx")]
        if matches:
            matches.sort(key=lambda f: os.path.getmtime(os.path.join(COMTRADE_EXPORTS_DIR, f)), reverse=True)
            return os.path.join(COMTRADE_EXPORTS_DIR, matches[0])
    return None

def locate_cleaned_workbook(config_path):
    candidates = []
    if config_path and os.path.exists(config_path):
        slug = get_config_field(config_path, "sector_details", "Sector slug")
        if slug:
            candidates.append(os.path.join(COMTRADE_EXPORTS_DIR, f"{slug}_comtrade_cleaned.xlsx"))
    for c in candidates:
        if c and os.path.exists(c):
            return c
    if os.path.exists(COMTRADE_EXPORTS_DIR):
        matches = [f for f in os.listdir(COMTRADE_EXPORTS_DIR) if f.endswith("_comtrade_cleaned.xlsx")]
        if matches:
            matches.sort(key=lambda f: os.path.getmtime(os.path.join(COMTRADE_EXPORTS_DIR, f)), reverse=True)
            return os.path.join(COMTRADE_EXPORTS_DIR, matches[0])
    return None

def diagnose_cleaned_workbook(cleaned_path):
    findings = []
    if not cleaned_path or not os.path.exists(cleaned_path):
        return [("error", "No cleaned workbook found yet for this sector. Run an output type that includes the cleaning step first.")]
    try:
        xls = pd.ExcelFile(cleaned_path)
    except Exception as e:
        return [("error", f"Could not open the cleaned workbook: {e}")]

    sheets = xls.sheet_names
    required = ["Master_Table", "World_Exports", "World_Imports", "India_Exports", "India_Imports", "Summary_Tables"]
    missing = [s for s in required if s not in sheets]
    if missing:
        findings.append(("warning", f"Cleaned workbook is missing expected sheet(s): {', '.join(missing)}."))

    competitor_sheets = ["Competitor_Exporters", "Competitor_Importers", "Competitor_Market_Share", "Competitor_Trends", "Competitor_Summary"]
    if not any(s in sheets for s in competitor_sheets):
        findings.append(("warning", "No competitor analysis sheets found yet — choose the 'Cleaned Excel Workbook (incl. Competitor Analysis)' output type (or run the Full Auto-Sequence) to add the 5 Competitor_* sheets."))

    try:
        master = pd.read_excel(xls, "Master_Table", header=1, dtype={"HS Code": str})
        master.columns = [str(c).strip() for c in master.columns]
        if master.empty:
            findings.append(("error", "Master_Table has no rows — the Comtrade API likely returned no data for this sector's HS codes/years/flows. Check the Errors_Log sheet in the downloaded workbook for the specific cause."))
        else:
            if "Sector Bucket" in master.columns:
                other_share = (master["Sector Bucket"].astype(str).str.strip().str.lower() == "other").mean()
                if other_share > 0.5:
                    findings.append(("warning", f"{other_share:.0%} of rows fall into the 'Other' sector bucket — the HS-code-to-bucket mapping (Step 2 'Buckets' tab) may not cover the HS codes that were actually fetched. Double-check the Buckets tab before re-running."))
            findings.append(("success", f"Master_Table has {len(master):,} rows" + (f" covering years {int(master['Year'].min())}–{int(master['Year'].max())}." if "Year" in master.columns and master["Year"].notna().any() else ".")))
    except Exception as e:
        findings.append(("error", f"Could not read Master_Table: {e}"))

    if "Errors_Log" in sheets:
        try:
            err_df = pd.read_excel(xls, "Errors_Log", header=1)
            if not err_df.empty:
                findings.append(("warning", f"Errors_Log contains {len(err_df)} logged issue(s) from this run — open the downloaded workbook's Errors_Log sheet for details."))
        except Exception:
            pass

    if not findings:
        findings.append(("success", "Cleaned workbook looks healthy."))
    return findings

def render_master_table_explorer(cleaned_path, sector_name, palette=None):
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
    flow_options = [f for f in ["World Exports", "World Imports", "India Exports", "India Imports"] if f in master["Flow Type"].unique()]

    years = sorted(master["Year"].dropna().astype(int).unique().tolist()) if "Year" in master.columns else []

    if "Reporter" in master.columns and "Trade Value USD Mn" in master.columns:
        reporter_rank = (master.groupby("Reporter")["Trade Value USD Mn"].sum().sort_values(ascending=False).index.tolist())
    else:
        reporter_rank = []

    bucket_options = sorted(master["Sector Bucket"].dropna().unique().tolist()) if "Sector Bucket" in master.columns else []

    c1, c2 = st.columns(2)
    with c1:
        yr_range = (st.slider("Year range", min(years), max(years), (min(years), max(years)), key="explore_years") if years else None)
        flow_sel = st.multiselect("Trade flow", flow_options, default=flow_options, key="explore_flow")
    with c2:
        bucket_sel = st.multiselect("Sector bucket / HS product group", bucket_options, default=[], key="explore_bucket", help="Leave empty to include all buckets.")
        default_countries = [c for c in reporter_rank if c.lower() == "india"] + [c for c in reporter_rank if c.lower() != "india"][:5]
        country_sel = st.multiselect("Reporter / countries to compare (sorted by trade value)", reporter_rank, default=default_countries[:6], key="explore_countries")

    filtered = master.copy()
    if yr_range: filtered = filtered[(filtered["Year"] >= yr_range[0]) & (filtered["Year"] <= yr_range[1])]
    if flow_sel: filtered = filtered[filtered["Flow Type"].isin(flow_sel)]
    if bucket_sel: filtered = filtered[filtered["Sector Bucket"].isin(bucket_sel)]
    if country_sel: filtered = filtered[filtered["Reporter"].isin(country_sel)]

    if filtered.empty:
        st.info("No rows match the selected filters — widen the year range or flow/country/bucket selection.")
        return

    metric = st.radio("Metric to chart", ["Trade Value (USD Mn)", "Unit Price (USD per quantity unit)"], horizontal=True, key="explore_metric")

    if metric == "Unit Price (USD per quantity unit)":
        has_qty = "Quantity" in filtered.columns and "Qty Unit" in filtered.columns
        priced = filtered[filtered["Quantity"].notna() & (filtered["Quantity"] > 0)] if has_qty else filtered.iloc[0:0]
        unit_options = sorted({u for u in priced.get("Qty Unit", pd.Series(dtype=str)).dropna().tolist() if u})

        if not unit_options:
            st.info("No quantity data reported for the current filter selection — unit price can't be computed. Try a different sector bucket/HS chapter (some products are value-only in Comtrade).")
            return

        qty_unit_sel = st.selectbox("Quantity unit (unit price is only comparable within a single unit)", unit_options, key="explore_qty_unit")
        scoped = priced[priced["Qty Unit"] == qty_unit_sel]

        agg = (scoped.groupby(["Year", "Reporter"], as_index=False).agg(**{"Trade Value USD": ("Trade Value USD", "sum"), "Quantity": ("Quantity", "sum")}))
        agg["Unit Price (USD)"] = (agg["Trade Value USD"] / agg["Quantity"]).round(4)
        agg = agg.rename(columns={"Reporter": "Country"})
        value_col, unit_label, chart_title, pie_title = "Unit Price (USD)", f"USD per {qty_unit_sel}", f"{sector_name} — unit price (USD per {qty_unit_sel})", f"{sector_name} unit price (USD/{qty_unit_sel})"
    else:
        agg = filtered.groupby(["Year", "Reporter"], as_index=False)["Trade Value USD Mn"].sum()
        agg = agg.rename(columns={"Reporter": "Country"})
        value_col, unit_label, chart_title, pie_title = "Trade Value USD Mn", "USD Mn", f"{sector_name} — filtered trade value (USD Mn)", f"{sector_name} trade value"

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
        with st.expander("View filtered rows"): st.dataframe(filtered, use_container_width=True)
        return
    else:
        fig = px.line(agg, x="Year", y=value_col, color="Country", markers=True, color_discrete_sequence=palette)

    fig.update_layout(title=chart_title, yaxis_title=unit_label, legend_title_text="", margin=dict(t=70))
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("View filtered rows"):
        st.dataframe(filtered, use_container_width=True)

def render_chart(df, id_col, chart_type, title, unit, kind="fixed", palette=None):
    palette = palette or PALETTE_PRESETS[DEFAULT_PALETTE]
    value_cols = [c for c in df.columns if c != id_col]

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
        else:  
            fig = px.bar(plot_df, x=id_col, y=value_col, color=id_col, color_discrete_sequence=palette)
            fig.update_layout(showlegend=False)

        fig.update_layout(title=title, yaxis_title=unit, xaxis_title="", margin=dict(t=70))
        return fig

    if kind == "trend":
        long_df = df.melt(id_vars=id_col, value_vars=value_cols, var_name="Year", value_name="Value")
        long_df["Year"] = long_df["Year"].astype(str)

        if chart_type == "Bar (Stacked)": fig = px.bar(long_df, x="Year", y="Value", color=id_col, barmode="stack", color_discrete_sequence=palette)
        elif chart_type == "Bar (Grouped)": fig = px.bar(long_df, x="Year", y="Value", color=id_col, barmode="group", color_discrete_sequence=palette)
        elif chart_type == "Area": fig = px.area(long_df, x="Year", y="Value", color=id_col, color_discrete_sequence=palette)
        elif chart_type == "Pie":
            last_year = value_cols[-1]
            fig = px.pie(df, names=id_col, values=last_year, color_discrete_sequence=palette)
            fig.update_layout(title=f"{title} — {last_year}", margin=dict(t=70))
            return fig
        else: 
            fig = px.line(long_df, x="Year", y="Value", color=id_col, markers=True, color_discrete_sequence=palette)

        fig.update_layout(title=title, yaxis_title=unit, legend_title_text="", margin=dict(t=70))
        return fig

    long_df = df.melt(id_vars=id_col, value_vars=value_cols, var_name="Series", value_name="Value")
    long_df[id_col] = long_df[id_col].astype(str)

    if chart_type == "Bar (Grouped)": fig = px.bar(long_df, x=id_col, y="Value", color="Series", barmode="group", color_discrete_sequence=palette)
    elif chart_type == "Bar (Stacked)": fig = px.bar(long_df, x=id_col, y="Value", color="Series", barmode="stack", color_discrete_sequence=palette)
    elif chart_type == "Line": fig = px.line(long_df, x=id_col, y="Value", color="Series", markers=True, color_discrete_sequence=palette)
    elif chart_type == "Area": fig = px.area(long_df, x=id_col, y="Value", color="Series", color_discrete_sequence=palette)
    else:  
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

    palette_names = list(PALETTE_PRESETS.keys())
    palette_choice = st.selectbox("🎨 Color theme", palette_names + ["Custom"], index=palette_names.index(DEFAULT_PALETTE), key="viz_palette")
    if palette_choice == "Custom":
        base = PALETTE_PRESETS[DEFAULT_PALETTE]
        st.caption("Pick colors for the first few series — remaining series reuse the preset below them.")
        pick_cols = st.columns(4)
        custom_palette = []
        for i in range(4):
            with pick_cols[i]: custom_palette.append(st.color_picker(f"Color {i + 1}", base[i % len(base)], key=f"custom_color_{i}"))
        active_palette = custom_palette + base[len(custom_palette):]
    else:
        active_palette = PALETTE_PRESETS[palette_choice]

    col1, col2 = st.columns([2, 1])
    with col1: chosen_label = st.selectbox("Choose a graph", list(label_map.values()), key="viz_choice")
    chosen_sheet = next(s for s, lbl in label_map.items() if lbl == chosen_label)

    df = pd.read_excel(xls, chosen_sheet)
    id_col = df.columns[0]
    meta = CHART_META[chosen_sheet]
    kind = meta.get("kind", "fixed")

    with col2: chart_type = st.selectbox("Chart type", CHART_TYPES, key="viz_chart_type")

    plot_df = df
    if kind == "ranking":
        max_n = len(df)
        n_options = sorted({n for n in [5, 10, 15, 20] if n <= max_n})
        if not n_options or max_n not in n_options: n_options.append(max_n)
        n_options = sorted(set(n_options))
        default_n = 10 if 10 in n_options else n_options[-1]
        top_n = st.selectbox("Show top", n_options, index=n_options.index(default_n), format_func=lambda n: f"Top {n}", key=f"topn_{chosen_sheet}")
        plot_df = df.head(top_n)
    elif kind == "trend":
        all_countries = df[id_col].astype(str).tolist()
        default_sel = all_countries[:5] if len(all_countries) > 5 else all_countries
        selected_countries = st.multiselect("Countries to compare (competitor analysis)", all_countries, default=default_sel, key=f"countries_{chosen_sheet}")
        plot_df = df[df[id_col].astype(str).isin(selected_countries)] if selected_countries else df.iloc[0:0]

    if plot_df.empty:
        st.info("Select at least one item above to render this chart.")
        return

    fig = render_chart(plot_df, id_col, chart_type, label_map[chosen_sheet], meta["unit"], kind=kind, palette=active_palette)
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("View underlying data"):
        st.dataframe(plot_df, use_container_width=True)

def render_geography_drilldown(config_file_path, cleaned_path, sector_name, api_key, palette=None):
    palette = palette or PALETTE_PRESETS[DEFAULT_PALETTE]
    st.subheader("🌍 Geography Drill-Down — Country-to-Country Trade")
    st.caption("Look up trade between any two specific countries — e.g. “how much does China import from "
               "the US” — or compare several partner countries' share of one country's trade in an item.")

    if not PLOTLY_OK:
        st.error("The `plotly` package is required for charts. Run `pip install plotly`, then restart the app.")
        return

    def _run_geo_fetch(cmd, spinner_text):
        log_box = st.empty()
        log_text = ""
        with st.spinner(spinner_text):
            try:
                process = subprocess.Popen(
                    cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                    env=SUBPROCESS_ENV, encoding="utf-8", errors="replace", cwd=COMTRADE_EXPORTS_DIR
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
        log_box = st.empty()
        log_text = ""
        with st.spinner(spinner_text):
            try:
                process = subprocess.Popen(
                    cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                    env=SUBPROCESS_ENV, encoding="utf-8", errors="replace", cwd=COMTRADE_EXPORTS_DIR
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
        if COMTRADE_DIR not in sys.path:
            sys.path.insert(0, COMTRADE_DIR)
        import sector_comtrade_pipeline as scp
        original_cwd = os.getcwd()
        os.chdir(COMTRADE_EXPORTS_DIR)
        try:
            ref = scp.get_country_reference()
        finally:
            os.chdir(original_cwd)
        return ref

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

    import datetime as _dt
    _auto_latest = _dt.date.today().year - 1
    try:
        config_path_full = os.path.join(COMTRADE_EXPORTS_DIR, config_file_path) if not os.path.isabs(config_file_path) else config_file_path
        dfy = pd.read_excel(config_path_full, "years", header=1)
        dfy.columns = [str(c).strip() for c in dfy.columns]
        cfg_start = int(float(dfy.iloc[0]["Start Year"]))
        _ey = dfy.iloc[0].get("End Year", None)
        cfg_end = (int(float(_ey)) if pd.notna(_ey) and str(_ey).strip().lower() not in ("", "auto", "latest", "nan")
                  else _auto_latest)
    except Exception:
        cfg_start, cfg_end = _auto_latest - 10, _auto_latest
    
    slider_max = max(cfg_end, _auto_latest)

    try:
        config_path_full = os.path.join(COMTRADE_EXPORTS_DIR, config_file_path) if not os.path.isabs(config_file_path) else config_file_path
        dfh = pd.read_excel(config_path_full, "hs_codes", header=1, dtype=str)
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
                script_path = os.path.join(COMTRADE_DIR, "sector_comtrade_pipeline.py")
                cmd = [sys.executable, script_path, "--config", config_file_path,
                      "--mode", "top_partners", "--reporter", reporter_code, "--top-n", str(tp_top_n)]
                if active_hs:
                    cmd += ["--hs", ",".join(active_hs)]
                if tp_years:
                    cmd += ["--years", ",".join(tp_years)]
                _run_top_partners_fetch(cmd, f"Finding {reporter_name}'s top trading partners...")

        tp_slug = get_config_field(os.path.join(COMTRADE_EXPORTS_DIR, config_file_path) if not os.path.isabs(config_file_path) else config_file_path, "sector_details", "Sector slug")
        tp_json_path = os.path.join(COMTRADE_EXPORTS_DIR, f"{tp_slug}_top_partners.json") if tp_slug else None
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
                script_path = os.path.join(COMTRADE_DIR, "sector_comtrade_pipeline.py")
                cmd = [sys.executable, script_path, "--config", config_file_path,
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
                script_path = os.path.join(COMTRADE_DIR, "sector_comtrade_pipeline.py")
                cmd = [sys.executable, script_path, "--config", config_file_path,
                      "--mode", "bilateral_compare", "--reporter", reporter_code,
                      "--partners", ",".join(partner_codes), "--flow", flow_code]
                if hs_arg:
                    cmd += ["--hs", ",".join(hs_arg)]
                if years_arg:
                    cmd += ["--years", ",".join(years_arg)]
                if not include_world:
                    cmd += ["--no-world-total"]
                _run_geo_fetch(cmd, f"Fetching {reporter_name}'s trade with {len(partner_names)} partner(s)...")

    geo_cleaned = cleaned_path or locate_cleaned_workbook(os.path.join(COMTRADE_EXPORTS_DIR, config_file_path) if not os.path.isabs(config_file_path) else config_file_path)
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

def run_comtrade_app():
    """Comtrade Sector Pipeline — auto-generate a sector config, run the
    download/clean/competitor/visuals pipeline, then explore results as
    interactive charts. All generated config/output files live under
    COMTRADE_EXPORTS_DIR so they don't collide with the other modules in
    this hub."""

    def _cfg_full_path(p):
        """Resolve a (possibly relative) config filename to its full path
        inside COMTRADE_EXPORTS_DIR, for direct pandas file IO. Subprocess
        calls instead pass the relative filename + cwd=COMTRADE_EXPORTS_DIR,
        matching the convention already used by the helper functions above."""
        return os.path.join(COMTRADE_EXPORTS_DIR, p) if p and not os.path.isabs(p) else p

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
            sector_choice = st.selectbox("Choose a Sector Template", list(SECTOR_PRESETS_COMTRADE.keys()))

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
                groups = SECTOR_PRESETS_COMTRADE[sector_choice]
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
                    cmd = [sys.executable, os.path.join(COMTRADE_DIR, "create_sector_config_from_qe.py")]
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
                            env=SUBPROCESS_ENV, encoding="utf-8", errors="replace", cwd=COMTRADE_EXPORTS_DIR,
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

        xls = pd.ExcelFile(_cfg_full_path(st.session_state.config_file_path))
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
                with pd.ExcelWriter(_cfg_full_path(st.session_state.config_file_path), engine="openpyxl") as writer:
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
            _dfy3 = pd.read_excel(_cfg_full_path(st.session_state.config_file_path), "years", header=1)
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
                        cmd = [sys.executable, os.path.join(COMTRADE_DIR, "sector_comtrade_pipeline.py"), "--config", st.session_state.config_file_path, "--mode", current_mode]
                        if year_override:
                            cmd += ["--start-year", str(year_override[0]), "--end-year", str(year_override[1])]

                        try:
                            process = subprocess.Popen(
                                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                                env=SUBPROCESS_ENV, encoding="utf-8", errors="replace", cwd=COMTRADE_EXPORTS_DIR,
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
                        cleaned_check = locate_cleaned_workbook(_cfg_full_path(st.session_state.config_file_path))
                        st.markdown("**Cleaned workbook diagnostics:**")
                        for level, msg in diagnose_cleaned_workbook(cleaned_check):
                            {"error": st.error, "warning": st.warning, "success": st.success}[level](msg)

        # Interactive Visuals Section: shows live, filterable charts built from
        # whatever sector the user is currently working with — no deck involved.
        chart_data_path = locate_chart_data_workbook(_cfg_full_path(st.session_state.config_file_path))
        if chart_data_path:
            st.divider()
            show_interactive_visuals(chart_data_path, st.session_state.sector_name or "Sector")

        # Master_Table explorer: free-form filtering (year range, trade flow,
        # reporter/competitor countries, sector bucket) over the cleaned workbook,
        # independent of the fixed Chart1-8 set above.
        cleaned_path_for_explore = locate_cleaned_workbook(_cfg_full_path(st.session_state.config_file_path))
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
            current_dir = COMTRADE_EXPORTS_DIR

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

# ==============================================================================
# 2. INDIVIDUAL APP MODULES
# ==============================================================================

def run_asi_app():
    st.title("🏭 Annual Survey of Industries (ASI)")

    available_options = [f"{code} - {ASI_NIC_NAMES[code]}" for code in sorted(ASI_NIC_NAMES.keys())]

    st.sidebar.header("🎯 Target Segment Filters")

    if "selected_nic_items" not in st.session_state: st.session_state.selected_nic_items = ["13 - Textiles"]
    if "custom_sector_name" not in st.session_state: st.session_state.custom_sector_name = "Textiles"
    if "sector_preset" not in st.session_state: st.session_state.sector_preset = "Textiles"

    def on_preset_change():
        preset = st.session_state.sector_preset
        if preset != "Custom Sector Entry":
            default_code = PRESET_MAP[preset][0]
            st.session_state.selected_nic_items = [f"{default_code} - {ASI_NIC_NAMES[default_code]}"]
            st.session_state.custom_sector_name = preset

    sector_preset = st.sidebar.selectbox(
        "Select Core Sector Template:",
        options=list(PRESET_MAP.keys()) + ["Custom Sector Entry"],
        index=3, key="sector_preset", on_change=on_preset_change
    )

    sector_name = st.sidebar.text_input("Analysis Label Name:", key="custom_sector_name")
    selected_friendly = st.sidebar.multiselect("Active Grouping NIC Codes:", options=available_options, key="selected_nic_items")
    nic_codes_list = [item.split(" - ")[0].strip() for item in selected_friendly]

    st.sidebar.markdown("---")
    st.sidebar.header("🎛️ Micro-Filtering Dashboard")

    with st.sidebar.expander("📍 Regional & Geographic Splits", expanded=False):
        region_filter = st.radio("Region Profile:", ["All Regions", "Rural Only", "Urban Only"])
        state_selection = st.multiselect("Isolate States (Leave empty for All India):", options=[f"{k} - {v}" for k, v in ASI_STATE_NAMES.items()])
        selected_state_codes = [s.split(" - ")[0] for s in state_selection]

    with st.sidebar.expander("🏭 Operational Parameters", expanded=False):
        scheme_filter = st.radio("Survey Tracking Scheme:", ["All Schemes", "Census Only (Large-Scale)", "Sample Only (Mid-Scale)"])
        exclude_zero_mult = st.checkbox("Drop Zero Multipliers (mult = 0)", value=True)
        exclude_zero_output = st.checkbox("Drop Zero Gross Output (costop = 0)", value=True)
        exclude_zero_workers = st.checkbox("Drop Zero Worker Records", value=True)

    st.sidebar.markdown("---")
    st.sidebar.header("📁 Advanced Dictionary Mapping")
    nic_mapping_file = st.sidebar.file_uploader("Optional: 5-Digit NIC Dictionary (CSV)", type=["csv"], help="Upload a CSV with 5-digit codes in the first column and industry names in the second column to unlock exact labels.")

    granular_nic_map = {}
    if nic_mapping_file:
        try:
            mapping_df = pd.read_csv(nic_mapping_file, dtype=str)
            code_col, name_col = mapping_df.columns[0], mapping_df.columns[1]
            granular_nic_map = dict(zip(mapping_df[code_col].str.strip(), mapping_df[name_col].str.strip()))
            st.sidebar.success(f"✅ Successfully loaded {len(granular_nic_map):,} exact industry names!")
        except Exception as e:
            st.sidebar.error("⚠️ Could not read dictionary. Make sure it's a valid CSV.")

    # Extended and comprehensive variable mapping dictionary
    ASI_VAR_LABELS = {
        "yr": "'24' for ASI 2023-2024", "blk": "Block code", "a1": "DSL", "a2": "PSL No.", 
        "a3": "Scheme code (Census-1, Sample-2)", "a4": "Ind. Code as per Frame (4-digit NIC)", 
        "a5": "Ind Code as per Return (5-digit NIC)", "a7": "State Code", "a8": "District code", 
        "a9": "Sector (Rural-1, Urban-2)", "a10": "RO/SRO code", "a11": "No. of units", 
        "a12": "Status of Unit", "bonus": "Bonus", "pf": "Provident Fund", 
        "welfare": "Welfare Expenses", "mwdays": "Number of working days (Manufacturing)", 
        "nwdays": "Number of working days (Non-Manufacturing)", "wdays": "Number of working days (Total)", 
        "costop": "Total Cost of Production", "expshare": "Share(%) of products/by-products directly exported", 
        "mult": "Multiplier", "ab01": "DSL (Block-A, Item 1)", "b02": "Type of organisation(code)", 
        "b03": "Corporate Identification Number (CIN)", "b04": "ISO Certification, 14000 Series", 
        "b05": "Year of initial production", "b06f": "Accounting year (From)", "b06t": "Accounting year (To)", 
        "b07": "Number of months of operation", "b08": "Share capital includes foreign entities?", 
        "b09": "R&D unit in factory?", "b11": "Formal training offered", "ac01": "DSL (Block-A, Item 1)", 
        "c 11": "SL. No.", "c 13": "Gross Value Opening as on", "c_14": "Gross Value of Addition due to Revaluation", 
        "c 15": "Gross Value of Actual addition", "c 16": "Gross Value of Deduction & adjustment", 
        "c 17": "Gross Value Closing as on", "c 18": "Depreciation Up to year beginning", 
        "c 19": "Depreciation Provided during the year", "c 110": "Depreciation due to Adjustment for sold/discarded", 
        "c 111": "Depreciation Up to year end", "c_112": "Net Value Opening as on", "c 113": "Net Value Closing as on",
        "ad01": "DSL (Block-A, Item 1)", "dii": "SL. No.", "d 13": "Opening (Rs.)", "d 14": "Closing (Rs.)",
        "ae01": "DSL (Block-A, Item 1)", "e 11": "SL. No.", "e 13": "Mandays Worked (Manufacturing)", 
        "e 14": "Mandays Worked (Non Manufacturing)", "e 15": "Mandays Worked (Total)", 
        "e 16": "Average Number of persons worked", "e 17": "No. of mandays paid for", "e 18": "Wages/salaries (in Rs.)",
        
        "f1": "Work done by others", "f2a": "Repair & maintenance (Building)", "f2b": "Repair & maintenance (Other)",
        "f3": "Operating expenses", "f4": "Own construction expenses", "f5": "Insurance Charges", 
        "f6": "Rent paid for Plant & Machinery", "f7": "R&D Expenses", "f8": "Rent paid for Buildings", 
        "f9": "Rent paid for land / royalties", "f10": "Interest paid", "f11": "Purchase value of goods resold", 
        "f12": "Inward transportation cost", "f13": "Outward transportation cost",
        
        "g1": "Receipts from mfg services", "g2": "Receipts from non-mfg services", "g3": "Value in electricity generated and sold",
        "g4": "Value of own construction", "g5": "Net balance of goods resold", "g6": "Rent received for Plant & Machinery",
        "g7": "Variation in stock of semi-finished goods", "g8": "Rent received for buildings", "g9": "Rent received for land / royalties",
        "g10": "Interest received", "g11": "Sale value of goods resold", "g12": "Other production subsidies",
        
        "h12": "Basic items consumed (Value)", "h13": "Non-basic chemicals consumed (Value)", "h14": "Packing items consumed (Value)",
        "h15": "Electricity own generated (Value)", "h16": "Electricity purchased (Value)", "h17": "Petrol, diesel, oil, lubricants (Value)",
        "h18": "Coal consumed (Value)", "h19": "Gas consumed (Value)", "h20": "Other fuel consumed (Value)", "h21": "Consumable store (Value)",
        
        "j13": "Item code (NPCMS) - Products", "j15": "Quantity manufactured", "j17": "Gross sale value (Rs.)",
        
        "imported_total_inputs": "Total Imported Inputs", "indigenous_total_inputs": "Total Indigenous Inputs",
        "depreciation_annexure": "Depreciation (Annexure VIII)", "nfcf_without_f7": "Net Fixed Capital Formation (Excl. R&D)",
        "materials_fuels_stores_opening": "Opening Stock: Materials, Fuels, Stores", "materials_fuels_stores_closing": "Closing Stock: Materials, Fuels, Stores",
        "semi_finished_opening": "Opening Stock: Semi-finished", "semi_finished_closing": "Closing Stock: Semi-finished",
        "finished_goods_opening": "Opening Stock: Finished Goods", "finished_goods_closing": "Closing Stock: Finished Goods",
        "total_employee_wages": "Total Employee Wages", "total_workers": "Total Workers", "women_workers": "Women Workers",
        "fixed_assets_net_closing": "Fixed Assets: Net Closing Value", "plant_machinery_gross_closing": "Plant & Machinery: Gross Closing Value",
        "actual_addition_fixed_assets": "Actual Addition to Fixed Assets", "working_capital": "Working Capital (Block D)",
        "total_inventory_closing": "Total Inventory (Closing)", "outstanding_loans": "Outstanding Loans (Block D)",
        "gross_sale_value_item12": "Gross Sale Value (Item 12)", "gst_item12": "GST (Item 12)",
        "excise_vat_other_taxes_item12": "Excise/VAT/Other Taxes (Item 12)", "other_distributive_expenses_item12": "Other Distributive Expenses (Item 12)",
        "subsidy_item12": "Subsidy (Item 12)", "ex_factory_value_item12": "Ex-factory Value (Item 12)",
        
        # Principal Characteristics
        "pc_01_number_of_factories": "PC 01: Number of Factories",
        "pc_02_factories_in_operation": "PC 02: Factories in Operation",
        "pc_03_fixed_capital": "PC 03: Fixed Capital",
        "pc_04_physical_working_capital": "PC 04: Physical Working Capital",
        "pc_05_working_capital": "PC 05: Working Capital",
        "pc_06_invested_capital": "PC 06: Invested Capital",
        "pc_07_gross_value_addition_fixed_capital": "PC 07: Gross Value of Addition to Fixed Capital",
        "pc_08_rent_paid_fixed_assets": "PC 08: Rent Paid for Fixed Assets",
        "pc_09_outstanding_loan": "PC 09: Outstanding Loan",
        "pc_10_interest_paid": "PC 10: Interest Paid",
        "pc_11_rent_received_fixed_assets": "PC 11: Rent Received for Fixed Assets",
        "pc_12_interest_received": "PC 12: Interest Received",
        "pc_13_gross_value_plant_machinery": "PC 13: Gross Value of Plant & Machinery",
        "pc_14_value_product_byproduct": "PC 14: Value of Product and By-Product",
        "pc_15_total_output": "PC 15: Total Output",
        "pc_16_fuels_consumed": "PC 16: Fuels Consumed",
        "pc_17_materials_consumed": "PC 17: Materials Consumed",
        "pc_18_total_inputs": "PC 18: Total Inputs",
        "pc_19_gross_value_added": "PC 19: Gross Value Added",
        "pc_20_depreciation": "PC 20: Depreciation",
        "pc_21_net_value_added": "PC 21: Net Value Added",
        "pc_22_net_fixed_capital_formation": "PC 22: Net Fixed Capital Formation",
        "pc_23_gross_fixed_capital_formation": "PC 23: Gross Fixed Capital Formation",
        "pc_24_addition_in_stock": "PC 24: Addition in Stock",
        "pc_25_gross_capital_formation": "PC 25: Gross Capital Formation",
        "pc_26_net_income": "PC 26: Net Income",
        "pc_27_net_profit": "PC 27: Net Profit"
    }

    def _norm_df(df, id_col):
        df.columns = df.columns.str.lower().str.strip()
        if id_col.lower() in df.columns:
            df = df.rename(columns={id_col.lower(): 'dsl'})
        elif 'a1' in df.columns:
            df = df.rename(columns={'a1': 'dsl'})
        df['dsl'] = df['dsl'].astype(str)
        return df

    st.markdown("### 📥 Primary Data Ingestion (All 10 Blocks Required)")
    
    st.info("Upload all 10 standard ASI blocks to unlock the 27 Principal Characteristics auto-computation engine.")
    
    c1, c2, c3, c4, c5 = st.columns(5)
    file_a = c1.file_uploader("Block A", type=["csv"])
    file_b = c2.file_uploader("Block B", type=["csv"])
    file_c = c3.file_uploader("Block C", type=["csv"])
    file_d = c4.file_uploader("Block D", type=["csv"])
    file_e = c5.file_uploader("Block E", type=["csv"])
    
    c6, c7, c8, c9, c10 = st.columns(5)
    file_f = c6.file_uploader("Block F", type=["csv"])
    file_g = c7.file_uploader("Block G", type=["csv"])
    file_h = c8.file_uploader("Block H", type=["csv"])
    file_i = c9.file_uploader("Block I", type=["csv"])
    file_j = c10.file_uploader("Block J", type=["csv"])

    if all([file_a, file_b, file_c, file_d, file_e, file_f, file_g, file_h, file_i, file_j]):
        if not nic_codes_list:
            st.error("❌ Configure at least one active NIC parameter in the control console to initialize calculations.")
        else:
            with st.spinner("Compiling database tables, syncing weights, and computing the 27 Principal Characteristics..."):
                # 1. Base Factory (Block A & B)
                df_A = _norm_df(pd.read_csv(file_a, dtype={"a1": str, "a5": str, "a7": str}), 'a1')
                if "a5" in df_A.columns: df_A["a5"] = df_A["a5"].astype(str).str.strip().str.zfill(5)
                if "a7" in df_A.columns: df_A["a7"] = df_A["a7"].astype(str).str.strip().str.zfill(2)
                
                df_B = _norm_df(pd.read_csv(file_b), 'ab01')
                
                # 2. Block C Processing
                df_C = _norm_df(pd.read_csv(file_c), 'ac01')
                for c in ['c_11', 'c_19', 'c_112', 'c_113', 'c_14', 'c_15', 'c_17']:
                    if c in df_C.columns: df_C[c] = pd.to_numeric(df_C[c], errors='coerce').fillna(0)
                    else: df_C[c] = 0

                annex_mask = df_C['c_11'].isin([1,2,3,4,5,6,7,9])
                c_annex = df_C[annex_mask].groupby('dsl').apply(lambda x: pd.Series({
                    'depreciation_annexure': x['c_19'].sum(),
                    'nfcf_without_f7': (x['c_113'] - x['c_112'] - x['c_14']).sum()
                })).reset_index()
                
                c_10 = df_C[df_C['c_11'] == 10].groupby('dsl').agg(
                    fixed_assets_net_closing=('c_113', 'first'),
                    actual_addition_fixed_assets=('c_15', 'first')
                ).reset_index()

                c_3 = df_C[df_C['c_11'] == 3].groupby('dsl').agg(
                    plant_machinery_gross_closing=('c_17', 'first')
                ).reset_index()

                df_C_merged = c_annex.merge(c_10, on='dsl', how='outer').merge(c_3, on='dsl', how='outer')

                # 3. Block D Processing
                df_D = _norm_df(pd.read_csv(file_d), 'ad01')
                for c in ['di1', 'di3', 'di4', 'd 11', 'd 13', 'd 14']:
                    if c in df_D.columns: df_D[c] = pd.to_numeric(df_D[c], errors='coerce')
                
                d_id = 'di1' if 'di1' in df_D.columns else 'd 11' if 'd 11' in df_D.columns else None
                d_op = 'di3' if 'di3' in df_D.columns else 'd 13' if 'd 13' in df_D.columns else None
                d_cl = 'di4' if 'di4' in df_D.columns else 'd 14' if 'd 14' in df_D.columns else None
                
                if all([d_id, d_op, d_cl]):
                    def ext_d(item_id, col_name, out_name):
                        return df_D[df_D[d_id] == item_id].groupby('dsl')[col_name].first().rename(out_name)
                    df_D_merged = pd.concat([
                        ext_d(4, d_op, 'materials_fuels_stores_opening'), ext_d(4, d_cl, 'materials_fuels_stores_closing'),
                        ext_d(5, d_op, 'semi_finished_opening'), ext_d(5, d_cl, 'semi_finished_closing'),
                        ext_d(6, d_op, 'finished_goods_opening'), ext_d(6, d_cl, 'finished_goods_closing'),
                        ext_d(7, d_cl, 'total_inventory_closing'), ext_d(16, d_cl, 'working_capital'),
                        ext_d(17, d_cl, 'outstanding_loans')
                    ], axis=1).reset_index()
                else: df_D_merged = pd.DataFrame(columns=['dsl'])

                # 4. Block E Processing
                df_E = _norm_df(pd.read_csv(file_e), 'ae01')
                e_id = 'ei1' if 'ei1' in df_E.columns else 'e 11'
                e_wages = 'ei8' if 'ei8' in df_E.columns else 'e 18'
                e_workers = 'ei6' if 'ei6' in df_E.columns else 'e 16'
                for c in [e_id, e_wages, e_workers]:
                    if c in df_E.columns: df_E[c] = pd.to_numeric(df_E[c], errors='coerce')
                
                if e_id in df_E.columns:
                    e10 = df_E[df_E[e_id] == 10].groupby('dsl')[e_wages].first().rename('total_employee_wages')
                    e6_tot = df_E[df_E[e_id] == 6].groupby('dsl')[e_workers].first().rename('total_workers')
                    e2_wom = df_E[df_E[e_id] == 2].groupby('dsl')[e_workers].first().rename('women_workers')
                    df_E_merged = pd.concat([e10, e6_tot, e2_wom], axis=1).reset_index()
                else: df_E_merged = pd.DataFrame(columns=['dsl'])

                # 5. Block F & G Processing
                df_F = _norm_df(pd.read_csv(file_f), 'af01')
                df_G = _norm_df(pd.read_csv(file_g), 'ag01')
                df_F_merged = df_F.groupby('dsl').sum(numeric_only=True).reset_index()
                df_G_merged = df_G.groupby('dsl').sum(numeric_only=True).reset_index()

                # 6. Block H Processing
                df_H = _norm_df(pd.read_csv(file_h), 'ah01')
                h_id = 'hi1' if 'hi1' in df_H.columns else 'h 11'
                h_val = 'hi6' if 'hi6' in df_H.columns else 'h 16'
                for c in [h_id, h_val]:
                    if c in df_H.columns: df_H[c] = pd.to_numeric(df_H[c], errors='coerce')
                if h_id in df_H.columns:
                    def ext_h(i, n): return df_H[df_H[h_id] == i].groupby('dsl')[h_val].first().rename(n)
                    df_H_merged = pd.concat([
                        ext_h(15, 'h15'), ext_h(16, 'h16'), ext_h(17, 'h17'), ext_h(18, 'h18'), 
                        ext_h(19, 'h19'), ext_h(20, 'h20'), ext_h(12, 'h12'), ext_h(13, 'h13'), 
                        ext_h(14, 'h14'), ext_h(21, 'h21'), ext_h(23, 'indigenous_total_inputs')
                    ], axis=1).reset_index()
                else: df_H_merged = pd.DataFrame(columns=['dsl'])

                # 7. Block I Processing
                df_I = _norm_df(pd.read_csv(file_i), 'ai01')
                i_id = 'ii1' if 'ii1' in df_I.columns else 'i 11'
                i_val = 'ii6' if 'ii6' in df_I.columns else 'i 16'
                for c in [i_id, i_val]: 
                    if c in df_I.columns: df_I[c] = pd.to_numeric(df_I[c], errors='coerce')
                if i_id in df_I.columns:
                    df_I_merged = df_I[df_I[i_id] == 7].groupby('dsl')[i_val].first().rename('imported_total_inputs').reset_index()
                else: df_I_merged = pd.DataFrame(columns=['dsl'])

                # 8. Block J Processing
                df_J = _norm_df(pd.read_csv(file_j), 'aj01')
                j_id = 'j11' if 'j11' in df_J.columns else 'j 11'
                
                for c in [j_id, 'j17', 'j18', 'j19', 'j110', 'j111', 'j113', 'j 17', 'j 18', 'j 19', 'j 110', 'j 111', 'j 113']:
                    if c in df_J.columns: df_J[c] = pd.to_numeric(df_J[c], errors='coerce')
                        
                j17_col = 'j17' if 'j17' in df_J.columns else 'j 17' if 'j 17' in df_J.columns else None
                j18_col = 'j18' if 'j18' in df_J.columns else 'j 18' if 'j 18' in df_J.columns else None
                j19_col = 'j19' if 'j19' in df_J.columns else 'j 19' if 'j 19' in df_J.columns else None
                j110_col = 'j110' if 'j110' in df_J.columns else 'j 110' if 'j 110' in df_J.columns else None
                j111_col = 'j111' if 'j111' in df_J.columns else 'j 111' if 'j 111' in df_J.columns else None
                j113_col = 'j113' if 'j113' in df_J.columns else 'j 113' if 'j 113' in df_J.columns else None

                if j_id in df_J.columns and j17_col and j18_col and j113_col:
                    agg_dict = {
                        'gross_sale_value_item12': (j17_col, 'first'),
                        'gst_item12': (j18_col, 'first'),
                        'ex_factory_value_item12': (j113_col, 'first')
                    }
                    if j19_col: agg_dict['excise_vat_other_taxes_item12'] = (j19_col, 'first')
                    if j110_col: agg_dict['other_distributive_expenses_item12'] = (j110_col, 'first')
                    if j111_col: agg_dict['subsidy_item12'] = (j111_col, 'first')

                    df_J_merged = df_J[df_J[j_id] == 12].groupby('dsl').agg(**agg_dict).reset_index()
                else: df_J_merged = pd.DataFrame(columns=['dsl'])

                # Merging all Blocks
                merged = df_A.copy()
                for d in [df_B, df_C_merged, df_D_merged, df_E_merged, df_F_merged, df_G_merged, df_H_merged, df_I_merged, df_J_merged]:
                    if not d.empty: 
                        overlap_cols = [col for col in d.columns if col in merged.columns and col != 'dsl']
                        d_clean = d.drop(columns=overlap_cols)
                        merged = merged.merge(d_clean, on='dsl', how='left')
                
                # Fill NAs
                num_cols_m = merged.select_dtypes(include=[np.number]).columns
                merged[num_cols_m] = merged[num_cols_m].fillna(0)

                # Ensure required columns for PCs exist
                required_cols = [
                    'a11', 'a12', 'fixed_assets_net_closing', 'total_inventory_closing', 'working_capital',
                    'actual_addition_fixed_assets', 'f9', 'outstanding_loans', 'f10', 'g9', 'g10',
                    'plant_machinery_gross_closing', 'gross_sale_value_item12', 'gst_item12', 'ex_factory_value_item12',
                    'excise_vat_other_taxes_item12', 'other_distributive_expenses_item12', 'subsidy_item12', 
                    'g1', 'g2', 'g3', 'g4', 'g6', 'g7', 'g8', 'g11', 'f7',
                    'h15', 'h16', 'h17', 'h18', 'h19', 'h20', 'h12', 'h13', 'h14', 'h21', 'imported_total_inputs',
                    'f1', 'f2a', 'f2b', 'f3', 'f4', 'f6', 'f8', 'f11', 'indigenous_total_inputs',
                    'depreciation_annexure', 'nfcf_without_f7', 'materials_fuels_stores_closing',
                    'materials_fuels_stores_opening', 'semi_finished_closing', 'semi_finished_opening',
                    'finished_goods_closing', 'finished_goods_opening', 'total_employee_wages',
                    'bonus', 'pf', 'welfare', 'total_workers', 'women_workers', 'costop', 'mult'
                ]
                for c in required_cols:
                    if c not in merged.columns: merged[c] = 0

                # ----------------------------------------------------
                # APPLY 27 PRINCIPAL CHARACTERISTICS FORMULAS
                # ----------------------------------------------------
                merged['pc_01_number_of_factories'] = merged['a11']
                merged['pc_02_factories_in_operation'] = np.where(merged['a12'].isin([1, 2, 3]), merged['a11'], 0)
                merged['pc_03_fixed_capital'] = merged['fixed_assets_net_closing']
                merged['pc_04_physical_working_capital'] = merged['total_inventory_closing']
                merged['pc_05_working_capital'] = merged['working_capital']
                merged['pc_06_invested_capital'] = merged['pc_03_fixed_capital'] + merged['pc_04_physical_working_capital']
                merged['pc_07_gross_value_addition_fixed_capital'] = merged['actual_addition_fixed_assets']
                merged['pc_08_rent_paid_fixed_assets'] = merged['f9']
                merged['pc_09_outstanding_loan'] = merged['outstanding_loans']
                merged['pc_10_interest_paid'] = merged['f10']
                merged['pc_11_rent_received_fixed_assets'] = merged['g9']
                merged['pc_12_interest_received'] = merged['g10']
                merged['pc_13_gross_value_plant_machinery'] = merged['plant_machinery_gross_closing']
                
                merged['pc_14_value_product_byproduct'] = merged['gross_sale_value_item12'] - (
                    merged['gst_item12'] + 
                    merged['excise_vat_other_taxes_item12'] + 
                    merged['other_distributive_expenses_item12'] - 
                    merged['subsidy_item12']
                )
                
                merged['pc_15_total_output'] = (
                    merged['ex_factory_value_item12'] + merged['g1'] + merged['g2'] + merged['g3'] +
                    merged['g4'] + merged['g6'] + merged['g7'] + merged['g8'] + merged['g11'] + merged['f7']
                )
                merged['pc_16_fuels_consumed'] = (
                    merged['h15'] + merged['h16'] + merged['h17'] + merged['h18'] + merged['h19'] + merged['h20']
                )
                merged['pc_17_materials_consumed'] = (
                    merged['h12'] + merged['h13'] + merged['h14'] + merged['h21'] + merged['imported_total_inputs']
                )
                merged['pc_18_total_inputs'] = (
                    merged['f1'] + merged['f2a'] + merged['f2b'] + merged['f3'] + merged['f4'] +
                    merged['f6'] + merged['f7'] + merged['f8'] + merged['f11'] +
                    merged['indigenous_total_inputs'] + merged['imported_total_inputs']
                )
                merged['pc_19_gross_value_added'] = merged['pc_15_total_output'] - merged['pc_18_total_inputs']
                merged['pc_20_depreciation'] = merged['depreciation_annexure']
                merged['pc_21_net_value_added'] = merged['pc_19_gross_value_added'] - merged['pc_20_depreciation']
                merged['pc_22_net_fixed_capital_formation'] = merged['nfcf_without_f7'] + merged['f7']
                merged['pc_23_gross_fixed_capital_formation'] = merged['pc_22_net_fixed_capital_formation'] + merged['pc_20_depreciation']
                merged['pc_24a_addition_stock_materials_fuels'] = merged['materials_fuels_stores_closing'] - merged['materials_fuels_stores_opening']
                merged['pc_24b_addition_stock_semi_finished'] = merged['semi_finished_closing'] - merged['semi_finished_opening']
                merged['pc_24c_addition_stock_finished_goods'] = merged['finished_goods_closing'] - merged['finished_goods_opening']
                merged['pc_24_addition_in_stock'] = (
                    merged['pc_24a_addition_stock_materials_fuels'] + merged['pc_24b_addition_stock_semi_finished'] + merged['pc_24c_addition_stock_finished_goods']
                )
                merged['pc_25_gross_capital_formation'] = merged['pc_23_gross_fixed_capital_formation'] + merged['pc_24_addition_in_stock']
                merged['pc_26_net_income'] = merged['pc_21_net_value_added'] - merged['f9'] - merged['f10']
                merged['pc_27_net_profit'] = merged['pc_26_net_income'] - merged['total_employee_wages'] - merged['bonus'] - merged['pf'] - merged['welfare']

                # PREVENT PERFORMANCE WARNING: DE-FRAGMENT MEMORY
                merged = merged.copy()

                # Pre-processing baseline filtering arrays
                merged["division"] = merged["a5"].astype(str).str[:2] if "a5" in merged.columns else "Unknown"
                merged["total_emoluments"] = merged["total_employee_wages"].fillna(0) + merged["bonus"].fillna(0) + merged["pf"].fillna(0) + merged["welfare"].fillna(0)
                
                filter_reasons = []
                if exclude_zero_mult: merged = merged[merged["mult"] > 0]
                if exclude_zero_output: merged = merged[merged["costop"] > 0] 
                if exclude_zero_workers: merged = merged[merged["total_workers"] > 0]
                    
                if region_filter == "Rural Only" and "a9" in merged.columns:
                    merged = merged[merged["a9"] == 1]
                    filter_reasons.append("Rural Facilities Only")
                elif region_filter == "Urban Only" and "a9" in merged.columns:
                    merged = merged[merged["a9"] == 2]
                    filter_reasons.append("Urban Facilities Only")
                    
                if selected_state_codes and "a7" in merged.columns:
                    merged = merged[merged["a7"].astype(str).isin(selected_state_codes)]
                    filter_reasons.append(f"Isolating {len(selected_state_codes)} State Entities")
                    
                if scheme_filter == "Census Only (Large-Scale)" and "a3" in merged.columns:
                    merged = merged[merged["a3"] == 1]
                    filter_reasons.append("Census Cohort")
                elif scheme_filter == "Sample Only (Mid-Scale)" and "a3" in merged.columns:
                    merged = merged[merged["a3"] == 2]
                    filter_reasons.append("Sample Representative Cohort")
                    
                st.success(f"⚡ **Ingestion Engine Connected:** `{merged.shape[0]:,}` operational rows matching active filter matrix.")
                
                if "a5" in merged.columns:
                    sector_mask = merged["a5"].astype(str).str.startswith(tuple(nic_codes_list))
                    sector_df = merged[sector_mask].copy()
                    other_df = merged[~sector_mask].copy()
                else:
                    sector_df = pd.DataFrame()
                    other_df = merged.copy()
                    
                full_df = merged.copy()
                def pop_estimate(df, col): return (df[col] * df["mult"]).sum() if col in df.columns else 0
                
                base_exclude = ["a3", "a9", "mult", "a7", "a5", "dsl", "year"]
                extra_metric_cols = [c for c in full_df.select_dtypes(include=[np.number]).columns if c not in base_exclude]
                
                if not sector_df.empty:
                    summary_rows = []
                    for label, curr_df in [(f"Sector: {sector_name}", sector_df), ("All Manufacturing", full_df), ("Other Industrial Baseline", other_df)]:
                        if curr_df.shape[0] == 0: continue
                        w_wk = pop_estimate(curr_df, "total_workers")
                        w_op = pop_estimate(curr_df, "pc_15_total_output")
                        w_em = pop_estimate(curr_df, "total_emoluments")
                        w_wom = pop_estimate(curr_df, "women_workers")
                        w_fixed = pop_estimate(curr_df, "pc_03_fixed_capital") 
                        summary_rows.append({
                            "Group Segment": label, "Surveyed Plants": curr_df.shape[0], "Est Population Size": int(round(curr_df["mult"].sum())),
                            "Output per Worker (Lakh)": (w_op / w_wk / 1e5) if w_wk > 0 else 0,
                            "Emp per Crore Investment": (w_wk / (w_fixed / 1e7)) if w_fixed > 0 else 0,
                            "Labour Income (%)": (w_em / w_op * 100) if w_op > 0 else 0,
                            "Women Workforce Share (%)": (w_wom / w_wk * 100) if w_wk > 0 else 0,
                            "Workforce Size (Mn)": w_wk / 1e6, "Capital Stock Asset (Cr)": w_fixed / 1e7
                        })
                    summary_table_df = pd.DataFrame(summary_rows)
                    
                    st.markdown("## 📊 Custom Visualization Studio & Matrix Builder")
                    studio_col1, studio_col2 = st.columns([1, 3])
                    
                    # ----------------------------------------------------
                    # UI DROPDOWN LIST ASSEMBLY
                    # ----------------------------------------------------
                    base_metrics = [
                        "Emp per Crore Investment", "Output per Worker (Lakh)", "Labour Income (%)", 
                        "Women Workforce Share (%)", "Plant Density Count", "Total Segment Workforce"
                    ]
                    
                    pc_options = []
                    other_options = []
                    
                    for c in extra_metric_cols:
                        if c.startswith('pc_'):
                            pc_name = ASI_VAR_LABELS.get(c.lower(), c.upper().replace('_', ' '))
                            pc_options.append(pc_name)
                        else:
                            label_key = c.split('_')[-1].lower() if c.startswith(('f_', 'g_', 'h_')) else c.lower()
                            dict_label = ASI_VAR_LABELS.get(label_key, ASI_VAR_LABELS.get(c.lower(), f"Raw Data: {c.upper()}"))
                            other_options.append(f"{dict_label} ({c})")
                            
                    pc_options.sort()
                    other_options.sort()
                    
                    # PCs -> Base Multi-variable formulas -> Everything else
                    metric_options = pc_options + base_metrics + other_options
                    
                    with studio_col1:
                        st.markdown("##### ⚙️ Layout Configuration")
                        chart_dimension = st.selectbox("Data Grouping (Categories):", options=["2-Digit NIC Division", "5-Digit Granular NIC", "Geographical State Location", "Rural vs Urban Region"])
                        chart_metric = st.selectbox("Primary Measurement (Y-Axis/Values):", options=metric_options)
                        chart_type = st.selectbox("Visualization Style:", ["Bar Chart", "Line Chart", "Scatter Plot", "Pie Chart"])
                        
                        scatter_x = None
                        if chart_type == "Scatter Plot": scatter_x = st.selectbox("Scatter X-Axis Metric:", options=metric_options, index=1)
                            
                        chart_palette = st.color_picker("Primary Analytical Accent Color:", "#1B3A5C")
                        min_sample_threshold = st.slider("Minimum Sample Cut-off (Plants per Category):", min_value=1, max_value=200, value=25)
                        max_display_items = st.slider("Max Chart Items (Top N):", min_value=5, max_value=50, value=20)
                        
                    if chart_dimension == "2-Digit NIC Division":
                        group_col = "division"; name_map = ASI_NIC_NAMES
                    elif chart_dimension == "5-Digit Granular NIC":
                        group_col = "a5"; name_map = {}
                        if "a5" in full_df.columns:
                            for code in full_df["a5"].unique():
                                if isinstance(code, str) and len(code) >= 2:
                                    if code in granular_nic_map:
                                        exact_name = granular_nic_map[code]
                                        if len(exact_name) > 35: exact_name = exact_name[:32] + "..."
                                        name_map[code] = f"{code} ({exact_name})"
                                    else:
                                        parent_name = ASI_NIC_NAMES.get(code[:2], "Other Industry")
                                        name_map[code] = f"{code} ({parent_name})"
                    elif chart_dimension == "Geographical State Location":
                        group_col = "a7"; name_map = ASI_STATE_NAMES
                    else:
                        group_col = "a9"; name_map = {"1": "Rural Ecosystems", "2": "Urban Corridors", 1: "Rural Ecosystems", 2: "Urban Corridors"}
                        
                    studio_rows = []
                    if group_col in full_df.columns:
                        for val_id, grp in full_df.groupby(group_col):
                            if grp.shape[0] < min_sample_threshold: continue
                            w_wk = pop_estimate(grp, "total_workers")
                            w_op = pop_estimate(grp, "pc_15_total_output")
                            w_em = pop_estimate(grp, "total_emoluments")
                            w_wom = pop_estimate(grp, "women_workers")
                            w_fixed = pop_estimate(grp, "pc_03_fixed_capital")
                            
                            is_tgt = False
                            if chart_dimension == "2-Digit NIC Division" and str(val_id) in [c[:2] for c in nic_codes_list]: is_tgt = True
                            elif chart_dimension == "5-Digit Granular NIC" and str(val_id) in nic_codes_list: is_tgt = True
                            
                            row_data = {
                                "id": val_id, "label_name": name_map.get(str(val_id), f"Code ID {val_id}"), "is_target_group": is_tgt,
                                "Emp per Crore Investment": (w_wk / (w_fixed / 1e7)) if w_fixed > 0 else 0,
                                "Output per Worker (Lakh)": (w_op / w_wk / 1e5) if w_wk > 0 else 0,
                                "Labour Income (%)": (w_em / w_op * 100) if w_op > 0 else 0,
                                "Women Workforce Share (%)": (w_wom / w_wk * 100) if w_wk > 0 else 0,
                                "Plant Density Count": grp.shape[0], "Total Segment Workforce": w_wk
                            }
                            
                            for ext_col in extra_metric_cols:
                                if ext_col.startswith('pc_'):
                                    dict_label = ASI_VAR_LABELS.get(ext_col.lower(), ext_col.upper().replace('_', ' '))
                                else:
                                    label_key = ext_col.split('_')[-1].lower() if ext_col.startswith(('f_', 'g_', 'h_')) else ext_col.lower()
                                    dict_label = f"{ASI_VAR_LABELS.get(label_key, ASI_VAR_LABELS.get(ext_col.lower(), 'Raw Data: ' + ext_col.upper()))} ({ext_col})"
                                
                                row_data[dict_label] = pop_estimate(grp, ext_col)
                                
                            studio_rows.append(row_data)
                        
                    studio_plot_df = pd.DataFrame(studio_rows)
                    
                    with studio_col2:
                        if studio_plot_df.empty:
                            st.warning("⚠️ No database rows passed the current minimum factory cut-off threshold.")
                        else:
                            studio_plot_df = studio_plot_df.sort_values(chart_metric, ascending=False).reset_index(drop=True)
                            plot_df = studio_plot_df.copy()
                            if len(plot_df) > max_display_items:
                                st.info(f"Chart showing Top {max_display_items} items (out of {len(plot_df)}).")
                                plot_df = plot_df.head(max_display_items)
                                
                            fig, ax = plt.subplots(figsize=(12, 6))
                            colors = [chart_palette if row["is_target_group"] else "#D2D7DF" for _, row in plot_df.iterrows()]
                            if not any(plot_df["is_target_group"]): colors = [chart_palette] * len(plot_df)
                                
                            if chart_type == "Bar Chart":
                                bars = ax.bar(plot_df["label_name"], plot_df[chart_metric], color=colors, edgecolor="#4A5568", linewidth=0.6)
                                ax.set_xticks(range(len(plot_df["label_name"])))
                                ax.set_xticklabels([textwrap.fill(str(x), width=16) for x in plot_df["label_name"]], rotation=45, ha="right", fontsize=9)
                                ax.set_ylabel(chart_metric, fontweight="bold")
                                for bar in bars[:5]:
                                    height = bar.get_height()
                                    ax.text(bar.get_x() + bar.get_width()/2., height + (height * 0.01), f'{height:.1f}', ha='center', va='bottom', fontsize=8)
                            elif chart_type == "Line Chart":
                                ax.plot(plot_df["label_name"], plot_df[chart_metric], marker='o', color=chart_palette, linewidth=2, markersize=8)
                                ax.fill_between(plot_df["label_name"], plot_df[chart_metric], color=chart_palette, alpha=0.1)
                                ax.set_xticks(range(len(plot_df["label_name"])))
                                ax.set_xticklabels([textwrap.fill(str(x), width=16) for x in plot_df["label_name"]], rotation=45, ha="right", fontsize=9)
                                ax.set_ylabel(chart_metric, fontweight="bold")
                            elif chart_type == "Scatter Plot":
                                scatter = ax.scatter(
                                    plot_df[scatter_x], plot_df[chart_metric], 
                                    s=np.clip(plot_df["Plant Density Count"]*2, 40, 600), 
                                    c=["#E53E3E" if r["is_target_group"] else chart_palette for _, r in plot_df.iterrows()],
                                    alpha=0.75, edgecolors="#2D3748"
                                )
                                ax.set_xlabel(scatter_x, fontweight="bold"); ax.set_ylabel(chart_metric, fontweight="bold")
                                for _, row in plot_df.iterrows():
                                    if row["is_target_group"] or row[chart_metric] == plot_df[chart_metric].max():
                                        ax.text(row[scatter_x], row[chart_metric], f" {row['label_name']}", fontsize=9, fontweight="bold")
                            elif chart_type == "Pie Chart":
                                pie_data = plot_df[chart_metric].clip(lower=0)
                                explode = [0.08 if row["is_target_group"] else 0 for _, row in plot_df.iterrows()]
                                if not any(plot_df["is_target_group"]) and len(explode) > 0: explode[0] = 0.08
                                ax.pie(pie_data, labels=plot_df["label_name"], autopct='%1.1f%%', startangle=140, colors=colors, explode=explode, textprops={'fontsize': 9}, wedgeprops={'edgecolor': 'white', 'linewidth': 1})
                            
                            plt.title(f"Dynamic Analysis: {chart_metric} analyzed across {chart_dimension}", fontweight="bold", pad=15, color="#1A202C")
                            if chart_type != "Pie Chart":
                                ax.spines['top'].set_visible(False)
                                ax.spines['right'].set_visible(False)
                                
                            fig.tight_layout()
                            st.pyplot(fig)
                    
                    st.markdown("## 🔍 Deep-Dive Microdata Audit Ledgers")
                    tab_led1, tab_led2, tab_led3 = st.tabs(["📊 National Comparative Matrix", "🔬 Studio Grid View Data", "💾 Bulk Export Vault"])
                    
                    with tab_led1:
                        st.markdown("##### National Macro Performance Architecture Ledger")
                        st.dataframe(summary_table_df.style.format({
                            'Surveyed Plants': '{:,.0f}', 'Est Population Size': '{:,.0f}', 'Output per Worker (Lakh)': '{:.2f}L', 
                            'Emp per Crore Investment': '{:.1f}', 'Labour Income (%)': '{:.2f}%', 'Women Workforce Share (%)': '{:.2f}%',
                            'Workforce Size (Mn)': '{:.3f} Mn', 'Capital Stock Asset (Cr)': '₹{:,.2f} Cr'
                        }))
                    with tab_led2:
                        st.markdown(f"##### Normalized Custom Output Metrics Ledger (Grouped by {chart_dimension})")
                        style_dict = {
                            'Emp per Crore Investment': '{:.2f}', 'Output per Worker (Lakh)': '{:.2f}L', 'Labour Income (%)': '{:.2f}%', 
                            'Women Workforce Share (%)': '{:.2f}%', 'Plant Density Count': '{:,.0f}', 'Total Segment Workforce': '{:,.0f}'
                        }
                        st.dataframe(studio_plot_df.style.format(style_dict, na_rep="-"))
                    with tab_led3:
                        st.markdown("##### Download Intelligence Workspace Pack")
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            summary_table_df.to_excel(writer, sheet_name='National Aggregates Overview', index=False)
                            studio_plot_df.to_excel(writer, sheet_name=f'Custom {chart_dimension[:15]} Ledger', index=False)
                        st.download_button(
                            label=f"📥 Download Selected Custom {sector_name} Export Book (Excel)",
                            data=excel_buffer.getvalue(),
                            file_name=f"ASI_Custom_{sector_name.replace(' ', '_')}_Studio_Report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.warning(f"⚠️ No dataset records matched your selected NIC configuration entry values.")

def run_downloader_app():
    st.title("📥 Screener Mass-Downloader (Local Tool)")
    
    st.markdown("""
    ### Why run locally?
    Cloud servers do not have screens, which means automated scripts cannot prompt you to manually log in and solve CAPTCHAs. To securely access your Screener.in account and download files directly to your machine, this tool must be run locally.
    
    ### 🛠️ Setup Instructions
    
    **1. Download the App Script**
    Click the button below to download the dedicated Python file.
    """)
    
    downloader_script = '''import streamlit as st
import time
import random
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

st.set_page_config(page_title="Screener Downloader UI", layout="centered")
st.title("Screener Mass-Download Interface")

# --- Configuration ---
DOWNLOAD_DIR = os.path.join(os.getcwd(), "screener_exports")
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

def setup_driver():
    """Configures Chrome for local automation and direct downloading."""
    chrome_options = Options()
    
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
        'source': "Object.defineProperty(navigator, 'webdriver', { get: () => undefined })"
    })
    
    return driver

# --- User Interface ---

st.markdown("Enter the tickers you want to download. The tool will open a browser, wait for you to log in, and then systematically extract the Excel data sheets.")

# Ticker Input Area
default_tickers = "RELIANCE, TCS, INFY, HDFCBANK, ITC"
tickers_input = st.text_area("Target Tickers (comma-separated):", value=default_tickers)

# Extraction Controls
if st.button("Launch Extraction Sequence", type="primary"):
    
    # Clean the input list
    raw_tickers = tickers_input.split(',')
    target_tickers = [t.strip().upper() for t in raw_tickers if t.strip()]
    
    if not target_tickers:
        st.error("Please enter at least one ticker.")
    else:
        st.info(f"Initialized sequence for {len(target_tickers)} companies. Launching browser...")
        
        try:
            driver = setup_driver()
            
            # 1. Auto-Detect Login Phase
            driver.get("https://www.screener.in/login/")
            
            with st.status("Waiting for manual login...", expanded=True) as status:
                st.write("1. A browser window has opened.")
                st.write("2. Please log in to Screener and solve any captchas.")
                st.write("3. The system is monitoring the URL and will proceed automatically once you reach the dashboard.")
                
                # Check the URL every second for up to 3 minutes
                logged_in = False
                for _ in range(180):
                    current_url = driver.current_url
                    if "login" not in current_url.lower():
                        logged_in = True
                        break
                    time.sleep(1)
                
                if not logged_in:
                    status.update(label="Login timed out.", state="error")
                    st.stop()
                    
                status.update(label="Authentication confirmed! Commencing downloads.", state="complete")

            # 2. Automated Download Loop
            progress_bar = st.progress(0)
            console = st.empty()
            
            success_count = 0
            failed_tickers = []
            
            for i, ticker in enumerate(target_tickers):
                console.info(f"Processing: {ticker} ({i+1}/{len(target_tickers)})")
                
                try:
                    url = f"https://www.screener.in/company/{ticker}/consolidated/"
                    driver.get(url)
                    time.sleep(random.uniform(2.5, 4.0))
                    
                    # Target and click the Export button via JS bypass
                    xpath = "//form[contains(@action, '/excel/')]//button | //button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'export to excel')]"
                    
                    export_btn = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, xpath))
                    )
                    
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", export_btn)
                    time.sleep(1.5)
                    driver.execute_script("arguments[0].click();", export_btn)
                    
                    success_count += 1
                    
                    # Mandatory Anti-Ban Delay
                    delay = random.uniform(8.0, 15.0)
                    console.warning(f"Successfully triggered {ticker}. Mimicking human delay for {delay:.1f}s...")
                    time.sleep(delay)
                    
                except Exception as e:
                    console.error(f"Failed to extract {ticker}. Skipping.")
                    failed_tickers.append(ticker)
                    time.sleep(5)
                
                # Update visual progress
                progress_bar.progress((i + 1) / len(target_tickers))
                
            driver.quit()
            
            # 3. Post-Run Summary
            st.success(f"Extraction Complete! Successfully downloaded {success_count} files.")
            st.markdown(f"**Files saved to:** `{DOWNLOAD_DIR}`")
            
            if failed_tickers:
                st.error(f"Failed to process the following tickers: {', '.join(failed_tickers)}")
                st.markdown("Ensure these tickers exist exactly as spelled on Screener.in.")
                
        except Exception as e:
            st.error(f"A critical error occurred: {e}")
            try:
                driver.quit()
            except:
                pass'''
    
    st.download_button(
        label="Download `app_downloader.py`",
        data=downloader_script,
        file_name="app_downloader.py",
        mime="text/x-python"
    )
    
    st.markdown("""
    **2. Install Dependencies**
    
    Open your terminal or command prompt and install the required libraries:
    ```bash
    pip install streamlit selenium webdriver-manager
    ```
    
    **3. Run the App**
    
    Navigate to the folder where you saved the script and run:
    ```bash
    streamlit run app_downloader.py
    ```
    This will open the interface in your local browser, allowing you to seamlessly log in to Screener and automatically extract the requested financial sheets to your machine.
    """)

def run_financial_app():
    st.title("Financial Analyzer")
    st.markdown("Upload multiple **Data Sheet** (Excel/CSV) exports from Screener.in to analyze and compare them.")
    uploaded_files = st.file_uploader("Upload Screener.in Exports", type=['xlsx', 'csv'], accept_multiple_files=True)

    if uploaded_files:
        company_data = {}
        with st.spinner("Processing files and aligning financial periods..."):
            for file in uploaded_files:
                company_name = file.name.replace(' - Data Sheet.csv', '').replace('.csv', '').replace('.xlsx', '')
                if file.name.endswith('.csv'): raw_df = pd.read_csv(file, header=None)
                else: raw_df = pd.read_excel(file, sheet_name='Data Sheet', header=None)
                    
                pl_df = extract_section(raw_df, 'PROFIT & LOSS', 'Quarters')
                bs_df = extract_section(raw_df, 'BALANCE SHEET', 'CASH FLOW:')
                
                if pl_df is not None and bs_df is not None:
                    company_data[company_name] = {'pl': pl_df, 'bs': bs_df}
                    
        if not company_data: st.error("Could not extract valid data blocks from the uploaded files.")
        else:
            st.success(f"Successfully processed {len(company_data)} companies.")
            pl_dict = {name: data['pl'] for name, data in company_data.items()}
            bs_dict = {name: data['bs'] for name, data in company_data.items()}
            
            avg_pl_df = calculate_aggregate_average(pl_dict)
            avg_bs_df = calculate_aggregate_average(bs_dict)
            
            avg_revenue_row = next((row for row in avg_pl_df.index if pd.notna(row) and ('sales' in str(row).lower() or 'revenue' in str(row).lower())), None)
            avg_cs_pl, avg_cs_bs = None, None
            if avg_revenue_row:
                avg_cs_pl = calculate_percentage_of_revenue(avg_pl_df, avg_pl_df.loc[avg_revenue_row])
                avg_cs_bs = calculate_percentage_of_revenue(avg_bs_df, avg_pl_df.loc[avg_revenue_row])

            tab1, tab2, tab3, tab4 = st.tabs(["Individual Analysis", "Aggregate Averages", "Comparative Trends", "Avg Common Size Trends"])
            
            with tab1:
                selected_company = st.selectbox("Select Company to View:", list(company_data.keys()))
                c_pl = company_data[selected_company]['pl']
                c_bs = company_data[selected_company]['bs']
                c_rev_row = next((r for r in c_pl.index if pd.notna(r) and ('sales' in str(r).lower() or 'revenue' in str(r).lower())), None)
                
                st.subheader(f"Financial Statements: {selected_company}")
                st.dataframe(c_pl, use_container_width=True)
                st.dataframe(c_bs, use_container_width=True)
                
                if c_rev_row:
                    st.subheader("Common Size Analysis (Base: Total Revenue)")
                    col1, col2 = st.columns(2)
                    with col1: st.dataframe(calculate_percentage_of_revenue(c_pl, c_pl.loc[c_rev_row]), use_container_width=True)
                    with col2: st.dataframe(calculate_percentage_of_revenue(c_bs, c_pl.loc[c_rev_row]), use_container_width=True)

            with tab2:
                st.markdown("### Industry / Group Averages")
                st.markdown(f"Displays the mathematical mean for every line item across all **{len(company_data)}** uploaded companies.")
                st.subheader("Average Profit & Loss")
                st.dataframe(avg_pl_df, use_container_width=True)
                st.subheader("Average Balance Sheet")
                st.dataframe(avg_bs_df, use_container_width=True)
                
                if avg_cs_pl is not None and avg_cs_bs is not None:
                    st.markdown("### Average Common Size Analysis")
                    st.markdown("Calculates the common size percentages using the **Average Line Items** divided by the **Average Total Revenue**.")
                    col1, col2 = st.columns(2)
                    with col1: st.dataframe(avg_cs_pl, use_container_width=True)
                    with col2: st.dataframe(avg_cs_bs, use_container_width=True)
                else: st.warning("Could not identify the Revenue row to calculate aggregate common size.")

            with tab3:
                st.subheader("Cross-Company Comparative Analysis")
                valid_items = list(dict.fromkeys([item for item in list(avg_pl_df.index) + list(avg_bs_df.index) if pd.notna(item)]))
                selected_item = st.selectbox("Select a line item to compare across companies:", valid_items, key="comp_trend")
                
                if selected_item:
                    fig = go.Figure()
                    for name, data in company_data.items():
                        target_df = data['pl'] if selected_item in data['pl'].index else data['bs']
                        if selected_item in target_df.index:
                            series = target_df.loc[selected_item].dropna()
                            fig.add_trace(go.Scatter(x=series.index, y=series.values, mode='lines+markers', name=name, opacity=0.6))
                    
                    target_avg_df = avg_pl_df if selected_item in avg_pl_df.index else avg_bs_df
                    if selected_item in target_avg_df.index:
                        avg_series = target_avg_df.loc[selected_item].dropna()
                        fig.add_trace(go.Scatter(x=avg_series.index, y=avg_series.values, mode='lines+markers', name='GROUP AVERAGE', line=dict(width=4, color='white', dash='dot')))
                    
                    fig.update_layout(title=f"Comparative Trend: {selected_item}", xaxis_title="Financial Year", yaxis_title="Reported Value")
                    st.plotly_chart(fig, use_container_width=True)

            with tab4:
                st.subheader("Industry-Wide Common Size Trends")
                if avg_cs_pl is not None and avg_cs_bs is not None:
                    valid_cs_items = list(dict.fromkeys([item for item in list(avg_cs_pl.index) + list(avg_cs_bs.index) if pd.notna(item)]))
                    selected_cs_item = st.selectbox("Select a metric to view its structural trend:", valid_cs_items, key="cs_trend")
                    
                    if selected_cs_item:
                        plot_data_cs = None
                        if selected_cs_item in avg_cs_pl.index: plot_data_cs = avg_cs_pl.loc[selected_cs_item]
                        elif selected_cs_item in avg_cs_bs.index: plot_data_cs = avg_cs_bs.loc[selected_cs_item]
                            
                        if plot_data_cs is not None:
                            plot_data_cs = pd.to_numeric(plot_data_cs, errors='coerce').dropna()
                            trend_df_cs = pd.DataFrame({"Financial Period": plot_data_cs.index, "Percentage of Revenue (%)": plot_data_cs.values})
                            fig_cs = px.line(trend_df_cs, x="Financial Period", y="Percentage of Revenue (%)", markers=True, title=f"Group Average: {selected_cs_item} as % of Revenue")
                            fig_cs.update_layout(yaxis_ticksuffix="%")
                            st.plotly_chart(fig_cs, use_container_width=True)
                else: st.warning("Average common size data is unavailable because a base revenue row could not be established.")

def run_plfs_app():
    st.title("📊 PLFS Analyzer")
    st.info("ℹ️ **Layout Status:** This workspace uses microdata layouts configured for the **2025 CPERV1 Scheme**. Use the **'⚙️ Configuration Hub'** to inject custom variables or recalibrate fixed-width coordinates.")

    st.sidebar.header("⚙️ Configuration Hub")
    layout_mode = st.sidebar.radio("Active Layout Blueprint:", options=["Standard 2025 Default Specs", "Custom Suffix/Byte Overrides"])

    base_colnames = ["state_raw", "sector_raw", "sex", "age_raw", "marital_raw", "edu_raw", "pas", "ind_pas", "mult"]

    if layout_mode == "Standard 2025 Default Specs":
        base_colspecs = [(15, 17), (14, 15), (47, 48), (48, 51), (51, 52), (52, 54), (78, 80), (80, 85), (340, 350)]
    else:
        st.sidebar.markdown("#### Manual Byte Adjustments")
        with st.sidebar.expander("Calibrate Core Fixed-Width Coordinates", expanded=True):
            state_s = st.number_input("State Start Byte", min_value=1, value=16)
            state_e = st.number_input("State End Byte", min_value=1, value=17)
            sec_s = st.number_input("Sector Start Byte", min_value=1, value=15)
            sec_e = st.number_input("Sector End Byte", min_value=1, value=15)
            sex_s = st.number_input("Sex Start Byte", min_value=1, value=48)
            sex_e = st.number_input("Sex End Byte", min_value=1, value=48)
            age_s = st.number_input("Age Start Byte", min_value=1, value=49)
            age_e = st.number_input("Age End Byte", min_value=1, value=51)
            marital_s = st.number_input("Marital Status Start Byte", min_value=1, value=52)
            marital_e = st.number_input("Marital Status End Byte", min_value=1, value=52)
            edu_s = st.number_input("Education Start Byte", min_value=1, value=53)
            edu_e = st.number_input("Education End Byte", min_value=1, value=54)
            pas_s = st.number_input("PAS Start Byte", min_value=1, value=79)
            pas_e = st.number_input("PAS End Byte", min_value=1, value=80)
            ind_s = st.number_input("NIC Industry Start Byte", min_value=1, value=81)
            ind_e = st.number_input("NIC Industry End Byte", min_value=1, value=85)
            mult_s = st.number_input("Weight Multiplier Start Byte", min_value=1, value=341)
            mult_e = st.number_input("Weight Multiplier End Byte", min_value=1, value=350)
            
        base_colspecs = [
            (int(state_s)-1, int(state_e)), (int(sec_s)-1, int(sec_e)), (int(sex_s)-1, int(sex_e)), 
            (int(age_s)-1, int(age_e)), (int(marital_s)-1, int(marital_e)), (int(edu_s)-1, int(edu_e)),
            (int(pas_s)-1, int(pas_e)), (int(ind_s)-1, int(ind_e)), (int(mult_s)-1, int(mult_e))
        ]

    st.sidebar.markdown("#### ➕ Add Custom Variables")
    custom_names_input = st.sidebar.text_input("Variable Names (comma separated)", placeholder="e.g., Training, SocialGroup")
    custom_bytes_input = st.sidebar.text_input("Byte Ranges (comma separated)", placeholder="e.g., 69-69, 70-71")

    final_colnames = base_colnames.copy()
    final_colspecs = base_colspecs.copy()
    c_names = []

    if custom_names_input and custom_bytes_input:
        c_names = [n.strip() for n in custom_names_input.split(",")]
        c_bytes_str = [b.strip() for b in custom_bytes_input.split(",")]
        
        if len(c_names) == len(c_bytes_str):
            for i, name in enumerate(c_names):
                try:
                    start_b, end_b = map(int, c_bytes_str[i].split("-"))
                    final_colnames.append(name)
                    final_colspecs.append((start_b - 1, end_b))
                except Exception as e:
                    st.sidebar.error(f"Error parsing byte range for {name}. Format should be Start-End (e.g., 10-12)")
        else:
            st.sidebar.error("Mismatched number of Custom Variables and Byte Ranges.")

    st.sidebar.markdown("---")
    st.sidebar.header("🎨 Global Theme Settings")
    primary_color = st.sidebar.color_picker("Primary Chart Color", value="#2B6CB0")
    secondary_color = st.sidebar.color_picker("Secondary Chart Color (Comparisons)", value="#D53F8C")

    st.markdown("### 📥 Primary Data Ingestion")
    
    use_custom_plfs = st.checkbox("Upload custom PLFS data file instead of using local default")
    
    uploaded_file = None
    if use_custom_plfs:
        uploaded_file = st.file_uploader("Upload CPERV1.TXT Data File", type=['txt'])
    else:
        default_plfs = os.path.join("data", "CPERV1.TXT")
        if os.path.exists(default_plfs):
            uploaded_file = default_plfs
            st.success(f"✅ Automatically loaded local PLFS default file from the `data/` directory.")
        else:
            st.info(f"💡 Default PLFS file (`CPERV1.TXT`) was not found in the `data/` directory. Please upload it manually by checking the box above.")

    if uploaded_file is not None:
        with st.spinner('Parsing PLFS microdata framework and mapping structures...'):
            full_population_df = process_plfs_base_data(uploaded_file, final_colspecs, final_colnames)
        
        if not full_population_df.empty:
            tab_macro, tab_mfg = st.tabs(["📈 Macro Labour Stats (LFPR & WPR)", "🏭 Manufacturing Deep-Dive"])
            
            with tab_macro:
                st.markdown("### National Labour Force Participation & Worker Population Ratio")
                st.write("Analyze structural changes in the workforce across core demographics, geography, and socio-economic markers.")
                
                macro_col1, macro_col2 = st.columns([1, 3])
                with macro_col1:
                    group_by_macro = st.selectbox(
                        "Primary Demographic Grouping:", 
                        options=["age_group", "gender_label", "region_label", "state_label", "education_level", "marital_status", "nic_2d", "worker_category"],
                        format_func=lambda x: {
                            "age_group": "Age Cohort", "gender_label": "Gender", "region_label": "Sector (Rural/Urban)", 
                            "state_label": "State / Union Territory", "education_level": "General Education Level",
                            "marital_status": "Marital Status", "nic_2d": "Industry (2-Digit NIC)", 
                            "worker_category": "Worker Category (Salaried/Casual)"
                        }[x]
                    )
                    adults_only = st.checkbox("Restrict to Age 15+ (Standard Definition)", value=True)
                    chart_display_mode = st.radio("Chart Display Mode:", options=["Combined Chart", "Separate Charts"])
                
                macro_df = full_population_df.copy()
                if adults_only: macro_df = macro_df[macro_df['age'] >= 15]

                agg_rows = []
                for val_id, grp in macro_df.groupby(group_by_macro):
                    if pd.isna(val_id) or str(val_id).strip() == "": continue
                    total_pop_wt = grp['final_wt'].sum()
                    lf_wt = grp[grp['is_lf']]['final_wt'].sum()
                    emp_wt = grp[grp['is_employed']]['final_wt'].sum()
                    lfpr = (lf_wt / total_pop_wt * 100) if total_pop_wt > 0 else 0
                    wpr = (emp_wt / total_pop_wt * 100) if total_pop_wt > 0 else 0
                    label = str(val_id)
                    if group_by_macro == "nic_2d": label = f"{label} - {PLFS_NIC_DICTIONARY.get(label, 'Other')}"
                        
                    agg_rows.append({
                        "Grouping": label, "LFPR (%)": lfpr, "WPR (%)": wpr,
                        "Total Population Est.": total_pop_wt, "Total Employed Est.": emp_wt
                    })
                    
                macro_plot_df = pd.DataFrame(agg_rows)
                
                with macro_col2:
                    if not macro_plot_df.empty:
                        if group_by_macro == "nic_2d":
                            macro_plot_df = macro_plot_df.sort_values("Total Employed Est.", ascending=False).head(15)
                            st.info("Displaying Top 15 Industries by Employment Volume.")
                        elif group_by_macro in ["state_label", "worker_category", "education_level"]:
                            macro_plot_df = macro_plot_df.sort_values("Total Population Est.", ascending=False)
                            if group_by_macro == "worker_category": st.info("Note: LFPR and WPR evaluate to 100% strictly within employed worker categories.")
                        else:
                            macro_plot_df = macro_plot_df.sort_values("Grouping")
                        
                        x = np.arange(len(macro_plot_df))
                        rot = 90 if group_by_macro in ["state_label", "worker_category", "education_level"] else 45
                        x_labels = [textwrap.fill(str(l), width=20) for l in macro_plot_df['Grouping']]
                        
                        if chart_display_mode == "Combined Chart":
                            fig, ax = plt.subplots(figsize=(10, 5))
                            width = 0.35
                            ax.bar(x - width/2, macro_plot_df['LFPR (%)'], width, label='LFPR', color=primary_color)
                            ax.bar(x + width/2, macro_plot_df['WPR (%)'], width, label='WPR', color=secondary_color)
                            ax.set_ylabel('Percentage (%)', fontweight='bold'); ax.set_title(f'LFPR and WPR distribution by {group_by_macro}', fontweight='bold')
                            ax.set_xticks(x); ax.set_xticklabels(x_labels, rotation=rot, ha="center" if rot==90 else "right")
                            ax.legend(); ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
                            fig.tight_layout()
                            st.pyplot(fig)
                        else:
                            fig1, ax1 = plt.subplots(figsize=(10, 4))
                            ax1.bar(x, macro_plot_df['LFPR (%)'], color=primary_color)
                            ax1.set_ylabel('LFPR (%)', fontweight='bold'); ax1.set_title(f'Labour Force Participation Rate (LFPR) by {group_by_macro}', fontweight='bold')
                            ax1.set_xticks(x); ax1.set_xticklabels(x_labels, rotation=rot, ha="center" if rot==90 else "right")
                            ax1.spines['top'].set_visible(False); ax1.spines['right'].set_visible(False)
                            fig1.tight_layout(); st.pyplot(fig1)

                            fig2, ax2 = plt.subplots(figsize=(10, 4))
                            ax2.bar(x, macro_plot_df['WPR (%)'], color=secondary_color)
                            ax2.set_ylabel('WPR (%)', fontweight='bold'); ax2.set_title(f'Worker Population Ratio (WPR) by {group_by_macro}', fontweight='bold')
                            ax2.set_xticks(x); ax2.set_xticklabels(x_labels, rotation=rot, ha="center" if rot==90 else "right")
                            ax2.spines['top'].set_visible(False); ax2.spines['right'].set_visible(False)
                            fig2.tight_layout(); st.pyplot(fig2)
                        
                        st.dataframe(macro_plot_df.style.format({'LFPR (%)': '{:.2f}%', 'WPR (%)': '{:.2f}%', 'Total Population Est.': '{:,.0f}', 'Total Employed Est.': '{:,.0f}'}))
                        
            with tab_mfg:
                workforce = full_population_df[full_population_df['is_employed']].copy()
                workforce = workforce[workforce['ind_pas'].str.match(r'^\d+$', na=False)]
                workforce['nic_int'] = workforce['ind_pas'].astype(int)
                mfg_base = workforce[(workforce['nic_int'] >= 10000) & (workforce['nic_int'] <= 33999)].copy()
                
                st.sidebar.markdown("---")
                st.sidebar.header("🎯 Demographics Filters (Mfg Tab)")
                keep_rural = st.sidebar.checkbox("Rural", value=True)
                keep_urban = st.sidebar.checkbox("Urban", value=True)
                keep_male = st.sidebar.checkbox("Male", value=True)
                keep_female = st.sidebar.checkbox("Female", value=True)
                
                allowed_sectors = [s for s, keep in zip(["Rural", "Urban"], [keep_rural, keep_urban]) if keep]
                allowed_genders = [g for g, keep in zip(["Male", "Female"], [keep_male, keep_female]) if keep]
                
                mfg_filtered = mfg_base[(mfg_base['region_label'].isin(allowed_sectors)) & (mfg_base['gender_label'].isin(allowed_genders))].copy()
                mfg_filtered['nic_desc'] = mfg_filtered['nic_2d'].map(lambda x: f"{x} - {PLFS_NIC_DICTIONARY.get(x, 'Unclassified Sub-sector')}")
                
                all_unique_targets = sorted(list(mfg_filtered['nic_2d'].dropna().unique()))
                friendly_options = sorted([mfg_filtered[mfg_filtered['nic_2d'] == code]['nic_desc'].iloc[0] for code in all_unique_targets if code])
                selected_friendly = st.sidebar.multiselect("Choose Industries to Display:", options=friendly_options, default=friendly_options[:5] if friendly_options else [])
                selected_codes = [s.split(" - ")[0].strip() for s in selected_friendly]
                
                final_df = mfg_filtered[mfg_filtered['nic_2d'].isin(selected_codes)].copy()
                
                if not final_df.empty:
                    tot_est = final_df['final_wt'].sum()
                    r_share = final_df[final_df['region_label'] == 'Rural']['final_wt'].sum() / tot_est * 100 if tot_est > 0 else 0
                    w_share = final_df[final_df['gender_label'] == 'Female']['final_wt'].sum() / tot_est * 100 if tot_est > 0 else 0
                    
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Filtered Worker Population", f"{tot_est / 1e6:.2f} Million")
                    c2.metric("Rural Structural Share", f"{r_share:.1f}%")
                    c3.metric("Female Representation Rate", f"{w_share:.1f}%")
                    
                    st.markdown("## 📊 Custom Visualization Studio")
                    studio_col1, studio_col2 = st.columns([1, 3])
                    grouping_options = ["Industry (NIC)", "State/UT", "Geographical Region", "Gender Grouping", "Worker Category", "Education Level", "Marital Status"] + (c_names if custom_names_input else [])
                    
                    with studio_col1:
                        chart_dimension = st.selectbox("Data Grouping (Categories):", options=grouping_options)
                        chart_metric = st.selectbox("Primary Measurement:", options=["Total Workforce Volume (Millions)", "Survey Sample Size (Raw Count)"])
                        chart_type = st.selectbox("Visualization Style:", ["Bar Chart", "Pie Chart"])
                    
                    if chart_dimension == "Industry (NIC)": group_col = "nic_desc"
                    elif chart_dimension == "State/UT": group_col = "state_label"
                    elif chart_dimension == "Geographical Region": group_col = "region_label"
                    elif chart_dimension == "Gender Grouping": group_col = "gender_label"
                    elif chart_dimension == "Worker Category": group_col = "worker_category"
                    elif chart_dimension == "Education Level": group_col = "education_level"
                    elif chart_dimension == "Marital Status": group_col = "marital_status"
                    else: group_col = chart_dimension 
                        
                    studio_rows = []
                    for val_id, grp in final_df.groupby(group_col):
                        studio_rows.append({"label_name": str(val_id), "Total Workforce Volume (Millions)": grp['final_wt'].sum() / 1e6, "Survey Sample Size (Raw Count)": len(grp)})
                    studio_plot_df = pd.DataFrame(studio_rows).sort_values(chart_metric, ascending=False)
                    
                    with studio_col2:
                        fig2, ax2 = plt.subplots(figsize=(10, 5))
                        if chart_type == "Bar Chart":
                            ax2.bar(studio_plot_df["label_name"].astype(str), studio_plot_df[chart_metric], color=primary_color)
                            ax2.set_xticks(range(len(studio_plot_df)))
                            rot = 90 if chart_dimension in ["State/UT", "Worker Category", "Education Level"] else 45
                            ax2.set_xticklabels([textwrap.fill(str(x), width=20) for x in studio_plot_df["label_name"]], rotation=rot, ha="center" if rot==90 else "right")
                            ax2.set_ylabel(chart_metric, fontweight="bold")
                            ax2.spines['top'].set_visible(False); ax2.spines['right'].set_visible(False)
                        else:
                            ax2.pie(studio_plot_df[chart_metric], labels=studio_plot_df["label_name"], autopct='%1.1f%%')
                        
                        fig2.tight_layout()
                        st.pyplot(fig2)
                else: st.warning("⚠️ No records matched your active demographic filter selections.")

# ==============================================================================
# 3. MASTER NAVIGATION & EXECUTION PIPELINE
# ==============================================================================

st.sidebar.title("App Navigation")
app_choice = st.sidebar.radio(
    "Select an Application to Run:",
    [
        "🏠 Dashboard Home", 
        "🏭 ASI", 
        "🌍 Comtrade",
        "📥 Screener Mass-Downloader", 
        "📈 Screener Financials Analyzer", 
        "📊 PLFS"
    ]
)

st.sidebar.markdown("---")

if app_choice == "🏠 Dashboard Home":
    st.title("Unified Analytics Hub")
    st.markdown("""
        **Select a module from the sidebar on the left to begin.**
    """)
elif app_choice == "🏭 ASI":
    run_asi_app()
elif app_choice == "🌍 Comtrade":
    run_comtrade_app()
elif app_choice == "📥 Screener Mass-Downloader":
    run_downloader_app()
elif app_choice == "📈 Screener Financials Analyzer":
    run_financial_app()
elif app_choice == "📊 PLFS":
    run_plfs_app()
