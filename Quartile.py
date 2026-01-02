import streamlit as st
import pandas as pd
from io import BytesIO

# ===============================
# PAGE CONFIG
# ===============================
st.set_page_config(
    page_title="Dynamic Metrics Report",
    page_icon="📊",
    layout="wide"
)

# ===============================
# CUSTOM CSS
# ===============================
st.markdown("""
<style>
.main-title {
    font-size: 34px;
    font-weight: 700;
}
.sub-title {
    font-size: 16px;
    color: #6b7280;
}
.section-title {
    font-size: 22px;
    font-weight: 600;
    margin-top: 25px;
}
.kpi-card {
    background-color: #ffffff;
    padding: 18px;
    border-radius: 14px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    text-align: center;
}
.kpi-label {
    font-size: 14px;
    color: #6b7280;
}
.kpi-value {
    font-size: 26px;
    font-weight: 700;
}
</style>
""", unsafe_allow_html=True)

# ===============================
# HEADER
# ===============================
st.markdown('<div class="main-title">📊 Dynamic Metrics Report</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Kategori & Format Performance Overview</div>', unsafe_allow_html=True)
st.divider()

# ===============================
# FILE UPLOAD (SIDEBAR)
# ===============================
with st.sidebar:
    st.header("📁 Data Source")
    format_file = st.file_uploader("Upload Format File", type=["xlsx"])
    category_file = st.file_uploader("Upload Kategori File", type=["xlsx"])

if not format_file or not category_file:
    st.info("⬅️ Upload **Format** dan **Kategori** file terlebih dahulu")
    st.stop()

# ===============================
# (SELURUH LOGIKA LOAD MAP TETAP SAMA)
# ===============================
# >>> COPY PASTE LOGIKA LOAD DATA KAMU DI SINI TANPA PERUBAHAN <<<

# ===============================
# FILTERS (SIDEBAR)
# ===============================
st.sidebar.divider()
st.sidebar.header("🎯 Filter")

# kategori
if "category_select" not in st.session_state:
    st.session_state["category_select"] = list(cont_map_cat.keys())

cat_mode = st.sidebar.radio(
    "Kategori Selection",
    ["Select All", "Clear All"],
    horizontal=True
)

st.session_state["category_select"] = (
    list(cont_map_cat.keys()) if cat_mode == "Select All" else []
)

st.session_state["category_select"] = st.sidebar.multiselect(
    "Pilih Kategori",
    cont_map_cat.keys(),
    default=st.session_state["category_select"]
)

# format
if "format_select" not in st.session_state:
    st.session_state["format_select"] = list(cont_map_fmt.keys())

fmt_mode = st.sidebar.radio(
    "Format Selection",
    ["Select All", "Clear All"],
    horizontal=True
)

st.session_state["format_select"] = (
    list(cont_map_fmt.keys()) if fmt_mode == "Select All" else []
)

st.session_state["format_select"] = st.sidebar.multiselect(
    "Pilih Format",
    cont_map_fmt.keys(),
    default=st.session_state["format_select"]
)

# ===============================
# KPI CARDS
# ===============================
total_value_mtd = display_df["Value MTD"].astype(float).sum()
total_value_ytd = display_df["Value YTD"].astype(float).sum()
avg_growth = display_df["Growth MTD"].astype(str).str.replace("%", "").astype(float).mean()

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Total Value MTD</div>
        <div class="kpi-value">{total_value_mtd:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Total Value YTD</div>
        <div class="kpi-value">{total_value_ytd:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    color = "green" if avg_growth >= 0 else "red"
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Avg Growth MTD</div>
        <div class="kpi-value" style="color:{color}">{avg_growth:.1f}%</div>
    </div>
    """, unsafe_allow_html=True)

# ===============================
# TABLE SECTION
# ===============================
st.markdown('<div class="section-title">📋 Detail Performance</div>', unsafe_allow_html=True)

def color_growth(val):
    try:
        v = float(str(val).replace("%", ""))
        if v > 0:
            return "color: green"
        elif v < 0:
            return "color: red"
    except:
        pass
    return ""

styled_df = display_df.style \
    .applymap(color_growth, subset=["Growth MTD", "%Gr L3M", "Growth YTD"]) \
    .set_properties(**{"text-align": "center"}) \
    .set_properties(subset=["Produk"], **{"text-align": "left", "color": "#2563eb"})

st.dataframe(styled_df, use_container_width=True, height=520)

# ===============================
# DOWNLOAD
# ===============================
st.divider()
st.download_button(
    "⬇️ Download Excel Report",
    data=output,
    file_name="Metrics_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
