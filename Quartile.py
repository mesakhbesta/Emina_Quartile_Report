import streamlit as st
import pandas as pd
from io import BytesIO

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    layout="wide",
    page_title="Dynamic Metrics Report"
)

st.title("📊 Dynamic Metrics Report Category & Format")
st.caption("Format & Kategori Performance Overview")
st.divider()

# =====================================================
# FILE UPLOAD
# =====================================================
with st.sidebar:
    st.header("📁 Upload Data")
    format_file = st.file_uploader("Format File (.xlsx)", type=["xlsx"])
    category_file = st.file_uploader("Kategori File (.xlsx)", type=["xlsx"])

if not format_file or not category_file:
    st.warning("⚠️ Silakan upload **Format File** dan **Kategori File**.")
    st.stop()

# =====================================================
# HELPER FUNCTIONS
# =====================================================
def parse_percent(val):
    try:
        if val is None or (isinstance(val, str) and val.strip() == ""):
            return 0
        if isinstance(val, str):
            return round(float(val.replace("%", "").replace(",", ".")), 1)
        return round(float(val) * 100, 1)
    except:
        return 0

def parse_number(val):
    try:
        if val is None or (isinstance(val, str) and val.strip() == ""):
            return 0
        return round(float(val), 0)
    except:
        return 0

def load_map(sheet, key_col, val_col, file, skip=0, parser=None):
    df = pd.read_excel(file, sheet_name=sheet, skiprows=skip)
    df = df.dropna(subset=[key_col])
    result = {}
    for _, r in df.iterrows():
        v = r[val_col] if val_col in r else 0
        if parser:
            v = parser(v)
        result[r[key_col]] = v
    return result

# =====================================================
# LOAD FORMAT METRICS
# =====================================================
cont_fmt = load_map(
    "Sheet 18",
    "Product P",
    "% of Total Current DO TP2 along Product P, Product P Hidden",
    format_file,
    parser=parse_percent
)

val_mtd_fmt = load_map("Sheet 1", "Product P", "Current DO", format_file, parser=parse_number)
val_ytd_fmt = load_map("Sheet 1", "Product P", "Current DO TP2", format_file, parser=parse_number)
gr_mtd_fmt = load_map("Sheet 4", "Product P", "vs LY", format_file, skip=1, parser=parse_percent)
gr_l3m_fmt = load_map("Sheet 3", "Product P", "vs L3M", format_file, skip=1, parser=parse_percent)
gr_ytd_fmt = load_map("Sheet 5", "Product P", "vs LY", format_file, skip=1, parser=parse_percent)
ach_mtd_fmt = load_map("Sheet 13", "Product P", "Current Achievement", format_file, parser=parse_percent)
ach_ytd_fmt = load_map("Sheet 14", "Product P", "Current Achievement TP2", format_file, parser=parse_percent)

# =====================================================
# LOAD CATEGORY METRICS
# =====================================================
cont_cat = load_map(
    "Sheet 18",
    "Product P",
    "% of Total Current DO TP2 along Product P, Product P Hidden",
    category_file,
    parser=parse_percent
)

val_mtd_cat = load_map("Sheet 1", "Product P", "Current DO", category_file, parser=parse_number)
val_ytd_cat = load_map("Sheet 1", "Product P", "Current DO TP2", category_file, parser=parse_number)
gr_mtd_cat = load_map("Sheet 4", "Product P", "vs LY", category_file, skip=1, parser=parse_percent)
gr_l3m_cat = load_map("Sheet 3", "Product P", "vs L3M", category_file, skip=1, parser=parse_percent)
gr_ytd_cat = load_map("Sheet 5", "Product P", "vs LY", category_file, skip=1, parser=parse_percent)
ach_mtd_cat = load_map("Sheet 13", "Product P", "Current Achievement", category_file, parser=parse_percent)
ach_ytd_cat = load_map("Sheet 14", "Product P", "Current Achievement TP2", category_file, parser=parse_percent)

# =====================================================
# FILTERS WITH RADIO (SAFE VERSION)
# =====================================================
with st.sidebar:
    st.header("🎯 Filter Data")

    # ---------- KATEGORI ----------
    if "cat_select" not in st.session_state:
        st.session_state.cat_select = list(cont_cat.keys())

    cat_mode = st.radio(
        "Kategori Selection",
        ["Select All", "Clear All"],
        key="cat_mode"
    )

    if cat_mode == "Select All":
        st.session_state.cat_select = list(cont_cat.keys())
    else:
        st.session_state.cat_select = []

    st.session_state.cat_select = st.multiselect(
        "Pilih Kategori",
        options=list(cont_cat.keys()),
        default=st.session_state.cat_select
    )

    st.divider()

    # ---------- FORMAT ----------
    if "fmt_select" not in st.session_state:
        st.session_state.fmt_select = list(cont_fmt.keys())

    fmt_mode = st.radio(
        "Format Selection",
        ["Select All", "Clear All"],
        key="fmt_mode"
    )

    if fmt_mode == "Select All":
        st.session_state.fmt_select = list(cont_fmt.keys())
    else:
        st.session_state.fmt_select = []

    st.session_state.fmt_select = st.multiselect(
        "Pilih Format",
        options=list(cont_fmt.keys()),
        default=st.session_state.fmt_select
    )

# =====================================================
# BUILD DISPLAY DATA (NO OTHERS)
# =====================================================
rows = []

for k in st.session_state.cat_select:
    rows.append([
        k,
        cont_cat.get(k, 0),
        val_mtd_cat.get(k, 0),
        val_ytd_cat.get(k, 0),
        gr_mtd_cat.get(k, 0),
        gr_l3m_cat.get(k, 0),
        gr_ytd_cat.get(k, 0),
        ach_mtd_cat.get(k, 0),
        ach_ytd_cat.get(k, 0),
    ])

for f in st.session_state.fmt_select:
    rows.append([
        f,
        cont_fmt.get(f, 0),
        val_mtd_fmt.get(f, 0),
        val_ytd_fmt.get(f, 0),
        gr_mtd_fmt.get(f, 0),
        gr_l3m_fmt.get(f, 0),
        gr_ytd_fmt.get(f, 0),
        ach_mtd_fmt.get(f, 0),
        ach_ytd_fmt.get(f, 0),
    ])

# =====================================================
# DATAFRAME & DISPLAY
# =====================================================
df = pd.DataFrame(rows, columns=[
    "Produk",
    "Cont YTD",
    "Value MTD",
    "Value YTD",
    "Growth MTD",
    "%Gr L3M",
    "Growth YTD",
    "Ach MTD",
    "Ach YTD"
])

pct_cols = ["Cont YTD", "Growth MTD", "%Gr L3M", "Growth YTD", "Ach MTD", "Ach YTD"]
for c in pct_cols:
    df[c] = df[c].apply(lambda x: f"{x:.1f}%")

st.subheader("📈 Performance Table")
st.dataframe(
    df.style.apply(
        lambda _: ["color: #1f77b4"] + [""] * (len(df.columns) - 1),
        axis=1
    ),
    use_container_width=True
)

# =====================================================
# DOWNLOAD EXCEL
# =====================================================
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Report", index=False)

output.seek(0)

st.download_button(
    "⬇️ Download Excel",
    output,
    "Metrics_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


