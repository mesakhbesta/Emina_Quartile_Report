import streamlit as st
import pandas as pd
from io import BytesIO

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(layout="wide", page_title="Dynamic Metrics Report")
st.title("📊 Dynamic Metrics Report Category & Format")
st.caption("Format & Kategori Performance Overview")
st.divider()

# =====================================================
# INIT SESSION STATE (WAJIB – BIAR GA RESET)
# =====================================================
if "lock_cat" not in st.session_state:
    st.session_state.lock_cat = False
if "lock_fmt" not in st.session_state:
    st.session_state.lock_fmt = False
if "cat_select" not in st.session_state:
    st.session_state.cat_select = []
if "fmt_select" not in st.session_state:
    st.session_state.fmt_select = []

# =====================================================
# FILE UPLOAD
# =====================================================
with st.sidebar:
    st.header("📁 Upload Data")
    format_file = st.file_uploader("Format File (.xlsx)", type=["xlsx"])
    category_file = st.file_uploader("Kategori File (.xlsx)", type=["xlsx"])

if not format_file or not category_file:
    st.warning("⚠️ Upload **Format File** dan **Kategori File**")
    st.stop()

# =====================================================
# HELPERS
# =====================================================
def parse_percent(val):
    try:
        if val is None or val == "":
            return 0
        if isinstance(val, str):
            return round(float(val.replace("%", "").replace(",", ".")), 1)
        return round(float(val) * 100, 1)
    except:
        return 0

def parse_number(val):
    try:
        if val is None or val == "":
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
        result[r[key_col]] = parser(v) if parser else v
    return result

def sanitize_selection(old, options, lock):
    if lock:
        return old
    return [x for x in old if x in options]

# =====================================================
# LOAD DATA
# =====================================================
cont_cat = load_map(
    "Sheet 18", "Product P",
    "% of Total Current DO TP2 along Product P, Product P Hidden",
    category_file, parser=parse_percent
)

cont_fmt = load_map(
    "Sheet 18", "Product P",
    "% of Total Current DO TP2 along Product P, Product P Hidden",
    format_file, parser=parse_percent
)

val_mtd_cat = load_map("Sheet 1", "Product P", "Current DO", category_file, parser=parse_number)
val_ytd_cat = load_map("Sheet 1", "Product P", "Current DO TP2", category_file, parser=parse_number)

val_mtd_fmt = load_map("Sheet 1", "Product P", "Current DO", format_file, parser=parse_number)
val_ytd_fmt = load_map("Sheet 1", "Product P", "Current DO TP2", format_file, parser=parse_number)

# =====================================================
# SIDEBAR FILTERS (🔒 LOCK STABLE)
# =====================================================
with st.sidebar:
    st.header("🎯 Filter Data")

    # ---------- KATEGORI ----------
    lock_cat = st.toggle(
        "🔒 Lock Kategori",
        value=st.session_state.lock_cat,
        key="lock_cat",
        help="Jika ON, pilihan kategori tidak berubah saat upload file"
    )

    cat_options = list(cont_cat.keys())

    # INIT DEFAULT JIKA KOSONG
    if not st.session_state.cat_select:
        st.session_state.cat_select = cat_options.copy()

    # SANITIZE (KECUALI LOCK)
    st.session_state.cat_select = sanitize_selection(
        st.session_state.cat_select, cat_options, lock_cat
    )

    cat_mode = st.radio(
        "Kategori Mode",
        ["Manual", "Select All", "Clear All"],
        index=0,
        disabled=lock_cat
    )

    if not lock_cat:
        if cat_mode == "Select All":
            st.session_state.cat_select = cat_options.copy()
        elif cat_mode == "Clear All":
            st.session_state.cat_select = []

    st.session_state.cat_select = st.multiselect(
        "Pilih Kategori",
        options=cat_options,
        default=st.session_state.cat_select,
        disabled=lock_cat
    )

    st.divider()

    # ---------- FORMAT ----------
    lock_fmt = st.toggle(
        "🔒 Lock Format",
        value=st.session_state.lock_fmt,
        key="lock_fmt",
        help="Jika ON, pilihan format tidak berubah saat upload file"
    )

    fmt_options = list(cont_fmt.keys())

    if not st.session_state.fmt_select:
        st.session_state.fmt_select = fmt_options.copy()

    st.session_state.fmt_select = sanitize_selection(
        st.session_state.fmt_select, fmt_options, lock_fmt
    )

    fmt_mode = st.radio(
        "Format Mode",
        ["Manual", "Select All", "Clear All"],
        index=0,
        disabled=lock_fmt
    )

    if not lock_fmt:
        if fmt_mode == "Select All":
            st.session_state.fmt_select = fmt_options.copy()
        elif fmt_mode == "Clear All":
            st.session_state.fmt_select = []

    st.session_state.fmt_select = st.multiselect(
        "Pilih Format",
        options=fmt_options,
        default=st.session_state.fmt_select,
        disabled=lock_fmt
    )

# =====================================================
# BUILD TABLE
# =====================================================
rows = []

for k in st.session_state.cat_select:
    rows.append([
        k,
        cont_cat.get(k, 0),
        val_mtd_cat.get(k, 0),
        val_ytd_cat.get(k, 0)
    ])

for f in st.session_state.fmt_select:
    rows.append([
        f,
        cont_fmt.get(f, 0),
        val_mtd_fmt.get(f, 0),
        val_ytd_fmt.get(f, 0)
    ])

df = pd.DataFrame(rows, columns=[
    "Produk", "Contribution YTD", "Value MTD", "Value YTD"
])

df["Contribution YTD"] = df["Contribution YTD"].apply(lambda x: f"{x:.1f}%")

# =====================================================
# DISPLAY
# =====================================================
st.subheader("📈 Performance Table")
st.dataframe(df, use_container_width=True)

# =====================================================
# DOWNLOAD
# =====================================================
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    df.to_excel(writer, index=False)

st.download_button(
    "⬇️ Download Excel",
    output.getvalue(),
    "Metrics_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
