import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("Dynamic Metrics Report (Format & Kategori)")

# ===============================
# FILE UPLOAD
# ===============================
with st.sidebar.expander("Upload Excel Files", expanded=True):
    format_file = st.file_uploader("Format File", type=["xlsx"])
    category_file = st.file_uploader("Kategori File", type=["xlsx"])

if not format_file or not category_file:
    st.warning("Please upload both Format and Kategori files")
    st.stop()

# ===============================
# HELPER FUNCTIONS
# ===============================
def parse_percent(val):
    """Parse persentase Excel ke float, hapus %"""
    if pd.isna(val) or str(val).strip() == "":
        return 0.0
    val = str(val).replace("%", "").replace(",", ".")
    try:
        return float(val)
    except:
        return 0.0

def parse_number(val):
    if pd.isna(val) or str(val).strip() == "":
        return 0
    try:
        return float(val)
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

# ===============================
# LOAD METRICS FORMAT
# ===============================
cont_map_fmt = load_map("Sheet 18", "Product P", "% of Total Current DO TP2 along Product P, Product P Hidden", format_file, parser=parse_percent)
value_mtd_fmt = load_map("Sheet 1", "Product P", "Current DO", format_file, parser=parse_number)
value_ytd_fmt = load_map("Sheet 1", "Product P", "Current DO TP2", format_file, parser=parse_number)
growth_mtd_fmt = load_map("Sheet 4", "Product P", "vs LY", skip=1, format_file, parser=parse_percent)
growth_l3m_fmt = load_map("Sheet 3", "Product P", "vs L3M", skip=1, format_file, parser=parse_percent)
growth_ytd_fmt = load_map("Sheet 5", "Product P", "vs LY", skip=1, format_file, parser=parse_percent)
ach_mtd_fmt = load_map("Sheet 13", "Product P", "Current Achievement", format_file, parser=parse_percent)
ach_ytd_fmt = load_map("Sheet 14", "Product P", "Current Achievement TP2", format_file, parser=parse_percent)

# ===============================
# LOAD METRICS CATEGORY
# ===============================
cont_map_cat = load_map("Sheet 18", "Product P", "% of Total Current DO TP2 along Product P, Product P Hidden", category_file, parser=parse_percent)
value_mtd_cat = load_map("Sheet 1", "Product P", "Current DO", category_file, parser=parse_number)
value_ytd_cat = load_map("Sheet 1", "Product P", "Current DO TP2", category_file, parser=parse_number)
growth_mtd_cat = load_map("Sheet 4", "Product P", "vs LY", skip=1, category_file, parser=parse_percent)
growth_l3m_cat = load_map("Sheet 3", "Product P", "vs L3M", skip=1, category_file, parser=parse_percent)
growth_ytd_cat = load_map("Sheet 5", "Product P", "vs LY", skip=1, category_file, parser=parse_percent)
ach_mtd_cat = load_map("Sheet 13", "Product P", "Current Achievement", category_file, parser=parse_percent)
ach_ytd_cat = load_map("Sheet 14", "Product P", "Current Achievement TP2", category_file, parser=parse_percent)

# ===============================
# FILTERS
# ===============================
st.sidebar.subheader("Filter Kategori")
if "category_select" not in st.session_state:
    st.session_state["category_select"] = list(cont_map_cat.keys())

cat_all_radio = st.sidebar.radio("Kategori Select All / Deselect All", ("Select All", "Deselect All"), key="cat_radio")
if cat_all_radio == "Select All":
    st.session_state["category_select"] = list(cont_map_cat.keys())
else:
    st.session_state["category_select"] = []

st.session_state["category_select"] = st.sidebar.multiselect(
    "Pilih Kategori", options=list(cont_map_cat.keys()),
    default=st.session_state["category_select"]
)

st.sidebar.subheader("Filter Format")
if "format_select" not in st.session_state:
    st.session_state["format_select"] = list(cont_map_fmt.keys())

fmt_all_radio = st.sidebar.radio("Format Select All / Deselect All", ("Select All", "Deselect All"), key="fmt_radio")
if fmt_all_radio == "Select All":
    st.session_state["format_select"] = list(cont_map_fmt.keys())
else:
    st.session_state["format_select"] = []

st.session_state["format_select"] = st.sidebar.multiselect(
    "Pilih Format", options=list(cont_map_fmt.keys()),
    default=st.session_state["format_select"]
)

# ===============================
# BUILD DISPLAY DATA
# ===============================
rows = []

# 1. Kategori selalu di atas
for k in st.session_state["category_select"]:
    rows.append([
        k,
        cont_map_cat.get(k,0),
        value_mtd_cat.get(k,0),
        value_ytd_cat.get(k,0),
        growth_mtd_cat.get(k,0),
        growth_l3m_cat.get(k,0),
        growth_ytd_cat.get(k,0),
        ach_mtd_cat.get(k,0),
        ach_ytd_cat.get(k,0)
    ])

# 2. Format yang dipilih
selected_fmt = st.session_state["format_select"]
for f in selected_fmt:
    rows.append([
        f,
        cont_map_fmt.get(f,0),
        value_mtd_fmt.get(f,0),
        value_ytd_fmt.get(f,0),
        growth_mtd_fmt.get(f,0),
        growth_l3m_fmt.get(f,0),
        growth_ytd_fmt.get(f,0),
        ach_mtd_fmt.get(f,0),
        ach_ytd_fmt.get(f,0)
    ])

# 3. Others = sum dari yang tidak dipilih
others_keys = [k for k in cont_map_fmt.keys() if k not in selected_fmt]
if others_keys:
    summed = ["Others"]
    summed.append(sum([cont_map_fmt[k] for k in others_keys]))
    summed.append(sum([value_mtd_fmt[k] for k in others_keys]))
    summed.append(sum([value_ytd_fmt[k] for k in others_keys]))
    summed.append(sum([growth_mtd_fmt[k] for k in others_keys]))
    summed.append(sum([growth_l3m_fmt[k] for k in others_keys]))
    summed.append(sum([growth_ytd_fmt[k] for k in others_keys]))
    summed.append(sum([ach_mtd_fmt[k] for k in others_keys]))
    summed.append(sum([ach_ytd_fmt[k] for k in others_keys]))
    rows.append(summed)

# ===============================
# CREATE DISPLAY DF
# ===============================
display_df = pd.DataFrame(rows, columns=[
    "Produk","Cont YTD","Value MTD","Value YTD",
    "Growth MTD","%Gr L3M","Growth YTD","Ach MTD","Ach YTD"
])

# Format persentase hanya untuk display
pct_cols = ["Cont YTD","Growth MTD","%Gr L3M","Growth YTD","Ach MTD","Ach YTD"]
for col in pct_cols:
    display_df[col] = display_df[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "")

# Styling biru untuk nama unit
def highlight_name(row):
    styles = [""] * len(row)
    styles[0] = "color: blue"
    return styles

st.dataframe(display_df.style.apply(highlight_name, axis=1), use_container_width=True)

# ===============================
# DOWNLOAD EXCEL
# ===============================
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    wb = writer.book
    ws = wb.add_worksheet("Report")
    writer.sheets["Report"] = ws

    header_fmt = wb.add_format({"bold": True, "align":"center","border":1})
    name_fmt = wb.add_format({"bold": False, "font_color":"blue","border":1})
    num_fmt = wb.add_format({"border":1, "num_format":"#,##0"})
    pct_g = wb.add_format({"border":1,"num_format":"0.0%","font_color":"green"})
    pct_r = wb.add_format({"border":1,"num_format":"0.0%","font_color":"red"})

    # Header
    for col_idx, col in enumerate(display_df.columns):
        ws.write(0, col_idx, col, header_fmt)

    # Rows
    for r_idx, row in enumerate(display_df.itertuples(index=False), start=1):
        ws.write(r_idx, 0, row[0], name_fmt)
        for c_idx, val in enumerate(row[1:], start=1):
            if pd.isna(val) or val=="":
                ws.write_blank(r_idx, c_idx, None, num_fmt)
            else:
                try:
                    val_num = float(str(val).replace("%",""))
                except:
                    val_num = val
                if display_df.columns[c_idx] in pct_cols:
                    ws.write_number(r_idx, c_idx, val_num/100, pct_g if val_num>=0 else pct_r)
                else:
                    ws.write_number(r_idx, c_idx, val_num, num_fmt)

    ws.set_column(0,0,40)
    ws.set_column(1,len(display_df.columns)-1,18)

output.seek(0)
st.download_button(
    "Download Excel",
    output,
    "Metrics_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
