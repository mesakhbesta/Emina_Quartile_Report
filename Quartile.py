import streamlit as st
import pandas as pd
from io import BytesIO

# ===============================
# PAGE CONFIG
# ===============================
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
    try:
        if val is None or (isinstance(val, str) and val.strip() == ""):
            return None
        if isinstance(val, str):
            return round(float(val.replace("%","").replace(",", ".")), 1)
        return round(float(val)*100, 1)
    except:
        return None

def parse_number(val):
    try:
        if val is None or (isinstance(val, str) and val.strip() == ""):
            return None
        return round(float(val),0)
    except:
        return None

def load_map(sheet, key_col, val_col, file, skip=0, parser=None):
    tmp = pd.read_excel(file, sheet_name=sheet, skiprows=skip)
    tmp = tmp.dropna(subset=[key_col])
    result = {}
    for _, r in tmp.iterrows():
        v = r[val_col] if val_col in r else None
        if parser:
            v = parser(v)
        result[r[key_col]] = v
    return result

# ===============================
# LOAD FORMAT METRICS
# ===============================
cont_map_fmt = load_map("Sheet 18", "Product P", "% of Total Current DO TP2 along Product P, Product P Hidden", file=format_file, parser=parse_percent)
value_mtd_fmt = load_map("Sheet 1", "Product P", "Current DO", file=format_file, parser=parse_number)
value_ytd_fmt = load_map("Sheet 1", "Product P", "Current DO TP2", file=format_file, parser=parse_number)
growth_mtd_fmt = load_map("Sheet 4", "Product P", "vs LY", skip=1, file=format_file, parser=parse_percent)
growth_l3m_fmt = load_map("Sheet 3", "Product P", "vs L3M", skip=1, file=format_file, parser=parse_percent)
growth_ytd_fmt = load_map("Sheet 5", "Product P", "vs LY", skip=1, file=format_file, parser=parse_percent)
ach_mtd_fmt = load_map("Sheet 13", "Product P", "Current Achievement", file=format_file, parser=parse_percent)
ach_ytd_fmt = load_map("Sheet 14", "Product P", "Current Achievement TP2", file=format_file, parser=parse_percent)

# ===============================
# LOAD CATEGORY METRICS
# ===============================
cont_map_cat = load_map("Sheet 18", "Product P", "% of Total Current DO TP2 along Product P, Product P Hidden", file=category_file, parser=parse_percent)
value_mtd_cat = load_map("Sheet 1", "Product P", "Current DO", file=category_file, parser=parse_number)
value_ytd_cat = load_map("Sheet 1", "Product P", "Current DO TP2", file=category_file, parser=parse_number)
growth_mtd_cat = load_map("Sheet 4", "Product P", "vs LY", skip=1, file=category_file, parser=parse_percent)
growth_l3m_cat = load_map("Sheet 3", "Product P", "vs L3M", skip=1, file=category_file, parser=parse_percent)
growth_ytd_cat = load_map("Sheet 5", "Product P", "vs LY", skip=1, file=category_file, parser=parse_percent)
ach_mtd_cat = load_map("Sheet 13", "Product P", "Current Achievement", file=category_file, parser=parse_percent)
ach_ytd_cat = load_map("Sheet 14", "Product P", "Current Achievement TP2", file=category_file, parser=parse_percent)

# ===============================
# FILTERS: SELECT ALL / DESELECT ALL RADIO
# ===============================
st.sidebar.subheader("Filter Kategori")
if "category_select" not in st.session_state:
    st.session_state["category_select"] = list(cont_map_cat.keys())

cat_all_radio = st.sidebar.radio("Kategori Select All / Deselect All", ("Select All", "Deselect All"))
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

# 1️⃣ Kategori di atas
for k in st.session_state["category_select"]:
    rows.append([
        k,
        cont_map_cat.get(k),
        value_mtd_cat.get(k),
        value_ytd_cat.get(k),
        growth_mtd_cat.get(k),
        growth_l3m_cat.get(k),
        growth_ytd_cat.get(k),
        ach_mtd_cat.get(k),
        ach_ytd_cat.get(k)
    ])

# 2️⃣ Format yang dipilih
for f in st.session_state["format_select"]:
    rows.append([
        f,
        cont_map_fmt.get(f),
        value_mtd_fmt.get(f),
        value_ytd_fmt.get(f),
        growth_mtd_fmt.get(f),
        growth_l3m_fmt.get(f),
        growth_ytd_fmt.get(f),
        ach_mtd_fmt.get(f),
        ach_ytd_fmt.get(f)
    ])

# 3️⃣ Others (Format yang tidak dipilih)
others_keys = [k for k in cont_map_fmt.keys() if k not in st.session_state["format_select"]]
if others_keys:
    summed = ["Others"]
    summed.append(sum([v for k in others_keys if (v:=cont_map_fmt.get(k)) is not None]))
    summed.append(sum([v for k in others_keys if (v:=value_mtd_fmt.get(k)) is not None]))
    summed.append(sum([v for k in others_keys if (v:=value_ytd_fmt.get(k)) is not None]))
    summed.append(sum([v for k in others_keys if (v:=growth_mtd_fmt.get(k)) is not None]))
    summed.append(sum([v for k in others_keys if (v:=growth_l3m_fmt.get(k)) is not None]))
    summed.append(sum([v for k in others_keys if (v:=growth_ytd_fmt.get(k)) is not None]))
    summed.append(sum([v for k in others_keys if (v:=ach_mtd_fmt.get(k)) is not None]))
    summed.append(sum([v for k in others_keys if (v:=ach_ytd_fmt.get(k)) is not None]))
    rows.append(summed)

# ===============================
# DISPLAY DATAFRAME
# ===============================
display_df = pd.DataFrame(rows, columns=[
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

# Styling: Name biru
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

    # Write header
    for col_idx, col in enumerate(display_df.columns):
        ws.write(0, col_idx, col, header_fmt)

    # Write rows
    for r_idx, row in enumerate(display_df.itertuples(index=False), start=1):
        ws.write(r_idx, 0, row[0], name_fmt)
        for c_idx, val in enumerate(row[1:], start=1):
            if pd.isna(val):
                ws.write_blank(r_idx, c_idx, None, num_fmt)
            else:
                # Sesuaikan format angka / persentase
                if display_df.columns[c_idx] in ["Cont YTD", "Growth MTD", "%Gr L3M", "Growth YTD", "Ach MTD", "Ach YTD"]:
                    if val >=0:
                        ws.write_number(r_idx, c_idx, val/100, pct_g)
                    else:
                        ws.write_number(r_idx, c_idx, val/100, pct_r)
                else:
                    ws.write_number(r_idx, c_idx, val, num_fmt)

    ws.set_column(0, 0, 40)
    ws.set_column(1, len(display_df.columns)-1, 18)

output.seek(0)
st.download_button(
    "Download Excel",
    output,
    "Metrics_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
