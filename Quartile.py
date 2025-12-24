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
            return 0.0
        if isinstance(val, str):
            return float(val.replace("%","").replace(",","."))
        return float(val)
    except:
        return 0.0

def parse_number(val):
    try:
        if val is None or (isinstance(val, str) and val.strip() == ""):
            return 0
        return float(val)
    except:
        return 0

def load_metrics(file):
    cont = pd.read_excel(file, sheet_name="Sheet 18", usecols=["Product P","% of Total Current DO TP2 along Product P, Product P Hidden"])
    value_mtd = pd.read_excel(file, sheet_name="Sheet 1", usecols=["Product P","Current DO"])
    value_ytd = pd.read_excel(file, sheet_name="Sheet 1", usecols=["Product P","Current DO TP2"])
    growth_mtd = pd.read_excel(file, sheet_name="Sheet 4", skiprows=1, usecols=["Product P","vs LY"])
    growth_l3m = pd.read_excel(file, sheet_name="Sheet 3", skiprows=1, usecols=["Product P","vs L3M"])
    growth_ytd = pd.read_excel(file, sheet_name="Sheet 5", skiprows=1, usecols=["Product P","vs LY"])
    ach_mtd = pd.read_excel(file, sheet_name="Sheet 13", usecols=["Product P","Current Achievement"])
    ach_ytd = pd.read_excel(file, sheet_name="Sheet 14", usecols=["Product P","Current Achievement TP2"])

    df = cont.rename(columns={"% of Total Current DO TP2 along Product P, Product P Hidden":"Cont YTD"})
    df = df.merge(value_mtd.rename(columns={"Current DO":"Value MTD"}), on="Product P", how="left")
    df = df.merge(value_ytd.rename(columns={"Current DO TP2":"Value YTD"}), on="Product P", how="left")
    df = df.merge(growth_mtd.rename(columns={"vs LY":"Growth MTD"}), on="Product P", how="left")
    df = df.merge(growth_l3m.rename(columns={"vs L3M":"%Gr L3M"}), on="Product P", how="left")
    df = df.merge(growth_ytd.rename(columns={"vs LY":"Growth YTD"}), on="Product P", how="left")
    df = df.merge(ach_mtd.rename(columns={"Current Achievement":"Ach MTD"}), on="Product P", how="left")
    df = df.merge(ach_ytd.rename(columns={"Current Achievement TP2":"Ach YTD"}), on="Product P", how="left")

    # parse numbers
    for col in ["Cont YTD","Growth MTD","%Gr L3M","Growth YTD","Ach MTD","Ach YTD"]:
        df[col] = df[col].apply(parse_percent)
    for col in ["Value MTD","Value YTD"]:
        df[col] = df[col].apply(parse_number)

    return df

# ===============================
# LOAD DATA
# ===============================
df_format = load_metrics(format_file)
df_category = load_metrics(category_file)

# ===============================
# FILTERS: SELECT ALL / DESELECT ALL RADIO
# ===============================
st.sidebar.subheader("Filter Kategori")
if "category_select" not in st.session_state:
    st.session_state["category_select"] = df_category["Product P"].tolist()

cat_radio = st.sidebar.radio("Kategori Select All / Deselect All", ["Select All","Deselect All"], key="cat_radio")
if cat_radio=="Select All":
    st.session_state["category_select"] = df_category["Product P"].tolist()
else:
    st.session_state["category_select"] = []

st.session_state["category_select"] = st.sidebar.multiselect(
    "Pilih Kategori", options=df_category["Product P"].tolist(),
    default=st.session_state["category_select"]
)

st.sidebar.subheader("Filter Format")
if "format_select" not in st.session_state:
    st.session_state["format_select"] = df_format["Product P"].tolist()

fmt_radio = st.sidebar.radio("Format Select All / Deselect All", ["Select All","Deselect All"], key="fmt_radio")
if fmt_radio=="Select All":
    st.session_state["format_select"] = df_format["Product P"].tolist()
else:
    st.session_state["format_select"] = []

st.session_state["format_select"] = st.sidebar.multiselect(
    "Pilih Format", options=df_format["Product P"].tolist(),
    default=st.session_state["format_select"]
)

# ===============================
# BUILD DISPLAY DATA
# ===============================
rows = []

# 1️⃣ Kategori di atas
cat_df = df_category[df_category["Product P"].isin(st.session_state["category_select"])]
rows.append(cat_df)

# 2️⃣ Format yang dipilih
fmt_df = df_format[df_format["Product P"].isin(st.session_state["format_select"])]
rows.append(fmt_df)

# 3️⃣ Others = sum metric dari format yang **tidak dipilih**
others_df = df_format[~df_format["Product P"].isin(st.session_state["format_select"])]
if not others_df.empty:
    # SUM metric persentase **langsung dalam angka %**
    others_sum = pd.DataFrame({
        "Product P":["Others"],
        "Cont YTD":[others_df["Cont YTD"].sum()],
        "Value MTD":[others_df["Value MTD"].sum()],
        "Value YTD":[others_df["Value YTD"].sum()],
        "Growth MTD":[others_df["Growth MTD"].sum()],
        "%Gr L3M":[others_df["%Gr L3M"].sum()],
        "Growth YTD":[others_df["Growth YTD"].sum()],
        "Ach MTD":[others_df["Ach MTD"].sum()],
        "Ach YTD":[others_df["Ach YTD"].sum()]
    })
    rows.append(others_sum)

display_df = pd.concat(rows, ignore_index=True)

# ===============================
# FORMAT PERCENTAGE FOR DISPLAY
# ===============================
pct_cols = ["Cont YTD","Growth MTD","%Gr L3M","Growth YTD","Ach MTD","Ach YTD"]

def fmt_pct(x):
    return f"{x:.1f}%" if pd.notna(x) else ""

for col in pct_cols:
    display_df[col] = display_df[col].apply(fmt_pct)

# Styling biru untuk nama unit
def highlight_name(row):
    styles = [""]*len(row)
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

    ws.set_column(0, 0, 40)
    ws.set_column(1, len(display_df.columns)-1, 18)

output.seek(0)
st.download_button(
    "Download Excel",
    output,
    "Metrics_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
