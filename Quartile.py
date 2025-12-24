import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("Metrics Report (Format & Kategori)")

# ===== FILE UPLOAD =====
with st.sidebar.expander("Upload Files", expanded=True):
    format_file = st.file_uploader("Format File", type=["xlsx"])
    category_file = st.file_uploader("Kategori File", type=["xlsx"])

if not format_file or not category_file:
    st.warning("Upload both Format and Kategori files")
    st.stop()

# ===== HELPER =====
def parse_number(val):
    try:
        if pd.isna(val):
            return 0
        return float(val)
    except:
        return 0

def load_data(file):
    df = pd.read_excel(file, sheet_name="Sheet 18", usecols=["Product P","% of Total Current DO TP2 along Product P, Product P Hidden"])
    df = df.rename(columns={"% of Total Current DO TP2 along Product P, Product P Hidden":"Cont YTD"})
    
    # Load other metrics
    df_value_mtd = pd.read_excel(file, sheet_name="Sheet 1", usecols=["Product P","Current DO"]).rename(columns={"Current DO":"Value MTD"})
    df_value_ytd = pd.read_excel(file, sheet_name="Sheet 1", usecols=["Product P","Current DO TP2"]).rename(columns={"Current DO TP2":"Value YTD"})
    df_growth_mtd = pd.read_excel(file, sheet_name="Sheet 4", skiprows=1, usecols=["Product P","vs LY"]).rename(columns={"vs LY":"Growth MTD"})
    df_growth_l3m = pd.read_excel(file, sheet_name="Sheet 3", skiprows=1, usecols=["Product P","vs L3M"]).rename(columns={"vs L3M":"%Gr L3M"})
    df_growth_ytd = pd.read_excel(file, sheet_name="Sheet 5", skiprows=1, usecols=["Product P","vs LY"]).rename(columns={"vs LY":"Growth YTD"})
    df_ach_mtd = pd.read_excel(file, sheet_name="Sheet 13", usecols=["Product P","Current Achievement"]).rename(columns={"Current Achievement":"Ach MTD"})
    df_ach_ytd = pd.read_excel(file, sheet_name="Sheet 14", usecols=["Product P","Current Achievement TP2"]).rename(columns={"Current Achievement TP2":"Ach YTD"})

    # Merge
    df = df.merge(df_value_mtd,on="Product P",how="left")
    df = df.merge(df_value_ytd,on="Product P",how="left")
    df = df.merge(df_growth_mtd,on="Product P",how="left")
    df = df.merge(df_growth_l3m,on="Product P",how="left")
    df = df.merge(df_growth_ytd,on="Product P",how="left")
    df = df.merge(df_ach_mtd,on="Product P",how="left")
    df = df.merge(df_ach_ytd,on="Product P",how="left")

    # Convert numeric columns
    for col in ["Value MTD","Value YTD"]:
        df[col] = df[col].apply(parse_number)

    return df

# ===== LOAD DATA =====
df_format = load_data(format_file)
df_category = load_data(category_file)

# ===== FILTERS =====
st.sidebar.subheader("Filter Kategori")
cat_all = df_category["Product P"].tolist()
if "category_select" not in st.session_state:
    st.session_state["category_select"] = cat_all.copy()

cat_radio = st.sidebar.radio("Kategori Select All / Deselect All", ["Select All","Deselect All"], key="cat_radio")
st.session_state["category_select"] = cat_all.copy() if cat_radio=="Select All" else []

st.session_state["category_select"] = st.sidebar.multiselect("Pilih Kategori", options=cat_all, default=st.session_state["category_select"])

st.sidebar.subheader("Filter Format")
fmt_all = df_format["Product P"].tolist()
if "format_select" not in st.session_state:
    st.session_state["format_select"] = fmt_all.copy()

fmt_radio = st.sidebar.radio("Format Select All / Deselect All", ["Select All","Deselect All"], key="fmt_radio")
st.session_state["format_select"] = fmt_all.copy() if fmt_radio=="Select All" else []

st.session_state["format_select"] = st.sidebar.multiselect("Pilih Format", options=fmt_all, default=st.session_state["format_select"])

# ===== BUILD DISPLAY =====
rows = []

# Kategori di atas
cat_df = df_category[df_category["Product P"].isin(st.session_state["category_select"])]
rows.append(cat_df)

# Format yang dipilih
fmt_df = df_format[df_format["Product P"].isin(st.session_state["format_select"])]
rows.append(fmt_df)

# Others dari format yang tidak dipilih
others_df = df_format[~df_format["Product P"].isin(st.session_state["format_select"])]
if not others_df.empty:
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

# ===== FORMAT PERSEN UNTUK DISPLAY =====
pct_cols = ["Cont YTD","Growth MTD","%Gr L3M","Growth YTD","Ach MTD","Ach YTD"]

def fmt_pct(x):
    try:
        return f"{float(x):.1f}%" if pd.notna(x) else ""
    except:
        return ""

for col in pct_cols:
    display_df[col] = display_df[col].apply(fmt_pct)

# Styling biru
def highlight_name(row):
    styles = [""]*len(row)
    styles[0] = "color: blue"
    return styles

st.dataframe(display_df.style.apply(highlight_name, axis=1), use_container_width=True)

# ===== DOWNLOAD EXCEL =====
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

    for col_idx, col in enumerate(display_df.columns):
        ws.write(0, col_idx, col, header_fmt)

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
