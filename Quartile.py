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
    if pd.isna(val):
        return None
    if isinstance(val, str):
        return round(float(val.replace("%","").replace(",", ".")), 1)
    return round(float(val)*100, 1)

def parse_number(val):
    if pd.isna(val):
        return None
    return round(float(val),0)

def load_metrics(file):
    # Ambil sheet pertama atau sheet spesifik sesuai kebutuhan
    df = pd.read_excel(file, sheet_name=0)
    df = df.dropna(subset=["Product P"])
    # Ambil metrics kolom yang ada
    metrics_cols = [c for c in df.columns if c != "Product P"]
    for col in metrics_cols:
        if "vs" in col or "Achievement" in col or "%Gr" in col:
            df[col] = df[col].apply(parse_percent)
        else:
            df[col] = df[col].apply(parse_number)
    return df

# ===============================
# LOAD DATA
# ===============================
format_df = load_metrics(format_file)
category_df = load_metrics(category_file)

# ===============================
# FILTERS: SELECT ALL / DESELECT ALL RADIO
# ===============================
st.sidebar.subheader("Filter Kategori")
if "category_select" not in st.session_state:
    st.session_state["category_select"] = list(category_df["Product P"].unique())

cat_all_radio = st.sidebar.radio("Kategori Select All / Deselect All", ("Select All", "Deselect All"))
if cat_all_radio == "Select All":
    st.session_state["category_select"] = list(category_df["Product P"].unique())
else:
    st.session_state["category_select"] = []

st.session_state["category_select"] = st.sidebar.multiselect(
    "Pilih Kategori", options=list(category_df["Product P"].unique()),
    default=st.session_state["category_select"]
)

st.sidebar.subheader("Filter Format")
if "format_select" not in st.session_state:
    st.session_state["format_select"] = list(format_df["Product P"].unique())

fmt_all_radio = st.sidebar.radio("Format Select All / Deselect All", ("Select All", "Deselect All"), key="fmt_radio")
if fmt_all_radio == "Select All":
    st.session_state["format_select"] = list(format_df["Product P"].unique())
else:
    st.session_state["format_select"] = []

st.session_state["format_select"] = st.sidebar.multiselect(
    "Pilih Format", options=list(format_df["Product P"].unique()),
    default=st.session_state["format_select"]
)

# ===============================
# BUILD DISPLAY DATA
# ===============================
display_rows = []

# 1️⃣ Kategori selalu di atas
for idx, row in category_df.iterrows():
    if row["Product P"] in st.session_state["category_select"]:
        display_rows.append({
            "Name": row["Product P"],
            **{c: row[c] for c in category_df.columns if c!="Product P"}
        })

# 2️⃣ Format yang dipilih
selected_format_df = format_df[format_df["Product P"].isin(st.session_state["format_select"])]
for idx, row in selected_format_df.iterrows():
    display_rows.append({
        "Name": row["Product P"],
        **{c: row[c] for c in format_df.columns if c!="Product P"}
    })

# 3️⃣ Others (Format saja)
others_df = format_df[~format_df["Product P"].isin(st.session_state["format_select"])]
if not others_df.empty:
    summed = {"Name": "Others"}
    metric_cols = [c for c in format_df.columns if c!="Product P"]
    for c in metric_cols:
        summed[c] = others_df[c].sum(min_count=1)
    display_rows.append(summed)

# ===============================
# CREATE DISPLAY DATAFRAME
# ===============================
display_df = pd.DataFrame(display_rows)

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
                if "%Gr" in display_df.columns[c_idx] or "Achievement" in display_df.columns[c_idx] or "vs" in display_df.columns[c_idx]:
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
