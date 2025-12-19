import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import xlsxwriter
from datetime import datetime

st.set_page_config(layout="wide")
st.title("Quartile Report Emina")

# =========================
# Sidebar: Upload + Cut Off
# =========================
cut_off_date = st.sidebar.date_input("📅 Tanggal Cut Off", value=datetime.today())

with st.sidebar.expander("📤 Upload FORMAT", expanded=True):
    format_files = {
        "Q1": st.file_uploader("Format Q1", type="xlsx"),
        "Q2": st.file_uploader("Format Q2", type="xlsx"),
        "Q3": st.file_uploader("Format Q3", type="xlsx"),
        "Q4": st.file_uploader("Format Q4", type="xlsx"),
    }

with st.sidebar.expander("📤 Upload KATEGORI", expanded=True):
    category_files = {
        "Q1": st.file_uploader("Kategori Q1", type="xlsx"),
        "Q2": st.file_uploader("Kategori Q2", type="xlsx"),
        "Q3": st.file_uploader("Kategori Q3", type="xlsx"),
        "Q4": st.file_uploader("Kategori Q4", type="xlsx"),
    }

# =========================
# Helper Functions
# =========================
def safe_read_excel(file, sheet, skiprows=0):
    """Baca sheet Excel aman, pakai openpyxl"""
    try:
        xls = pd.ExcelFile(file, engine='openpyxl')
        if sheet not in xls.sheet_names:
            st.warning(f"⚠️ Sheet {sheet} tidak ditemukan.")
            return pd.DataFrame()
        df = pd.read_excel(file, sheet_name=sheet, skiprows=skiprows, engine='openpyxl')
        df.columns = df.columns.astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Gagal baca sheet {sheet}: {e}")
        return pd.DataFrame()

def pick_highest_q(files):
    for q in ["Q4", "Q3", "Q2", "Q1"]:
        if files.get(q):
            return files[q]
    return None

def process_files(files, is_category=False):
    if not any(files.values()):
        return pd.DataFrame()

    result = {}
    # YTD
    for q, col in zip(["Q1","Q2","Q3","Q4"], ["YTD Q1","YTD Q2","YTD Q3","YTD Q4"]):
        if files.get(q):
            df = safe_read_excel(files[q], "Sheet 5", skiprows=1)
            if df.empty:
                continue
            key = next((c for c in df.columns if "Product" in c), None)
            val = next((c for c in df.columns if "vs LY" in c), None)
            if key and val:
                df[key] = df[key].astype(str)
                df = df[~df[key].str.lower().eq("grand total")]
                if is_category:
                    df = df[~df[key].str.lower().eq("others")]
                result[col] = df.set_index(key)[val]

    latest = pick_highest_q(files)
    if not latest:
        return pd.DataFrame()

    # CONT
    df = safe_read_excel(latest, "Sheet 18", skiprows=0)
    if not df.empty:
        key = next((c for c in df.columns if "Product" in c), None)
        val = next((c for c in df.columns if "Total Current DO TP2" in c), None)
        if key and val:
            df[key] = df[key].astype(str)
            df = df[~df[key].str.lower().eq("grand total")]
            if is_category:
                df = df[~df[key].str.lower().eq("others")]
            result["Cont"] = df.set_index(key)[val]

    # MTD
    df = safe_read_excel(latest, "Sheet 4", skiprows=1)
    if not df.empty:
        key = next((c for c in df.columns if "Product" in c), None)
        val = next((c for c in df.columns if "vs LY" in c), None)
        if key and val:
            df[key] = df[key].astype(str)
            df = df[~df[key].str.lower().eq("grand total")]
            if is_category:
                df = df[~df[key].str.lower().eq("others")]
            result["MTD"] = df.set_index(key)[val]

    # %Gr L3M
    df = safe_read_excel(latest, "Sheet 3", skiprows=1)
    if not df.empty:
        key = next((c for c in df.columns if "Product" in c), None)
        val = next((c for c in df.columns if "vs L3M" in c), None)
        if key and val:
            df[key] = df[key].astype(str)
            df = df[~df[key].str.lower().eq("grand total")]
            if is_category:
                df = df[~df[key].str.lower().eq("others")]
            result["%Gr L3M MTD"] = df.set_index(key)[val]

    return pd.DataFrame(result).fillna(0)

# =========================
# Proses hanya jika ada file di-upload
if any(format_files.values()) or any(category_files.values()):
    df_format = process_files(format_files, is_category=False)
    df_category = process_files(category_files, is_category=True)
else:
    st.info("⬅️ Upload minimal satu file (Format atau Kategori)")
    st.stop()

# =========================
# Debug: cek sheet & kolom
st.write("Debug: Format DataFrame")
st.write(df_format)
st.write("Debug: Kategori DataFrame")
st.write(df_category)

# =========================
# Filter Selection
format_options = df_format.index.astype(str).tolist() if not df_format.empty else ["-Data Kosong-"]
category_options = df_category.index.astype(str).tolist() if not df_category.empty else ["-Data Kosong-"]

selected_format = st.sidebar.multiselect("Filter Format", format_options)
selected_category = st.sidebar.multiselect("Filter Kategori", category_options)

if (selected_format == ["-Data Kosong-"]) or (selected_category == ["-Data Kosong-"]):
    st.warning("⚠️ Data kosong, silakan cek file yang di-upload dan sheet/kolomnya")
    st.stop()

if not selected_format and not selected_category:
    st.info("⬅️ Silakan pilih minimal satu filter (Format atau Kategori) untuk menampilkan data")
    st.stop()

# =========================
# Siapkan Data Format (tambahkan "Others" jika perlu)
df_fmt_final = pd.DataFrame()
if selected_format and not df_format.empty:
    selected_df = df_format.loc[selected_format]
    others_df = df_format.drop(selected_format, errors="ignore")
    if not others_df.empty:
        others = others_df.sum().to_frame().T
        others.index = ["Others"]
        df_fmt_final = pd.concat([selected_df, others])
    else:
        df_fmt_final = selected_df

# =========================
# Siapkan Data Kategori
df_cat_final = pd.DataFrame()
if selected_category and not df_category.empty:
    df_cat_final = df_category.loc[selected_category]

# =========================
# Merge Display
display_frames = []
if not df_cat_final.empty:
    df_cat_display = df_cat_final.copy()
    df_cat_display.index = ["Kategori - " + str(i) for i in df_cat_display.index]
    display_frames.append(df_cat_display)
if not df_fmt_final.empty:
    df_fmt_display = df_fmt_final.copy()
    df_fmt_display.index = ["Format - " + str(i) for i in df_fmt_display.index]
    display_frames.append(df_fmt_display)

df_final_display = pd.concat(display_frames)

# =========================
# MultiIndex Columns
columns = pd.MultiIndex.from_tuples([
    ("Sell In YTD", "Cont"),
    ("Growth", "YTD Q1"),
    ("Growth", "YTD Q2"),
    ("Growth", "YTD Q3"),
    ("Growth", "MTD"),
    ("Growth", "%Gr L3M MTD"),
    ("Growth", "YTD Q4"),
])
df_display = pd.DataFrame(index=df_final_display.index, columns=columns)
for c in df_final_display.columns:
    if c == "Cont":
        df_display[("Sell In YTD","Cont")] = df_final_display[c]/100
    elif c in df_display.columns.get_level_values(1):
        df_display[("Growth",c)] = df_final_display[c]/100

df_display = df_display.applymap(lambda x: f"{x*100:.2f}%" if pd.notnull(x) else "0.00%")

# =========================
# Cut-off tanggal Bahasa Indonesia
bulan_id = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}
cut_off_str = f"{cut_off_date.day} {bulan_id[cut_off_date.month]} {cut_off_date.year}"

st.markdown(f"**Cut Off: {cut_off_str}**")
st.dataframe(df_display, use_container_width=True)

# =========================
# Download Excel
def to_excel(df, cut_off_str):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'nan_inf_to_errors': True})
    ws = wb.add_worksheet("Report")

    header = wb.add_format({'bold': True, 'align': 'center', 'border': 1})
    cell = wb.add_format({'align': 'center', 'border': 1})
    percent_fmt = wb.add_format({'align': 'center','border':1,'num_format':'0.00%'})
    cut_off_fmt = wb.add_format({'bold': True, 'align':'left'})

    ws.write(0, 0, f"Cut Off: {cut_off_str}", cut_off_fmt)
    ws.write(1, 0, "Format/Kategori", header)
    ws.write(1, 1, "Sell In YTD", header)
    ws.merge_range(1,2,1,7,"Growth", header)

    headers = ["Cont","YTD Q1","YTD Q2","YTD Q3","MTD","%Gr L3M MTD","YTD Q4"]
    ws.write_row(2,1,headers,header)

    for r, (idx,row) in enumerate(df.iterrows()):
        ws.write(r+3,0,idx,cell)
        for c,val in enumerate(row):
            try:
                ws.write(r+3,c+1,float(val.strip('%'))/100, percent_fmt)
            except:
                ws.write(r+3,c+1,val,cell)

    ws.set_column(0,0,25)
    ws.set_column(1,1,12)
    ws.set_column(2,7,12)

    wb.close()
    output.seek(0)
    return output

st.download_button(
    "⬇️ Download Excel",
    to_excel(df_display, cut_off_str),
    "Quartile_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
