
import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta

st.title("Compare Rekening Koran vs Invoice - Output 10 Kolom Final")

st.header("Periode Rekonsiliasi")
start_date = st.date_input("Tanggal Mulai")
end_date = st.date_input("Tanggal Selesai")

st.header("Upload Rekening Koran (Data 1)")
file1 = st.file_uploader("Upload Excel Rekening Koran", type=["xls", "xlsx"], key="file1")

st.header("Upload Invoice (Data 2)")
file2 = st.file_uploader("Upload Excel Invoice", type=["xls", "xlsx"], key="file2")

def translate_bulan(text):
    bulan_map = {
        "JAN": "JAN", "FEB": "FEB", "MAR": "MAR", "APR": "APR", "MEI": "MAY",
        "JUN": "JUN", "JUL": "JUL", "AGU": "AUG", "SEP": "SEP", "OKT": "OCT",
        "NOV": "NOV", "DES": "DEC"
    }
    for indo, eng in bulan_map.items():
        text = text.replace(indo, eng)
    return text

def safe_strptime(s):
    try:
        return datetime.strptime(s, "%d %b %Y")
    except ValueError:
        return None

def expand_trx_dates(desc):
    if pd.isnull(desc):
        return []

    match1 = re.search(r'TRX TGL ([0-9]{2} [A-Z]{3})(?:-([0-9]{2} [A-Z]{3}))? ([0-9]{4})', desc)
    if match1:
        start_str = translate_bulan(f"{match1.group(1)} {match1.group(3)}")
        end_str = start_str
        if match1.group(2):
            end_str = translate_bulan(f"{match1.group(2)} {match1.group(3)}")
        start_date = safe_strptime(start_str)
        end_date = safe_strptime(end_str)
        if start_date and end_date:
            return [(start_date + timedelta(days=i)).strftime("%d %b %Y")
                    for i in range((end_date - start_date).days + 1)]

    match2 = re.search(r'TRX TGL ([0-9]{2})(?:-([0-9]{2}))? ([A-Z]{3}) ([0-9]{4})', desc)
    if match2:
        hari1 = match2.group(1)
        hari2 = match2.group(2) or match2.group(1)
        bulan = match2.group(3)
        tahun = match2.group(4)
        start_str = translate_bulan(f"{hari1} {bulan} {tahun}")
        end_str = translate_bulan(f"{hari2} {bulan} {tahun}")
        start_date = safe_strptime(start_str)
        end_date = safe_strptime(end_str)
        if start_date and end_date:
            return [(start_date + timedelta(days=i)).strftime("%d %b %Y")
                    for i in range((end_date - start_date).days + 1)]

    return []

if file1 and file2 and start_date and end_date:
    tanggal_list = pd.date_range(start=start_date, end=end_date)
    df_output = pd.DataFrame({"Tanggal": tanggal_list.strftime("%d %b %Y").str.upper().str.strip()})

    df1 = pd.read_excel(file1)
    df1["Post Date"] = pd.to_datetime(df1["Post Date"], dayfirst=True, errors='coerce')
    df1 = df1.dropna(subset=["Post Date", "Amount"])
    df1 = df1[(df1["Branch"].str.contains("UNIT E-CHANNEL", na=False)) & (df1["Amount"] > 100_000_000)].copy()
    df1["Tanggal TRX List"] = df1["Description"].apply(expand_trx_dates)

    df2 = pd.read_excel(file2)
    df2["TANGGAL INVOICE"] = pd.to_datetime(df2["TANGGAL INVOICE"], errors='coerce')
    df2 = df2.dropna(subset=["TANGGAL INVOICE", "HARGA"])
    df2["Tanggal"] = df2["TANGGAL INVOICE"].dt.strftime("%d %b %Y").str.upper().str.strip()
    df2_grouped = df2.groupby("Tanggal")["HARGA"].sum().reset_index()

    df_output["Post Date"] = None
    df_output["Branch"] = None
    df_output["Journal No."] = None
    df_output["Description"] = None
    df_output["Amount"] = 0
    df_output["Invoice"] = 0
    df_output["Selisih"] = 0
    df_output["Db/Cr"] = None
    df_output["Balance"] = None

    for idx, row in df_output.iterrows():
        tanggal_row = row["Tanggal"].strip().upper()
        for _, d1_row in df1.iterrows():
            if tanggal_row in [t.strip().upper() for t in d1_row["Tanggal TRX List"]]:
                df_output.at[idx, "Post Date"] = d1_row["Post Date"]
                df_output.at[idx, "Branch"] = d1_row["Branch"]
                df_output.at[idx, "Journal No."] = d1_row["Journal No."]
                df_output.at[idx, "Description"] = d1_row["Description"]
                df_output.at[idx, "Amount"] = d1_row["Amount"]
                df_output.at[idx, "Db/Cr"] = d1_row["Db/Cr"]
                df_output.at[idx, "Balance"] = d1_row["Balance"]
                break

    df_output = pd.merge(df_output, df2_grouped, on="Tanggal", how="left")
    df_output["Invoice"] = df_output["HARGA"].fillna(0)
    df_output.drop(columns=["HARGA"], inplace=True)
    df_output["Selisih"] = df_output["Amount"] - df_output["Invoice"]

    st.write("Jumlah baris final:", len(df_output))
    st.header("Hasil Compare Detail (10 Kolom)")
    st.dataframe(df_output.fillna(""))

    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_output.to_excel(writer, index=False, sheet_name='Compare Hasil')
    output.seek(0)

    st.download_button(
        label="Download Hasil Compare (Excel)",
        data=output,
        file_name="hasil_compare.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Silakan upload kedua file dan pilih periode untuk melanjutkan.")
