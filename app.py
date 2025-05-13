
import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta

st.title("Compare Rekening Koran vs Invoice - Output 10 Kolom Final")

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

def extract_trx_range(desc):
    if pd.isnull(desc):
        return None
    match1 = re.search(r'TRX TGL ([0-9]{2} [A-Z]{3})(?:-([0-9]{2} [A-Z]{3}))? ([0-9]{4})', desc)
    if match1:
        start = f"{match1.group(1)} {match1.group(3)}"
        end = f"{match1.group(2)} {match1.group(3)}" if match1.group(2) else start
        return f"{start} - {end}" if start != end else start
    match2 = re.search(r'TRX TGL ([0-9]{2})(?:-([0-9]{2}))? ([A-Z]{3}) ([0-9]{4})', desc)
    if match2:
        hari1 = match2.group(1)
        hari2 = match2.group(2) or match2.group(1)
        bulan = match2.group(3)
        tahun = match2.group(4)
        start = f"{hari1} {bulan} {tahun}"
        end = f"{hari2} {bulan} {tahun}"
        return f"{start} - {end}" if start != end else start
    return None

def fix_invoice_summing(df1, df2):
    df2["Tanggal"] = df2["TANGGAL INVOICE"].dt.strftime("%d %b %Y").str.upper().str.strip()
    invoice_by_date = df2.groupby("Tanggal")["HARGA"].sum().to_dict()

    def sum_invoice(trx_range):
        if pd.isnull(trx_range):
            return 0
        if "-" not in trx_range:
            trx_range = translate_bulan(trx_range.strip())
            trx_date = safe_strptime(trx_range)
            if trx_date:
                key = trx_date.strftime("%d %b %Y").upper()
                return invoice_by_date.get(key, 0)
            return 0
        start_str, end_str = trx_range.split("-")
        start_date = safe_strptime(translate_bulan(start_str.strip()))
        end_date = safe_strptime(translate_bulan(end_str.strip()))
        if not start_date or not end_date:
            return 0
        date_list = [(start_date + timedelta(days=i)).strftime("%d %b %Y").upper()
                     for i in range((end_date - start_date).days + 1)]
        return sum(invoice_by_date.get(d, 0) for d in date_list)

    df1["Invoice"] = df1["Tanggal"].apply(sum_invoice)
    df1["Selisih"] = df1["Amount"] - df1["Invoice"]
    return df1

if file1 and file2:
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    df1["Post Date"] = pd.to_datetime(df1["Post Date"], dayfirst=True, errors='coerce')
    df1 = df1.dropna(subset=["Post Date", "Amount"])
    df1 = df1[(df1["Branch"].str.contains("UNIT E-CHANNEL", na=False)) & (df1["Amount"] > 100_000_000)].copy()
    df1["Tanggal"] = df1["Description"].apply(extract_trx_range)
    df1 = df1.dropna(subset=["Tanggal"])

    df2["TANGGAL INVOICE"] = pd.to_datetime(df2["TANGGAL INVOICE"], errors='coerce')
    df2 = df2.dropna(subset=["TANGGAL INVOICE", "HARGA"])

    df1 = fix_invoice_summing(df1, df2)

    df_final = df1[["Tanggal", "Post Date", "Branch", "Journal No.", "Description",
                    "Amount", "Invoice", "Selisih", "Db/Cr", "Balance"]]

    st.header("Output 10 Kolom Final")
    st.dataframe(df_final.fillna(""))

    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Compare Hasil')
    output.seek(0)

    st.download_button(
        label="Download Output Excel",
        data=output,
        file_name="hasil_compare_fixed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Silakan upload kedua file untuk melanjutkan.")
