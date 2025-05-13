import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.title("Perbandingan Transaksi Harian: Invoice vs Rekening Koran (Unit E-Channel > 100jt)")

# Upload dua file Excel
file_rekening = st.file_uploader("Upload Data 1 - Rekening Koran (Excel)", type=["xls", "xlsx"])
file_invoice = st.file_uploader("Upload Data 2 - Invoice (Excel)", type=["xls", "xlsx"])

if file_rekening and file_invoice:
    # Load data
    rekening_df = pd.read_excel(file_rekening)
    invoice_df = pd.read_excel(file_invoice)

    # Format dan parsing tanggal
    rekening_df['Post Date'] = pd.to_datetime(rekening_df['Post Date'], dayfirst=True, errors='coerce').dt.date
    invoice_df['TANGGAL INVOICE'] = pd.to_datetime(invoice_df['TANGGAL INVOICE'], errors='coerce')

    # Bersihkan nilai Amount dan Harga
    rekening_df['Amount'] = rekening_df['Amount'].astype(str).str.replace('[^0-9]', '', regex=True).astype(float)
    invoice_df['HARGA'] = pd.to_numeric(invoice_df['HARGA'], errors='coerce')

    # Filter hanya unit E-CHANNEL dan amount > 100jt dari rekening koran
    echannel_df = rekening_df[(rekening_df['Branch'].str.contains("E-CHANNEL", case=False, na=False)) & (rekening_df['Amount'] > 100_000_000)].copy()

    # Ekstrak tanggal dari kolom Description berdasarkan pola TRX TGL
    def extract_trx_tgl(text):
        match = re.search(r'TRX TGL ([0-9]{1,2} [A-Z]{3}(?:-[0-9]{1,2} [A-Z]{3})? [0-9]{4})', str(text))
        return match.group(1) if match else None

    echannel_df['Tanggal'] = echannel_df['Description'].apply(extract_trx_tgl)

    # Fungsi bantu konversi teks TGL ke list tanggal
    def trx_tgl_to_date_range(tgl_str):
        try:
            parts = tgl_str.split('-')
            if len(parts) == 2:
                start = pd.to_datetime(parts[0] + tgl_str[-5:], format='%d %b %Y')
                end = pd.to_datetime(parts[1], format='%d %b %Y')
                return pd.date_range(start, end)
            else:
                return [pd.to_datetime(tgl_str, format='%d %b %Y')]
        except:
            return []

    # Hitung total invoice sesuai tanggal di kolom Tanggal
    def sum_invoice_for_trx(tgl_str):
        total = 0
        for tgl in trx_tgl_to_date_range(tgl_str):
            total += invoice_df[invoice_df['TANGGAL INVOICE'].dt.date == tgl.date()]['HARGA'].sum()
        return total

    echannel_df['invoice'] = echannel_df['Tanggal'].apply(sum_invoice_for_trx)

    # Hitung selisih
    echannel_df['selisih'] = echannel_df['invoice'] - echannel_df['Amount']

    # Buat output 10 kolom sesuai spesifikasi
    final_df = echannel_df[[
        'Tanggal',
        'Post Date',
        'Branch',
        'Journal No.',
        'Description',
        'Amount',
        'invoice',
        'selisih',
        'Db/Cr',
        'Balance'
    ]]

    # Tampilkan preview
    st.subheader("Preview Output (Unit E-Channel > 100jt)")
    st.dataframe(final_df)

    # Export to Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Compare')
        writer.save()
    st.download_button(
        label="📥 Download Hasil dalam Excel",
        data=output.getvalue(),
        file_name="hasil_compare_unit_echannel_100jt.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
