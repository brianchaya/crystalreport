import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.title("Donor Type & RM Auto-Fill (Batch Report)")

st.markdown(
    """
Upload **Data Sponsor** (sumber Donor Type & RM) dan **Batch Daily Report** (file yang mau diisi).
App akan mencocokkan `PartnerID` di Batch Report dengan `ID_Partner` di Data Sponsor, lalu:
- Kolom **ChequeDate** akan diisi **Donor Type**
- Kolom **BankName** akan diisi **RM**
"""
)

sponsor_file = st.file_uploader("Upload Data Sponsor (.xlsx)", type=["xlsx"])
batch_file = st.file_uploader("Upload Batch Daily Report (.xlsx)", type=["xlsx"])


# ==============================
# LOAD SPONSOR LOOKUP (ID_Partner -> Type, RM)
# ==============================
def load_sponsor_lookup(file):
    xls = pd.ExcelFile(file)

    df = None
    for sheet in xls.sheet_names:
        tmp = pd.read_excel(xls, sheet_name=sheet)
        tmp.columns = tmp.columns.astype(str).str.strip()
        cols_upper = [c.upper() for c in tmp.columns]
        if "ID_PARTNER" in cols_upper and "TYPE" in cols_upper and "RM" in cols_upper:
            df = tmp
            break

    if df is None:
        st.error("Sheet dengan kolom ID_Partner, Type, dan RM tidak ditemukan di file Data Sponsor.")
        st.stop()

    df.columns = [c.strip().upper() for c in df.columns]
    df = df[["ID_PARTNER", "TYPE", "RM"]].copy()

    # Bersihkan ID_Partner jadi string angka rapi (tanpa .0)
    def clean_id(x):
        if pd.isna(x):
            return None
        try:
            return str(int(float(x))).strip()
        except (ValueError, TypeError):
            return str(x).strip()

    df["ID_PARTNER"] = df["ID_PARTNER"].apply(clean_id)
    df["TYPE"] = df["TYPE"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    df["RM"] = df["RM"].apply(lambda x: "" if pd.isna(x) else str(x).strip())

    # Kalau ada ID_Partner duplikat, ambil baris pertama yang punya Type/RM terisi
    df = df.sort_values(by=["TYPE", "RM"], ascending=False)
    df = df.drop_duplicates(subset="ID_PARTNER", keep="first")

    lookup = df.set_index("ID_PARTNER")[["TYPE", "RM"]].to_dict(orient="index")
    return lookup


def clean_partner_id(x):
    if x is None:
        return None
    try:
        return str(int(float(x))).strip()
    except (ValueError, TypeError):
        return str(x).strip()


# ==============================
# PROCESS BATCH REPORT (preserve original layout)
# ==============================
def process_batch_report(file, lookup):
    wb = openpyxl.load_workbook(file)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    matched = 0
    not_found = 0
    not_found_ids = []

    # Catatan penting: di file batch report ini, label header "PartnerID",
    # "ChequeDate", "BankName" tertulis di kolom C/F/G, TAPI data riil di
    # bawahnya selalu mulai dari kolom A dengan urutan tetap:
    # A=PartnerID, B=Partner Name, C=ChequeDate, D=BankName, E=Payment Amount.
    # Jadi kita pakai posisi data riil (bukan posisi label header).
    COL_PARTNERID = 1
    COL_CHEQUEDATE = 3
    COL_BANKNAME = 4

    in_batch_block = False

    for row in range(1, ws.max_row + 1):
        cell_c = ws.cell(row, 3).value  # cek label header "PartnerID" di kolom C

        if cell_c is not None and str(cell_c).strip() == "PartnerID":
            in_batch_block = True
            continue

        if not in_batch_block:
            continue

        partner_id_raw = ws.cell(row, COL_PARTNERID).value

        # Berhenti / skip baris non-data (Total Batch Amount, baris kosong, dll)
        if partner_id_raw is None:
            in_batch_block = False
            continue

        try:
            float(partner_id_raw)
        except (ValueError, TypeError):
            in_batch_block = False
            continue

        partner_id = clean_partner_id(partner_id_raw)

        if partner_id in lookup:
            donor_type = lookup[partner_id]["TYPE"]
            rm = lookup[partner_id]["RM"]

            ws.cell(row, COL_CHEQUEDATE).value = donor_type
            ws.cell(row, COL_BANKNAME).value = rm

            matched += 1
        else:
            not_found += 1
            not_found_ids.append(partner_id)

    output = BytesIO()
    wb.save(output)

    return output.getvalue(), matched, not_found, not_found_ids


# ==============================
# MAIN
# ==============================
if sponsor_file and batch_file:

    lookup = load_sponsor_lookup(sponsor_file)
    result_bytes, matched, not_found, not_found_ids = process_batch_report(batch_file, lookup)

    col1, col2 = st.columns(2)
    col1.metric("Matched (terisi)", matched)
    col2.metric("Tidak Ditemukan", not_found)

    if not_found_ids:
        with st.expander("Lihat PartnerID yang tidak ditemukan di Data Sponsor"):
            st.write(sorted(set(not_found_ids)))

    st.success("Selesai! Kolom ChequeDate sudah berisi Donor Type, kolom BankName sudah berisi RM.")

    st.download_button(
        "Download Hasil (Batch Report Updated)",
        result_bytes,
        "Batch_Report_FILLED.xlsx"
    )
else:
    st.info("Silakan upload kedua file untuk memulai.")
