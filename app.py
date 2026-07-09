import streamlit as st
import pandas as pd
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy
from collections import defaultdict
from io import BytesIO

st.title("Crystal Report")

sponsor_file = st.file_uploader("Upload Data Sponsor (.xlsx)", type=["xlsx"])
batch_file   = st.file_uploader("Upload Batch Daily Report (.xls)", type=["xls"])


# ==============================
# LOAD SPONSOR LOOKUP
# ==============================
def load_sponsor_lookup(file):
    xls = pd.ExcelFile(file)
    df  = None
    for sheet in xls.sheet_names:
        tmp = pd.read_excel(xls, sheet_name=sheet)
        tmp.columns = tmp.columns.astype(str).str.strip()
        cols_up = [c.upper() for c in tmp.columns]
        if "ID_PARTNER" in cols_up and "TYPE" in cols_up and "RM" in cols_up:
            df = tmp
            break
    if df is None:
        st.error("Kolom ID_Partner / Type / RM tidak ditemukan di Data Sponsor.")
        st.stop()
    df.columns = [c.strip().upper() for c in df.columns]
    df = df[["ID_PARTNER", "TYPE", "RM"]].copy()
    def cid(x):
        if pd.isna(x): return None
        try:   return str(int(float(x))).strip()
        except: return str(x).strip()
    df["ID_PARTNER"] = df["ID_PARTNER"].apply(cid)
    df["TYPE"] = df["TYPE"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    df["RM"]   = df["RM"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    df = df.sort_values(["TYPE", "RM"], ascending=False)
    df = df.drop_duplicates(subset="ID_PARTNER", keep="first")
    return df.set_index("ID_PARTNER")[["TYPE", "RM"]].to_dict(orient="index")


def clean_pid(x):
    if x is None or x == "": return None
    try:   return str(int(float(x))).strip()
    except: return str(x).strip()


# ==============================
# PROCESS
# ==============================
def process(file_bytes, lookup):

    wb_in = xlrd.open_workbook(file_contents=file_bytes, formatting_info=True)
    ws_in = wb_in.sheet_by_index(0)
    nrows = ws_in.nrows

    # ------------------------------------------------------------------
    # Pass 1: Kumpulkan desig → {Mass/Middle/Major: total_amount} per batch
    # Col (0-indexed) di data rows:
    #   col0 = PartnerID, col4 = amount, col5 = designation
    # ------------------------------------------------------------------
    batch_desig_list = []
    cur_desig        = None
    in_batch         = False

    for r in range(nrows):
        row = ws_in.row_values(r)

        if str(row[2]).strip() == 'PartnerID':          # header batch
            if cur_desig is not None:
                batch_desig_list.append(cur_desig)
            cur_desig = defaultdict(lambda: defaultdict(float))
            in_batch  = True
            continue

        if str(row[0]).strip() == 'Total Batch Amount':
            in_batch = False
            continue

        if not in_batch:
            continue

        pid = clean_pid(row[0])
        if pid is None:
            in_batch = False
            continue

        amount      = row[4] if isinstance(row[4], (int, float)) else 0
        designation = str(row[5]).strip() if row[5] != '' else ''

        if pid in lookup and designation:
            dt = lookup[pid]['TYPE']
            if dt in ('Mass', 'Middle', 'Major'):
                cur_desig[designation][dt] += amount

    if cur_desig is not None:
        batch_desig_list.append(cur_desig)

    # ------------------------------------------------------------------
    # Salin workbook PERSIS (semua formatting, font, ukuran, dll terjaga)
    # ------------------------------------------------------------------
    wb_out = xl_copy(wb_in)
    ws_out = wb_out.get_sheet(0)

    # Style untuk nominal Rp di kolom Mass/Middle/Major (summary)
    rp_style = xlwt.easyxf(
        num_format_str=r'[$Rp-409]#,##0.00_);\([$Rp-409]#,##0.00\)'
    )

    # ------------------------------------------------------------------
    # Pass 2: Tulis DonorType, RM, Mass, Middle, Major ke sel yang tepat
    #
    # Struktur header batch (col, 0-indexed):
    #   col2="PartnerID"  col3="Partner Name"  col4="ChequeDate"
    #   col5="BankName"   col6="Payment Amount" col7="Designation"
    #
    # Di data rows, col6 & col7 KOSONG → kita isi:
    #   col6 → Donor Type (ganti header "Payment Amount" → "Donor Type")
    #   col7 → RM         (ganti header "Designation"    → "RM")
    #
    # Struktur summary:
    #   col0="Acount Fund"  col1="Designation"  col2="Description"  col3="Amount"
    #   Tambah: col4="Mass"  col5="Middle"  col6="Major"
    # ------------------------------------------------------------------
    matched       = 0
    not_found     = []
    in_batch      = False
    in_summary    = False
    batch_idx     = -1
    cur_desig_ref = None

    for r in range(nrows):
        row      = ws_in.row_values(r)
        col0_str = str(row[0]).strip()
        col2_str = str(row[2]).strip()

        # ---- Header batch ----
        if col2_str == 'PartnerID':
            in_batch      = True
            in_summary    = False
            batch_idx    += 1
            cur_desig_ref = batch_desig_list[batch_idx] \
                            if batch_idx < len(batch_desig_list) else {}
            # Rename header col6 → "Donor Type", col7 → "RM"
            ws_out.write(r, 6, 'Donor Type')
            ws_out.write(r, 7, 'RM')
            continue

        # ---- Akhir data batch ----
        if col0_str == 'Total Batch Amount':
            in_batch = False
            continue

        # ---- Masuk section For Finance ----
        if col0_str == 'For Finance':
            in_summary = True
            in_batch   = False
            continue

        # ---- Data rows batch: isi col6=DonorType, col7=RM ----
        if in_batch:
            pid = clean_pid(row[0])
            if pid is None:
                in_batch = False
                continue
            donor_type = ''
            rm         = ''
            if pid in lookup:
                donor_type = lookup[pid]['TYPE']
                rm         = lookup[pid]['RM']
                matched   += 1
            else:
                not_found.append(pid)
            ws_out.write(r, 6, donor_type)
            ws_out.write(r, 7, rm)
            continue

        # ---- Summary header: tambah Mass / Middle / Major ----
        if in_summary and col0_str == 'Acount Fund':
            ws_out.write(r, 4, 'Mass')
            ws_out.write(r, 5, 'Middle')
            ws_out.write(r, 6, 'Major')
            continue

        # ---- Summary data rows ----
        if in_summary:
            non_empty = [v for v in row if v not in ('', None)]
            if not non_empty:
                in_summary = False
                continue
            # Baris subtotal: hanya satu angka di col0
            if len(non_empty) == 1 and isinstance(row[0], (int, float)):
                in_summary = False
                continue

            designation = str(row[1]).strip() if row[1] not in ('', None) else ''
            if designation and cur_desig_ref:
                ta     = cur_desig_ref.get(designation, {})
                mass   = ta.get('Mass',   0)
                middle = ta.get('Middle', 0)
                major  = ta.get('Major',  0)

                ws_out.write(r, 4, mass   if mass   > 0 else '-',
                             rp_style if mass   > 0 else xlwt.Style.default_style)
                ws_out.write(r, 5, middle if middle > 0 else '-',
                             rp_style if middle > 0 else xlwt.Style.default_style)
                ws_out.write(r, 6, major  if major  > 0 else '-',
                             rp_style if major  > 0 else xlwt.Style.default_style)

    output = BytesIO()
    wb_out.save(output)
    return output.getvalue(), matched, list(set(not_found))


# ==============================
# MAIN
# ==============================
if sponsor_file and batch_file:
    lookup     = load_sponsor_lookup(sponsor_file)
    file_bytes = batch_file.read()

    result_bytes, matched, not_found = process(file_bytes, lookup)

    col1, col2 = st.columns(2)
    col1.metric("Matched (terisi)", matched)
    col2.metric("Tidak Ditemukan",  len(not_found))

    if not_found:
        with st.expander("PartnerID tidak ditemukan di Data Sponsor"):
            st.write(sorted(not_found))

    st.success("Selesai!")

    st.download_button(
        "Download Hasil",
        result_bytes,
        "Crystal_Report_FILLED.xls",
        mime="application/vnd.ms-excel"
    )

else:
    st.info("Silakan upload kedua file untuk memulai.")

exit code 0
Done
