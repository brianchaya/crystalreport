import streamlit as st
import pandas as pd
import xlrd
import openpyxl
from collections import defaultdict
from io import BytesIO

st.title("Crystal Report")

sponsor_file = st.file_uploader("Upload Data Sponsor (.xlsx)", type=["xlsx"])
batch_file   = st.file_uploader("Upload Batch Daily Report (.xls / .xlsx)", type=["xls", "xlsx"])


# ==============================
# LOAD SPONSOR LOOKUP
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
        st.error("Sheet dengan kolom ID_Partner, Type, RM tidak ditemukan di Data Sponsor.")
        st.stop()

    df.columns = [c.strip().upper() for c in df.columns]
    df = df[["ID_PARTNER", "TYPE", "RM"]].copy()

    def clean_id(x):
        if pd.isna(x): return None
        try: return str(int(float(x))).strip()
        except: return str(x).strip()

    df["ID_PARTNER"] = df["ID_PARTNER"].apply(clean_id)
    df["TYPE"] = df["TYPE"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    df["RM"]   = df["RM"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    df = df.sort_values(by=["TYPE", "RM"], ascending=False)
    df = df.drop_duplicates(subset="ID_PARTNER", keep="first")
    return df.set_index("ID_PARTNER")[["TYPE", "RM"]].to_dict(orient="index")


def clean_pid(x):
    if x is None or x == "": return None
    try: return str(int(float(x))).strip()
    except: return str(x).strip()


# ==============================
# PROCESS XLS → XLSX
# Aturan:
#   - Batch data row  : tulis DonorType ke col E (idx 5, 1-based), RM ke col F (idx 6)
#   - Summary header  : tambah label Mass/Middle/Major setelah Amount
#   - Summary data row: tambah nilai Mass/Middle/Major
#   - Semua baris lain: copy persis apa adanya
# ==============================
def process(file_bytes, lookup):
    wb_in  = xlrd.open_workbook(file_contents=file_bytes)
    ws_in  = wb_in.sheet_by_index(0)

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = wb_in.sheet_names()[0]

    nrows   = ws_in.nrows
    ncols   = ws_in.ncols
    matched = 0
    not_found = []

    # ---- Pass 1: build desig_type_amount per batch ----
    # Kolom data (0-indexed xlrd):
    #   col0 = PartnerID data, col4 = Amount data, col5 = Designation data
    # Setelah kita tulis ke col4 (1-based openpyxl = E) dengan DonorType,
    # data Amount di col4 akan tertimpa. Jadi kita baca dulu di pass 1.

    # Kumpulkan per-batch: desig → {Mass, Middle, Major}
    batch_desig = []       # list of dict per batch
    cur_desig   = None
    in_batch    = False

    for r in range(nrows):
        row = ws_in.row_values(r)
        # Deteksi header batch: col2 (0-idx) == 'PartnerID'
        if str(row[2]).strip() == 'PartnerID':
            if cur_desig is not None:
                batch_desig.append(cur_desig)
            cur_desig = defaultdict(lambda: defaultdict(float))
            in_batch  = True
            continue
        if not in_batch or cur_desig is None:
            continue
        if str(row[0]).strip() == 'Total Batch Amount':
            in_batch = False
            continue
        # Data row: col0 harus angka (PartnerID)
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
        batch_desig.append(cur_desig)

    # ---- Pass 2: copy semua baris, terapkan perubahan ----
    in_batch      = False
    in_summary    = False
    batch_idx     = -1
    cur_desig_ref = None

    for r in range(nrows):
        row = ws_in.row_values(r)
        out_row_idx = r + 1   # openpyxl 1-based

        # Tulis semua kolom dulu (copy apa adanya)
        for c in range(ncols):
            val = row[c]
            # xlrd kadang kembalikan float untuk angka bulat, normalkan
            if isinstance(val, float) and val == int(val):
                val = int(val)
            ws_out.cell(row=out_row_idx, column=c + 1, value=val)

        # ---- Deteksi baris header batch ----
        if str(row[2]).strip() == 'PartnerID':
            in_batch   = True
            in_summary = False
            batch_idx += 1
            cur_desig_ref = batch_desig[batch_idx] if batch_idx < len(batch_desig) else {}
            continue

        # ---- Deteksi akhir batch data ----
        if str(row[0]).strip() == 'Total Batch Amount':
            in_batch = False
            continue

        # ---- Deteksi masuk summary ----
        if str(row[0]).strip() == 'For Finance':
            in_summary = True
            in_batch   = False
            continue

        # ---- Batch data row ----
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
            # Tulis ke col E (idx 5, 1-based) dan col F (idx 6, 1-based)
            ws_out.cell(row=out_row_idx, column=5, value=donor_type)
            ws_out.cell(row=out_row_idx, column=6, value=rm)
            continue

        # ---- Summary header: "Acount Fund" ----
        if in_summary and str(row[0]).strip() == 'Acount Fund':
            ws_out.cell(row=out_row_idx, column=5, value='Mass')
            ws_out.cell(row=out_row_idx, column=6, value='Middle')
            ws_out.cell(row=out_row_idx, column=7, value='Major')
            continue

        # ---- Summary data row ----
        if in_summary:
            # Keluar dari summary kalau baris kosong atau angka tunggal (subtotal)
            non_empty = [v for v in row if v != '' and v is not None]
            if not non_empty:
                in_summary = False
                continue
            # Subtotal row: hanya 1 angka di col0
            if len(non_empty) == 1 and isinstance(row[0], (int, float)):
                in_summary = False
                continue
            # Baris data summary: col1 = designation
            designation = str(row[1]).strip() if row[1] != '' else ''
            if designation and cur_desig_ref:
                type_amt = cur_desig_ref.get(designation, {})
                mass   = type_amt.get('Mass',   0)
                middle = type_amt.get('Middle', 0)
                major  = type_amt.get('Major',  0)
                ws_out.cell(row=out_row_idx, column=5, value=mass   if mass   > 0 else '-')
                ws_out.cell(row=out_row_idx, column=6, value=middle if middle > 0 else '-')
                ws_out.cell(row=out_row_idx, column=7, value=major  if major  > 0 else '-')

    output = BytesIO()
    wb_out.save(output)
    return output.getvalue(), matched, list(set(not_found))


# ==============================
# MAIN
# ==============================
if sponsor_file and batch_file:
    lookup     = load_sponsor_lookup(sponsor_file)
    file_bytes = batch_file.read()

    if not batch_file.name.lower().endswith('.xls'):
        st.error("Upload file .xls (bukan .xlsx).")
        st.stop()

    result_bytes, matched, not_found = process(file_bytes, lookup)

    col1, col2 = st.columns(2)
    col1.metric("Matched (terisi)", matched)
    col2.metric("Tidak Ditemukan", len(not_found))

    if not_found:
        with st.expander("PartnerID tidak ditemukan di Data Sponsor"):
            st.write(sorted(not_found))

    st.success("Selesai!")

    st.download_button(
        "Download Hasil",
        result_bytes,
        "Crystal_Report_FILLED.xlsx"
    )
else:
    st.info("Silakan upload kedua file untuk memulai.")
