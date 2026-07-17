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


def is_summary_data_row(row):
    """Cek apakah baris ini adalah data row di summary (bukan header, bukan subtotal, bukan kosong)."""
    non_empty = [v for v in row if v not in ('', None)]
    if not non_empty:
        return False, 'empty'
    if len(non_empty) == 1 and isinstance(row[0], (int, float)):
        return False, 'subtotal'
    if str(row[0]).strip() == 'Acount Fund':
        return False, 'header'
    return True, 'data'


# ==============================
# PROCESS
# ==============================
def process(file_bytes, lookup):

    wb_in = xlrd.open_workbook(file_contents=file_bytes, formatting_info=True)
    ws_in = wb_in.sheet_by_index(0)
    nrows = ws_in.nrows

    # ------------------------------------------------------------------
    # Pass 1: Per batch, kumpulkan:
    #   - desig_amounts : desig → {Mass, Middle, Major, Unknown: float}
    #   - desig_count   : desig → berapa kali muncul di summary-nya
    #
    # Col (0-indexed) di batch data rows:
    #   col0=PartnerID  col4=amount  col5=designation
    # Col (0-indexed) di summary data rows:
    #   col0=acct_fund  col1=designation  col2=description  col3=amount
    # ------------------------------------------------------------------
    all_desig_amounts = []   # list of dict per batch
    all_desig_counts  = []   # list of dict per batch

    cur_amounts = None
    cur_counts  = None
    in_batch    = False
    in_summary  = False

    for r in range(nrows):
        row      = ws_in.row_values(r)
        col0_str = str(row[0]).strip()
        col2_str = str(row[2]).strip()

        # ---- Mulai batch baru ----
        if col2_str == 'PartnerID':
            # Simpan batch sebelumnya jika ada
            if cur_amounts is not None:
                all_desig_amounts.append(cur_amounts)
                all_desig_counts.append(cur_counts)
            cur_amounts = defaultdict(lambda: defaultdict(float))
            cur_counts  = defaultdict(int)
            in_batch    = True
            in_summary  = False
            continue

        if col0_str == 'Total Batch Amount':
            in_batch = False
            continue

        if col0_str == 'For Finance':
            in_summary = True
            in_batch   = False
            continue

        # ---- Batch data row: hitung desig → type → amount ----
        if in_batch:
            pid = clean_pid(row[0])
            if pid is None:
                in_batch = False
                continue
            amount      = row[4] if isinstance(row[4], (int, float)) else 0
            designation = str(row[5]).strip() if row[5] != '' else ''
            if designation:
                if pid in lookup:
                    dt = lookup[pid]['TYPE']
                    if dt in ('Mass', 'Middle', 'Major'):
                        cur_amounts[designation][dt] += amount
                    else:
                        # Donor type kosong / bukan Mass-Middle-Major → Unknown
                        cur_amounts[designation]['Unknown'] += amount
                else:
                    # PartnerID tidak ada di Data Sponsor → Unknown
                    cur_amounts[designation]['Unknown'] += amount
            continue

        # ---- Summary rows: hitung berapa kali tiap designation muncul ----
        if in_summary:
            is_data, kind = is_summary_data_row(row)
            if kind == 'empty':
                in_summary = False
                continue
            if kind == 'subtotal':
                in_summary = False
                continue
            if kind == 'header':
                continue
            # Data row
            designation = str(row[1]).strip() if row[1] not in ('', None) else ''
            if designation:
                cur_counts[designation] += 1

    # Simpan batch terakhir
    if cur_amounts is not None:
        all_desig_amounts.append(cur_amounts)
        all_desig_counts.append(cur_counts)

    # ------------------------------------------------------------------
    # Salin workbook PERSIS (semua formatting, font, ukuran, dll terjaga)
    # ------------------------------------------------------------------
    wb_out = xl_copy(wb_in)
    ws_out = wb_out.get_sheet(0)

    rp_style = xlwt.easyxf(
        num_format_str=r'[$Rp-409]#,##0.00_);\([$Rp-409]#,##0.00\)'
    )
    def write_amount(r, c, val):
        if isinstance(val, (int, float)) and val > 0:
            ws_out.write(r, c, val, rp_style)
        else:
            ws_out.write(r, c, '-')

    sum_mass = 0
    sum_middle = 0
    sum_major = 0
    sum_unknown = 0
    
    # ------------------------------------------------------------------
    # Pass 2: Tulis ke file
    # ------------------------------------------------------------------
    matched       = 0
    not_found     = []
    in_batch      = False
    in_summary    = False
    batch_idx     = -1
    cur_amounts_r = None   # desig_amounts untuk batch ini
    cur_counts_r  = None   # desig_counts  untuk batch ini

    for r in range(nrows):
        row      = ws_in.row_values(r)
        col0_str = str(row[0]).strip()
        col2_str = str(row[2]).strip()

        # ---- Header batch ----
        if col2_str == 'PartnerID':
            in_batch      = True
            in_summary    = False
            batch_idx    += 1
            cur_amounts_r = all_desig_amounts[batch_idx] \
                            if batch_idx < len(all_desig_amounts) else {}
            cur_counts_r  = all_desig_counts[batch_idx]  \
                            if batch_idx < len(all_desig_counts)  else {}
            # Rename header: col6 → "Donor Type", col7 → "RM"
            ws_out.write(r, 6, 'Donor Type')
            ws_out.write(r, 7, 'RM')
            continue

        if col0_str == 'Total Batch Amount':
            in_batch = False
            continue

        if col0_str == 'For Finance':
            in_summary = True
            in_batch = False
        
            sum_mass = 0
            sum_middle = 0
            sum_major = 0
            sum_unknown = 0
            continue

        # ---- Batch data row: isi Donor Type & RM ----
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

        # ---- Summary header: tambah Mass / Middle / Major / Unknown ----
        if in_summary and col0_str == 'Acount Fund':
            ws_out.write(r, 4, 'Mass')
            ws_out.write(r, 5, 'Middle')
            ws_out.write(r, 6, 'Major')
            ws_out.write(r, 7, 'Unknown')
            continue

        # ---- Summary data rows ----
        if in_summary:
            is_data, kind = is_summary_data_row(row)
            if kind == 'empty':
                total_row = r + 1
                
                ws_out.write(r, 2, 'Total :')
                write_amount(r, 3, sum_mass + sum_middle + sum_major + sum_unknown)
                write_amount(r, 4, sum_mass)
                write_amount(r, 5, sum_middle)
                write_amount(r, 6, sum_major)
                write_amount(r, 7, sum_unknown)
            
                in_summary = False
                continue
            if kind == 'subtotal':
                ws_out.write(r, 2, 'Total :')
                write_amount(r, 3, sum_mass + sum_middle + sum_major + sum_unknown)
                write_amount(r, 4, sum_mass)
                write_amount(r, 5, sum_middle)
                write_amount(r, 6, sum_major)
                write_amount(r, 7, sum_unknown)
            
                in_summary = False
                continue
            if kind == 'header':
                continue

            designation = str(row[1]).strip() if row[1] not in ('', None) else ''
            if not designation:
                continue

            # Designation muncul lebih dari 1x di summary ini → tulis "?"
            count = cur_counts_r.get(designation, 1)
            if count > 1:
                ws_out.write(r, 4, '?')
                ws_out.write(r, 5, '?')
                ws_out.write(r, 6, '?')
                ws_out.write(r, 7, '?')
            else:
                ta      = cur_amounts_r.get(designation, {})
                mass    = ta.get('Mass',    0)
                middle  = ta.get('Middle',  0)
                major   = ta.get('Major',   0)
                unknown = ta.get('Unknown', 0)
                write_amount(r, 4, mass)
                write_amount(r, 5, middle)
                write_amount(r, 6, major)
                write_amount(r, 7, unknown)
                sum_mass += mass
                sum_middle += middle
                sum_major += major
                sum_unknown += unknown

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
