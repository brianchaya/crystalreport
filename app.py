import streamlit as st
import pandas as pd
import xlrd
import openpyxl
from openpyxl.styles import Alignment
from collections import defaultdict, OrderedDict
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
        if pd.isna(x):
            return None
        try:
            return str(int(float(x))).strip()
        except (ValueError, TypeError):
            return str(x).strip()

    df["ID_PARTNER"] = df["ID_PARTNER"].apply(clean_id)
    df["TYPE"] = df["TYPE"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    df["RM"]   = df["RM"].apply(lambda x: "" if pd.isna(x) else str(x).strip())

    df = df.sort_values(by=["TYPE", "RM"], ascending=False)
    df = df.drop_duplicates(subset="ID_PARTNER", keep="first")

    return df.set_index("ID_PARTNER")[["TYPE", "RM"]].to_dict(orient="index")


def clean_pid(x):
    if x is None or x == "":
        return None
    try:
        return str(int(float(x))).strip()
    except (ValueError, TypeError):
        return str(x).strip()


# ==============================
# PARSE XLS (row by row)
# ==============================
def parse_xls(file_bytes):
    """
    Parse the xls into a list of blocks. Each block is either:
      {'type': 'title', 'text': str}
      {'type': 'batch', 'header_text': str, 'data': [(pid, name, cheque_raw, bank_raw, amount, designation)], 'total': float}
      {'type': 'summary', 'header_batch': str, rows: [(acct_fund, designation, description, amount)], 'total': float}
      {'type': 'grand_total', 'amount': float}
      {'type': 'footer', 'text': str}
    """
    wb = xlrd.open_workbook(file_contents=file_bytes)
    ws = wb.sheet_by_index(0)

    blocks = []
    r = 0
    nrows = ws.nrows

    # --- Title row (row 0) ---
    if nrows > 0:
        blocks.append({'type': 'title', 'text': ws.row_values(0)[0]})
        r = 1

    while r < nrows:
        row = ws.row_values(r)

        # Skip fully empty
        if all(v == '' or v is None for v in row):
            r += 1
            continue

        col2 = str(row[2]).strip() if row[2] != '' else ''
        col0 = row[0]
        col0_str = str(col0).strip()

        # --- Batch header ---
        if col2 == 'PartnerID':
            header_text = str(row[0]).strip()
            batch_data = []
            r += 1
            while r < nrows:
                dr = ws.row_values(r)
                dc0 = dr[0]
                # End of batch data rows
                if str(dc0).strip() == 'Total Batch Amount':
                    total_batch = dr[1] if isinstance(dr[1], (int, float)) else 0
                    r += 1
                    break
                # Skip empty rows mid-batch
                if all(v == '' or v is None for v in dr):
                    r += 1
                    continue
                # Check if it's a numeric PartnerID row
                pid = None
                try:
                    pid = str(int(float(dc0))).strip()
                except (ValueError, TypeError):
                    pass
                if pid is None:
                    # Not a data row, stop
                    break
                name       = str(dr[1]).strip() if dr[1] != '' else ''
                cheque_raw = dr[2]   # date serial
                bank_raw   = str(dr[3]).strip() if dr[3] != '' else ''
                amount     = dr[4] if isinstance(dr[4], (int, float)) else 0
                designation = str(dr[5]).strip() if dr[5] != '' else ''
                batch_data.append((pid, name, cheque_raw, bank_raw, amount, designation))
                r += 1
            blocks.append({
                'type': 'batch',
                'header_text': header_text,
                'data': batch_data,
                'total': total_batch
            })
            continue

        # --- "For Finance" summary section ---
        if col0_str == 'For Finance':
            summary_rows = []
            r += 1
            while r < nrows:
                sr = ws.row_values(r)
                sc0 = sr[0]
                sc0_str = str(sc0).strip()
                # Skip empty
                if all(v == '' or v is None for v in sr):
                    r += 1
                    continue
                # Header row "Acount Fund, Designation..."
                if sc0_str == 'Acount Fund':
                    r += 1
                    continue
                # Total row: col0 is a number, rest empty
                if isinstance(sc0, (int, float)) and sc0 != 0 and all(v == '' or v is None for v in sr[1:]):
                    summary_total = sc0
                    r += 1
                    break
                # Check if next section starts (batch header or grand total)
                try:
                    float(sc0)
                    # could be summary total or grand total
                    summary_total = sc0
                    r += 1
                    break
                except (ValueError, TypeError):
                    pass
                if str(sr[2]).strip() == 'PartnerID' or sc0_str == 'For Finance':
                    break
                # Summary data row
                acct_fund   = sc0_str
                designation = str(sr[1]).strip() if sr[1] != '' else ''
                description = str(sr[2]).strip() if sr[2] != '' else ''
                amount      = sr[3] if isinstance(sr[3], (int, float)) else 0
                summary_rows.append((acct_fund, designation, description, amount))
                r += 1
            blocks.append({
                'type': 'summary',
                'rows': summary_rows,
                'total': summary_total
            })
            continue

        # --- Grand total (standalone big number at bottom) ---
        # Only if not already covered above
        try:
            val = float(col0)
            if all(v == '' or v is None for v in row[1:]):
                blocks.append({'type': 'grand_total', 'amount': val})
                r += 1
                continue
        except (ValueError, TypeError):
            pass

        # --- Footer / print date row ---
        if 'Print Date' in col0_str or 'D:\\' in col0_str:
            blocks.append({'type': 'footer', 'text': col0_str})
            r += 1
            continue

        r += 1

    return blocks


# ==============================
# PROCESS & WRITE OUTPUT XLSX
# ==============================
def process_and_write(blocks, lookup):
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Cristal report"

    matched = 0
    not_found_ids = []
    out_row = 1

    def write_row(values, bold=False, wrap=False):
        nonlocal out_row
        for col_idx, val in enumerate(values, start=1):
            cell = ws_out.cell(row=out_row, column=col_idx, value=val)
            if bold:
                cell.font = openpyxl.styles.Font(bold=True)
            if wrap:
                cell.alignment = Alignment(wrap_text=True)
        out_row += 1

    # ---- For each batch, build desig_type_amount GLOBALLY across ALL batches ----
    # (summary rows may summarize across batches per designation)
    # Actually each summary is per-batch, so build per-batch inside loop.

    batch_blocks = [b for b in blocks if b['type'] == 'batch']
    summary_blocks = [b for b in blocks if b['type'] == 'summary']

    # Build desig_type_amount per batch index
    # We assume summary[i] corresponds to batch[i]
    desig_type_amount_per_batch = []
    for batch in batch_blocks:
        desig_type = defaultdict(lambda: defaultdict(float))
        for pid, name, cheque_raw, bank_raw, amount, designation in batch['data']:
            if designation == '':
                continue
            donor_type = ''
            if pid in lookup:
                donor_type = lookup[pid]['TYPE']
            if donor_type in ('Mass', 'Middle', 'Major') and donor_type != '':
                desig_type[designation][donor_type] += amount
        desig_type_amount_per_batch.append(desig_type)

    # ---- Write output ----
    batch_idx = 0
    summary_idx = 0

    for block in blocks:
        btype = block['type']

        if btype == 'title':
            write_row([block['text']], bold=True, wrap=True)
            ws_out.row_dimensions[out_row - 1].height = 50
            out_row += 1  # blank after title

        elif btype == 'batch':
            desig_type = desig_type_amount_per_batch[batch_idx]
            batch_idx += 1

            # --- Header row (no Designation col) ---
            header_text = block['header_text']
            # Recalculate total items after merge
            raw_data = block['data']

            # Merge by PartnerID within this batch
            merged = OrderedDict()
            for pid, name, cheque_raw, bank_raw, amount, designation in raw_data:
                if pid not in merged:
                    merged[pid] = {
                        'name': name,
                        'cheque_raw': cheque_raw,
                        'bank_raw': bank_raw,
                        'amount': 0.0
                    }
                merged[pid]['amount'] += amount

            total_items = len(merged)
            # Rebuild header text with corrected item count
            import re
            new_header = re.sub(r'Total Batch Items\s*:\s*\d+',
                                f'Total Batch Items :{total_items}',
                                header_text)

            write_row(
                [new_header, f'Total Batch Items :{total_items}',
                 'PartnerID', 'Partner Name', 'ChequeDate', 'BankName', 'Payment Amount'],
                bold=True, wrap=True
            )
            ws_out.row_dimensions[out_row - 1].height = 60

            # --- Merged data rows ---
            for pid, info in merged.items():
                donor_type = ''
                rm = ''
                if pid in lookup:
                    donor_type = lookup[pid]['TYPE']
                    rm = lookup[pid]['RM']
                    matched += 1
                else:
                    not_found_ids.append(pid)

                write_row([
                    int(pid) if pid.isdigit() else pid,
                    info['name'],
                    donor_type,   # replaces ChequeDate
                    rm,            # replaces BankName
                    info['amount']
                ])

            # --- Total Batch Amount ---
            write_row(['Total Batch Amount', block['total']], bold=True)
            out_row += 1  # blank

        elif btype == 'summary':
            if summary_idx < len(desig_type_amount_per_batch):
                desig_type = desig_type_amount_per_batch[summary_idx]
            else:
                desig_type = defaultdict(lambda: defaultdict(float))
            summary_idx += 1

            write_row(['For Finance'], bold=True)
            write_row(
                ['Acount Fund', 'Designation', 'Description', 'Amount', 'Mass', 'Middle', 'Major'],
                bold=True
            )

            for acct_fund, designation, description, amount in block['rows']:
                type_amounts = desig_type.get(designation, {})
                mass   = type_amounts.get('Mass', 0)
                middle = type_amounts.get('Middle', 0)
                major  = type_amounts.get('Major', 0)
                write_row([
                    acct_fund,
                    designation,
                    description,
                    amount,
                    mass   if mass   > 0 else '-',
                    middle if middle > 0 else '-',
                    major  if major  > 0 else '-',
                ])

            # Summary total
            write_row([block['total']], bold=True)
            out_row += 1  # blank

        elif btype == 'grand_total':
            write_row([block['amount']], bold=True)
            out_row += 1

        elif btype == 'footer':
            write_row([block['text']])

    # Auto-width
    for col in ws_out.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    lines = str(cell.value).split('\n')
                    cell_max = max(len(l) for l in lines)
                    max_len = max(max_len, cell_max)
            except Exception:
                pass
        ws_out.column_dimensions[col_letter].width = min(max_len + 4, 50)

    output = BytesIO()
    wb_out.save(output)
    return output.getvalue(), matched, list(set(not_found_ids))


# ==============================
# MAIN
# ==============================
if sponsor_file and batch_file:
    lookup = load_sponsor_lookup(sponsor_file)

    file_bytes = batch_file.read()
    
    # Detect xls vs xlsx
    if batch_file.name.lower().endswith('.xls'):
        blocks = parse_xls(file_bytes)
    else:
        st.error("Format .xlsx belum didukung untuk file baru ini — gunakan file .xls.")
        st.stop()

    result_bytes, matched, not_found_ids = process_and_write(blocks, lookup)

    batch_count   = sum(1 for b in blocks if b['type'] == 'batch')
    summary_count = sum(1 for b in blocks if b['type'] == 'summary')

    col1, col2, col3 = st.columns(3)
    col1.metric("Batch diproses", batch_count)
    col2.metric("Matched (terisi)", matched)
    col3.metric("Tidak Ditemukan", len(not_found_ids))

    if not_found_ids:
        with st.expander("PartnerID tidak ditemukan di Data Sponsor"):
            st.write(sorted(not_found_ids))

    st.success("Selesai!")

    st.download_button(
        "Download Hasil",
        result_bytes,
        "Crystal_Report_FILLED.xlsx"
    )
else:
    st.info("Silakan upload kedua file untuk memulai.")
