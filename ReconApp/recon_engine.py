# ======================================================
# STREAMLIT RECON-FILE GENERATOR (NO LOGIC MODIFIED)
# ======================================================

import streamlit as st
import pandas as pd
from pathlib import Path
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
import openpyxl.utils
import time

# ------------------------------------------------------
# Streamlit page setup
# ------------------------------------------------------
st.set_page_config(
    page_title="Recon File Generator",
    layout="wide"
)

st.title("🧾 EE Recon File Generator")
st.write("Upload the Trial Balance and Entries, input ICP Code, and generate the full Recon Workbook.")

# ------------------------------------------------------
# REQUIRED INTERNAL FILES (NOT uploaded by user)
# ------------------------------------------------------
MAPPING_FILE = Path("ReconApp/static/mapping.xlsx")
PLC_FILE = Path("ReconApp/static/PLC.xlsx")

if not MAPPING_FILE.exists():
    st.error("❌ mapping.xlsx not found in the app folder. Please upload it once at deployment.")
    st.stop()

if not PLC_FILE.exists():
    st.error("❌ PLC.xlsx not found in the app folder. Please upload it once at deployment.")
    st.stop()

# ------------------------------------------------------
# USER INPUTS
# ------------------------------------------------------

st.header("1. Upload Required Files")

trial_balance_upload = st.file_uploader(
    "Upload Trial Balance",
    type=["xlsx"],
    accept_multiple_files=False
)

entries_upload = st.file_uploader(
    "Upload Entries",
    type=["xlsx"],
    accept_multiple_files=False
)

st.header("2. Enter ICP Code")
ICP = st.text_input("ICP Code (e.g. SKPVAB)").strip().upper()

# ------------------------------------------------------
# LOAD INTERNAL FILES
# ------------------------------------------------------
@st.cache_data
def load_internal_mapping():
    book = pd.read_excel(MAPPING_FILE, sheet_name=None)

    if "account_mapping" not in book or "mapping_directory" not in book:
        raise KeyError("mapping.xlsx must contain 'account_mapping' and 'mapping_directory' sheets.")

    return book["account_mapping"], book["mapping_directory"]

@st.cache_data
def load_internal_plc():
    return pd.read_excel(PLC_FILE, engine="openpyxl")

mapping_accounts_df, mapping_dir_df = load_internal_mapping()
plc_df = load_internal_plc()

# ------------
# Workbook function
# -----------

 # === CREATE WORKBOOK ===
def build_workbook(trial_balance, entries, map_dir, ICP):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # === TRIAL BALANCE TAB ===
    ws_tb = wb.create_sheet("Trial Balance (YTD)", 0)

    # Write headers
    for c_idx, col in enumerate(trial_balance.columns, 1):
        ws_tb.cell(1, c_idx, col).font = Font(bold=True)

    # Write data
    for r_idx, row in enumerate(trial_balance.itertuples(index=False), 2):
        for c_idx, val in enumerate(row, 1):
            ws_tb.cell(r_idx, c_idx, val)

    # Autofit
    for col in ws_tb.columns:
        max_len = max((len(str(c.value)) if c.value else 0 for c in col), default=0)
        ws_tb.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = max_len + 2

    # === ACCOUNT OVERVIEW SECTION ===
    sheet_order = list(map_dir["sheet"].unique()) + ["Unmapped"]
    cols_needed = ["Posting Date", "Description", "Document No.", "ICP CODE", "GAAP Code", "Amount (LCY)"]
    cols_present = [c for c in cols_needed if c in entries.columns]

    sheet_status = {}
    account_anchor = {}  # {acc_no: (sheetname, row)}

    for sheet_name in sheet_order:

        subset = trial_balance[trial_balance["sheet_group"] == sheet_name]
        if subset.empty:
            continue

        ws = wb.create_sheet(sheet_name[:31])
        row_cursor = 3
        sheet_mismatch = False
        account_count = 0

        # Loop accounts
        for _, tb in subset.sort_values("No.").iterrows():

            acc_no = str(tb["No."])
            acc_name = tb.get("Name", "")
            tb_bal = tb.get("Balance at Date", 0.0)

            acc_df = entries[entries["G/L Account No."] == acc_no].copy()
            if acc_df.empty:
                continue

            account_count += 1
            acc_df = remove_internal_zeroes(acc_df)

            # ========= SPECIAL CASE 731000 (Totals per ICP) ==========
            if acc_no == "731000":
                acc_df = (
                    acc_df.groupby("ICP CODE", as_index=False)["Amount (LCY)"].sum()
                )
                acc_df["Description"] = acc_df["ICP CODE"]
                acc_df["Document No."] = ""
                acc_df["GAAP Code"] = ""
                acc_df = acc_df[["Description", "Document No.", "ICP CODE", "GAAP Code", "Amount (LCY)"]]

                net_sum = round(acc_df["Amount (LCY)"].sum(), 2)

                header = ws.cell(row=row_cursor, column=1, value=f"{acc_no} - {acc_name}")
                account_anchor[acc_no] = (ws.title, row_cursor)
                header.fill = green_fill if abs(net_sum - tb_bal) <= tolerance else red_fill

                if abs(net_sum - tb_bal) > tolerance:
                    sheet_mismatch = True

                row_cursor += 1
                block_start = row_cursor

                # Column headers
                cols_731 = ["Description", "Document No.", "ICP CODE", "GAAP Code", "Amount (LCY)"]
                for c_idx, col in enumerate(cols_731, 1):
                    ws.cell(row=row_cursor, column=c_idx, value=col).font = Font(bold=True)
                    ws.cell(row=row_cursor, column=c_idx).fill = header_fill
                row_cursor += 1

                # Entries
                for _, e in acc_df.iterrows():
                    for c_idx, col in enumerate(cols_731, 1):
                        ws.cell(row=row_cursor, column=c_idx, value=e.get(col, ""))
                        ws.cell(row=row_cursor, column=c_idx).fill = entry_fill
                    row_cursor += 1

                # Totals
                ws.cell(row=row_cursor, column=4, value="Account Total").font = Font(bold=True)
                ws.cell(row=row_cursor, column=5, value=net_sum).font = Font(bold=True)
                for c in range(1, 6):
                    ws.cell(row=row_cursor, column=c).fill = total_fill

                apply_borders(ws, block_start, row_cursor, 1, 5)
                row_cursor += 3
                continue

            # ========= SPECIAL CASE 321000 (same logic as 731000) ==========
            if acc_no == "321000":
                acc_df = (
                    acc_df.groupby("ICP CODE", as_index=False)["Amount (LCY)"].sum()
                )
                acc_df["Description"] = acc_df["ICP CODE"]
                acc_df["Document No."] = ""
                acc_df["GAAP Code"] = ""
                acc_df = acc_df[["Description", "Document No.", "ICP CODE", "GAAP Code", "Amount (LCY)"]]

                net_sum = round(acc_df["Amount (LCY)"].sum(), 2)

                header = ws.cell(row=row_cursor, column=1, value=f"{acc_no} - {acc_name}")
                account_anchor[acc_no] = (ws.title, row_cursor)
                header.fill = green_fill if abs(net_sum - tb_bal) <= tolerance else red_fill

                if abs(net_sum - tb_bal) > tolerance:
                    sheet_mismatch = True

                row_cursor += 1
                block_start = row_cursor

                cols_321 = ["Description", "Document No.", "ICP CODE", "GAAP Code", "Amount (LCY)"]
                for c_idx, col in enumerate(cols_321, 1):
                    ws.cell(row=row_cursor, column=c_idx, value=col).font = Font(bold=True)
                    ws.cell(row=row_cursor, column=c_idx).fill = header_fill
                row_cursor += 1

                for _, e in acc_df.iterrows():
                    for c_idx, col in enumerate(cols_321, 1):
                        ws.cell(row=row_cursor, column=c_idx, value=e.get(col, ""))
                        ws.cell(row=row_cursor, column=c_idx).fill = entry_fill
                    row_cursor += 1

                ws.cell(row=row_cursor, column=4, value="Account Total").font = Font(bold=True)
                ws.cell(row=row_cursor, column=5, value=net_sum).font = Font(bold=True)
                for c in range(1, 6):
                    ws.cell(row=row_cursor, column=c).fill = total_fill

                apply_borders(ws, block_start, row_cursor, 1, 5)
                row_cursor += 3
                continue

            # ========= SPECIAL CASE 390000–399999 (Totals-only with note) ==========
            if acc_no.isdigit() and 390000 <= int(acc_no) <= 399999:
                net_sum = round(acc_df["Amount (LCY)"].sum(), 2)

                header = ws.cell(row=row_cursor, column=1, value=f"{acc_no} - {acc_name}")
                account_anchor[acc_no] = (ws.title, row_cursor)
                header.fill = green_fill if abs(net_sum - tb_bal) <= tolerance else red_fill

                if abs(net_sum - tb_bal) > tolerance:
                    sheet_mismatch = True

                row_cursor += 1
                block_start = row_cursor

                # Row 1 — labels
                ws.cell(row=row_cursor, column=1, value="Note").font = Font(bold=True)
                ws.cell(row=row_cursor, column=6, value="Account Total").font = Font(bold=True)
                for c in range(1, 7):
                    ws.cell(row=row_cursor, column=c).fill = total_fill
                row_cursor += 1

                # Row 2 — content
                ws.cell(row=row_cursor, column=1, value="(See documentation)")
                ws.cell(row=row_cursor, column=6, value=net_sum)
                for c in range(1, 7):
                    ws.cell(row=row_cursor, column=c).fill = entry_fill

                apply_borders(ws, block_start, row_cursor, 1, 6)
                row_cursor += 3
                continue

            # ========= NORMAL ACCOUNTS ==========
            acc_df = acc_df.sort_values("Posting Date", ascending=False)

            net_sum = round(acc_df["Amount (LCY)"].sum(), 2)

            header = ws.cell(row=row_cursor, column=1, value=f"{acc_no} - {acc_name}")
            account_anchor[acc_no] = (ws.title, row_cursor)
            header.fill = green_fill if abs(net_sum - tb_bal) <= tolerance else red_fill

            if abs(net_sum - tb_bal) > tolerance:
                sheet_mismatch = True

            row_cursor += 1
            block_start = row_cursor

            # Column headers
            for c_idx, col in enumerate(cols_present, 1):
                ws.cell(row=row_cursor, column=c_idx, value=col).font = Font(bold=True)
                ws.cell(row=row_cursor, column=c_idx).fill = header_fill
            row_cursor += 1

            # Entries
            for _, e in acc_df.iterrows():
                for c_idx, col in enumerate(cols_present, 1):
                    ws.cell(row=row_cursor, column=c_idx, value=e.get(col, ""))
                    ws.cell(row=row_cursor, column=c_idx).fill = entry_fill
                row_cursor += 1

            # Totals
            ws.cell(row=row_cursor, column=len(cols_present)-1, value="Account Total").font = Font(bold=True)
            ws.cell(row=row_cursor, column=len(cols_present),   value=net_sum).font = Font(bold=True)

            for c in range(1, len(cols_present) + 1):
                ws.cell(row=row_cursor, column=c).fill = total_fill

            apply_borders(ws, block_start, row_cursor, 1, len(cols_present))
            row_cursor += 3

        sheet_status[sheet_name] = {
            "mismatches": int(sheet_mismatch),
            "accounts": account_count
        }

    return wb, sheet_status, account_anchor


# ------------------------------------------------------
# BUTTON
# ------------------------------------------------------

generate = st.button("Generate Recon File", type="primary")

if generate:
    if trial_balance_upload is None or entries_upload is None or ICP == "":
        st.error("Please upload both files and input ICP Code.")
        st.stop()

    with st.spinner("Generating Recon File… please wait, this may take up to 20 seconds…"):

        # LOAD USER FILES
        trial_balance_df = pd.read_excel(trial_balance_upload)
        entries_df = pd.read_excel(entries_upload)

        # The next section calls your full original logic…
        # The function is defined in MESSAGE 2
        try:
            output_bytes = build_recon_workbook(
                trial_balance_df,
                entries_df,
                mapping_accounts_df,
                mapping_dir_df,
                plc_df,
                ICP
            )
        except Exception as e:
            st.error(f"❌ Error during generation: {e}")
            raise

        st.success("Recon File Generated Successfully!")

        st.download_button(
            label="⬇️ Download Reconciliation_Mapped.xlsx",
            data=output_bytes,
            file_name="Reconciliation_Mapped.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# === LOAD DATA (Streamlit will pass file-like objects instead of paths) ===
def load_input_files(trial_balance_file, entries_file, mapping_file, icp_code):
    """
    Returns:
        trial_balance, entries, mapping_dir, acct_to_code, code_to_meta, ICP
    """
    ICP = icp_code.strip().upper()

    trial_balance = pd.read_excel(trial_balance_file)
    entries = pd.read_excel(entries_file)

    acct_to_code, code_to_meta, map_dir = load_mapping(mapping_file)
    trial_balance = apply_mapping(trial_balance, acct_to_code, code_to_meta)

    # Normalize
    entries.rename(
        columns=lambda c: "Amount (LCY)" if str(c).strip().lower() in ["amount", "amount (lcy)"] else c,
        inplace=True
    )
    entries.rename(
        columns=lambda c: "ICP CODE" if str(c).strip().lower() == "icp code" else c,
        inplace=True
    )

    trial_balance["Balance at Date"] = trial_balance["Balance at Date"].apply(to_float)
    entries["G/L Account No."] = entries["G/L Account No."].apply(normalize_account)
    entries["Amount (LCY)"] = entries["Amount (LCY)"].apply(to_float)
    entries["Posting Date"] = pd.to_datetime(entries["Posting Date"], errors="coerce").dt.date

    return trial_balance, entries, acct_to_code, code_to_meta, map_dir, ICP


# === EXCEL STYLES ===
thin = Side(border_style="thin", color="000000")
thick = Side(border_style="medium", color="000000")

green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
red_fill   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
header_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
entry_fill  = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
total_fill  = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")


# === BORDER DRAWING (unchanged) ===
def apply_borders(ws, top, bottom, left, right):
    for r in range(top, bottom + 1):
        for c in range(left, right + 1):
            ws.cell(r, c).border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Thick borders
    for c in range(left, right + 1):
        ws.cell(top, c).border    = Border(top=thick, left=thin, right=thin, bottom=thin)
        ws.cell(bottom, c).border = Border(bottom=thick, left=thin, right=thin, top=thin)

    for r in range(top, bottom + 1):
        ws.cell(r, left).border  = Border(left=thick, top=thin, bottom=thin, right=thin)
        ws.cell(r, right).border = Border(right=thick, top=thin, bottom=thin, left=thin)

    # Corners
    ws.cell(top, left).border     = Border(top=thick, left=thick, right=thin, bottom=thin)
    ws.cell(top, right).border    = Border(top=thick, right=thick, left=thin, bottom=thin)
    ws.cell(bottom, left).border  = Border(bottom=thick, left=thick, right=thin, top=thin)
    ws.cell(bottom, right).border = Border(bottom=thick, right=thick, left=thin, top=thin)


# === INTERNAL ZEROING (unchanged) ===
def remove_internal_zeroes(df, tol=0.01):
    df = df.sort_values("Posting Date", ascending=True).reset_index(drop=True)
    keep = [True] * len(df)

    for i in range(len(df)):
        if not keep[i]:
            continue
        for j in range(i + 1, len(df)):
            same_keys = (
                df.loc[i, "Document No."] == df.loc[j, "Document No."]
                and df.loc[i, "ICP CODE"] == df.loc[j, "ICP CODE"]
                and df.loc[i, "GAAP Code"] == df.loc[j, "GAAP Code"]
            )

            if same_keys and keep[j] and abs(df.loc[i, "Amount (LCY)"] + df.loc[j, "Amount (LCY)"]) <= tol:
                keep[i] = keep[j] = False
                break

    df = df[keep].copy()

    df["cum"] = df["Amount (LCY)"].cumsum().round(2)
    zero_indices = df.index[abs(df["cum"]) <= tol].tolist()

    if zero_indices:
        df = df.loc[zero_indices[-1] + 1:].copy()

    return df


# -----------------------
# ✅ MESSAGE 3 — Finalize workbook: Frontpage, save, color tabs
# -----------------------

def finalize_and_save(trial_balance, entries, map_dir, ICP,
                      build_fn=build_workbook,
                      plc_filename="PLC.xlsx"):
    """
    Build full workbook (calls build_workbook), add front page, format numbers/dates,
    save workbook and color sheet tabs with xlwings.
    Returns path to saved file.
    """
    # 1) Build sheets for accounts
    wb, sheet_status, account_anchor = build_fn(trial_balance, entries, map_dir, ICP)

    # 2) FRONT PAGE (PLC header cards + autogenerated comments + hyperlinks)
    from warnings import filterwarnings
    filterwarnings("ignore", message="Data Validation extension is not supported and will be removed")

    ws_front = wb.create_sheet("Frontpage", 0)
    ws_front["A1"] = "EE Reconciliation Overview"
    ws_front["A1"].font = Font(size=16, bold=True)

    # Load PLC (optional but expected)
    plc_path = Path(folder) / plc_filename
    plc_norm = None
    try:
        plc_df = pd.read_excel(plc_path, engine="openpyxl")
        plc_norm = plc_df.copy()
        plc_norm.columns = [c.strip() for c in plc_norm.columns]
        required = ["ICP code", "Company name", "Accountant", "Controller"]
        missing_cols = [c for c in required if c not in plc_norm.columns]
        if missing_cols:
            raise KeyError(f"Missing columns in PLC.xlsx: {missing_cols}")
        plc_norm["ICP code_norm"] = plc_norm["ICP code"].astype(str).str.strip().str.upper()
        plc_norm["Company name"]   = plc_norm["Company name"].astype(str).str.strip()
        plc_norm["Accountant"]     = plc_norm["Accountant"].astype(str).str.strip()
        plc_norm["Controller"]     = plc_norm["Controller"].astype(str).str.strip()
    except FileNotFoundError:
        plc_norm = None
    except Exception as e:
        plc_norm = None

    # Add PLC cards for selected ICP(s)
    selected_icps = [ICP]
    row_ptr = 3
    for icp in selected_icps:
        icp_key = str(icp).strip().upper()
        row = plc_norm.loc[plc_norm["ICP code_norm"] == icp_key] if plc_norm is not None else pd.DataFrame()
        ws_front.cell(row_ptr, 1, "ICP code").font = Font(bold=True)
        ws_front.cell(row_ptr, 2, icp_key)

        ws_front.cell(row_ptr + 1, 1, "Company name").font = Font(bold=True)
        ws_front.cell(row_ptr + 2, 1, "Accountant").font   = Font(bold=True)
        ws_front.cell(row_ptr + 3, 1, "Controller").font   = Font(bold=True)

        if not row.empty:
            ws_front.cell(row_ptr + 1, 2, row.iloc[0]["Company name"])
            ws_front.cell(row_ptr + 2, 2, row.iloc[0]["Accountant"])
            ws_front.cell(row_ptr + 3, 2, row.iloc[0]["Controller"])
        else:
            ws_front.cell(row_ptr + 1, 2, "Not found in PLC.xlsx")
            ws_front.cell(row_ptr + 2, 2, "—")
            ws_front.cell(row_ptr + 3, 2, "—")

        apply_borders(ws_front, top=row_ptr, bottom=row_ptr+3, left=1, right=2)
        row_ptr += 6

    row_ptr += 1
    ws_front.cell(row_ptr, 1, "Automatically generated comments:").font = Font(bold=True, underline="single")
    row_ptr += 2

    # Build quick checks (same as earlier logic)
    comments = []
    mask_200_399 = (
        trial_balance["No."].astype(str).str.isdigit() &
        trial_balance["No."].astype(int).between(200000, 399999)
    )
    negatives = trial_balance.loc[
        mask_200_399 & (trial_balance["Balance at Date"] < 0),
        ["No.", "Name", "Balance at Date"]
    ]

    mask_400_plus = (
        trial_balance["No."].astype(str).str.isdigit() &
        (trial_balance["No."].astype(int) >= 400000)
    )
    positives = trial_balance.loc[
        mask_400_plus & (trial_balance["Balance at Date"] > 0),
        ["No.", "Name", "Balance at Date"]
    ]

    if not negatives.empty:
        comments.append(f"{len(negatives)} account(s) in the 200000–399999 range have negative balances.")
    if not positives.empty:
        comments.append(f"{len(positives)} account(s) in the 400000+ range have positive balances.")

    if 'sheet_status' in locals():
        mismatched_sheets = sum(v['mismatches'] for v in sheet_status.values())
        if mismatched_sheets > 0:
            comments.append(f"{mismatched_sheets} sheet(s) contain out-of-balance accounts exceeding tolerance {tolerance}.")
        else:
            comments.append("All sheets appear balanced within tolerance limits.")

    if not comments:
        comments.append("No issues detected based on configured checks.")

    # Write comments
    start_row = row_ptr
    for i, comment in enumerate(comments, start=start_row):
        ws_front.cell(i, 1, f"• {comment}")
    row_ptr = start_row + len(comments) + 2

    # Helper to set hyperlinks to account anchors
    def set_hyperlink(cell, acc_no):
        acc = str(acc_no)
        if acc in account_anchor:
            sheet_name, anchor_row = account_anchor[acc]
            sheet_ref = f"'{sheet_name}'" if not sheet_name.isalnum() else sheet_name
            cell.hyperlink = f"#{sheet_ref}!A{anchor_row}"
            cell.style = "Hyperlink"

    # Detailed lists
    if not negatives.empty:
        ws_front.cell(row_ptr, 1, "Negative balances (200000–399999):").font = Font(bold=True)
        row_ptr += 1
        for _, r in negatives.iterrows():
            acc = str(r["No."])
            c = ws_front.cell(row_ptr, 1, acc)
            set_hyperlink(c, acc)
            ws_front.cell(row_ptr, 2, r.get("Name", ""))
            val_cell = ws_front.cell(row_ptr, 3, r["Balance at Date"])
            val_cell.number_format = "#,##0.00"
            row_ptr += 1

    if not positives.empty:
        ws_front.cell(row_ptr, 1, "Positive balances (400000+):").font = Font(bold=True)
        row_ptr += 1
        for _, r in positives.iterrows():
            acc = str(r["No."])
            c = ws_front.cell(row_ptr, 1, acc)
            set_hyperlink(c, acc)
            ws_front.cell(row_ptr, 2, r.get("Name", ""))
            val_cell = ws_front.cell(row_ptr, 3, r["Balance at Date"])
            val_cell.number_format = "#,##0.00"
            row_ptr += 1

    # Footer metadata
    row_ptr += 2
    ws_front.cell(row_ptr, 1, f"Generated on: {pd.Timestamp.now():%Y-%m-%d %H:%M}")
    ws_front.cell(row_ptr + 1, 1, f"Tolerance used: {tolerance}")
    ws_front.cell(row_ptr + 2, 1, f"Source files located in: {folder}")

    # 3) FORMAT numeric cells (comma thousands) and dates across all sheets
    amount_fmt = "#,##0.00"
    date_fmt = "yyyy-mm-dd"
    for ws in wb.worksheets:
        # apply number formats to cells that look like amounts or dates
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                # Date-like (we stored Posting Date as datetime.date earlier)
                if isinstance(cell.value, (pd.Timestamp,)) or (hasattr(cell.value, "year") and isinstance(cell.value, (int,)) is False and getattr(cell.value, "isoformat", None)):
                    # Try to coerce to date string if it's pandas Timestamp or date-like
                    try:
                        # If it's a date object, set date format
                        if isinstance(cell.value, (pd.Timestamp,)) or getattr(cell.value, "isoformat", None):
                            # Only set format if cell contains a date object
                            try:
                                pd.to_datetime(cell.value)
                                cell.number_format = date_fmt
                            except:
                                pass
                    except Exception:
                        pass
                # Numbers: use amount format for columns named "Amount" or values that are numeric
                if isinstance(cell.value, (int, float)):
                    # Heuristic: columns with "Amount" in header or numeric cell in last column
                    cell.number_format = amount_fmt

    # 4) Auto-fit columns & hide gridlines
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        for col in ws.columns:
            max_len = max((len(str(c.value)) if c.value else 0 for c in col), default=0)
            ws.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = max_len + 2

    print(f"✅ Workbook created successfully:\n{saved_path}")
    return str(saved_path)


# ============================================================
# PUBLIC FUNCTION CALLED FROM STREAMLIT
# ============================================================

def generate_reconciliation_file(trial_balance_file, entries_file, icp_code):
    """
    Streamlit passes uploaded files (file-like objects) to this function.
    This function returns the Excel file as bytes.
    """

    # 1. Load trial balance & entries (from uploaded files)
    trial_balance_df = pd.read_excel(trial_balance_file)
    entries_df = pd.read_excel(entries_file)

    # 2. Load internal mapping files
    mapping_accounts_df = pd.read_excel("ReconApp/static/mapping.xlsx", sheet_name="account_mapping")
    mapping_dir_df      = pd.read_excel("ReconApp/static/mapping.xlsx", sheet_name="mapping_directory")

    plc_df = pd.read_excel("ReconApp/static/PLC.xlsx")

    # 3. Build workbook in memory
    output_path = "Reconciliation_Mapped.xlsx"

    saved_path = finalize_and_save(
        trial_balance_df,
        entries_df,
        mapping_dir_df,
        icp_code,
        build_fn=build_recon_workbook,
        plc_filename="PLC.xlsx",
        output_path=output_path
    )

    # 4. Return file bytes to Streamlit
    with open(saved_path, "rb") as f:
        return f.read()





