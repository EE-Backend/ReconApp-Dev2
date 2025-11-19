# recon_engine.py
import pandas as pd
from pathlib import Path
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
import openpyxl.utils
import time

# === CONFIG / Defaults ===
TOLERANCE = 0.001

# By default mapping and plc files are expected in ./static/
BASE_DIR = Path(__file__).parent
STATIC_DIR = BASE_DIR / "static"
DEFAULT_MAPPING = STATIC_DIR / "mapping.xlsx"
DEFAULT_PLC = STATIC_DIR / "PLC.xlsx"


# === HELPERS ===
def normalize_account(val):
    if pd.isna(val):
        return ""
    return "".join(ch for ch in str(val) if ch.isdigit())


def to_float(val):
    if pd.isna(val):
        return 0.0
    val = str(val).replace(",", ".").replace(" ", "")
    try:
        return round(float(val), 2)
    except:
        return 0.0


def _normalize_code(x):
    """Return consistent code strings (e.g., 101, 101.0, ' 101 ') -> '101'."""
    if pd.isna(x):
        return None
    try:
        f = float(x)
        return str(int(f)) if f.is_integer() else str(f).rstrip("0").rstrip(".")
    except:
        s = str(x).strip()
        if s.replace(".", "", 1).isdigit():
            try:
                f = float(s)
                return str(int(f)) if f.is_integer() else s
            except:
                pass
        return s


# === MAPPING LOAD/APPLY ===
def load_mapping(mapping_path=None):
    """
    Returns: acct_to_code (dict), code_to_meta (dict), map_dir (DataFrame)
    Expects mapping_path to be an Excel workbook with sheets:
      - account_mapping (with columns: Account no., Mapping)
      - mapping_directory (with columns: code, header, sheet)
    """
    mapping_path = Path(mapping_path) if mapping_path else DEFAULT_MAPPING
    if not mapping_path.exists():
        raise FileNotFoundError(f"Mapping file not found at {mapping_path}")

    book = pd.read_excel(mapping_path, sheet_name=None)

    if "account_mapping" not in book or "mapping_directory" not in book:
        raise KeyError("mapping.xlsx must include sheets 'account_mapping' and 'mapping_directory'")

    map_accounts = book["account_mapping"].rename(columns={"Mapping": "code"}).copy()
    map_dir = book["mapping_directory"].copy()

    # Normalize
    map_accounts["Account no."] = map_accounts["Account no."].astype(str).str.strip().apply(normalize_account)
    map_accounts["code"] = map_accounts["code"].apply(_normalize_code)

    map_dir["code"] = map_dir["code"].apply(_normalize_code)
    map_dir["header"] = map_dir["header"].astype(str).str.strip()
    map_dir["sheet"] = map_dir["sheet"].astype(str).str.strip()

    acct_to_code = map_accounts.set_index("Account no.")["code"].to_dict()
    code_to_meta = map_dir.set_index("code")[["header", "sheet"]].to_dict(orient="index")

    return acct_to_code, code_to_meta, map_dir


def apply_mapping(trial_balance, acct_to_code, code_to_meta, bank_code="19", bank_ranges=((390000, 399999),)):
    tb = trial_balance.copy()
    tb["No."] = tb["No."].astype(str).str.strip().apply(normalize_account)
    tb["code"] = tb["No."].map(acct_to_code)

    def _in_ranges(acc_str, ranges):
        return bool(acc_str) and acc_str.isdigit() and any(lo <= int(acc_str) <= hi for lo, hi in ranges)

    unmapped = tb["code"].isna()
    tb.loc[unmapped & tb["No."].apply(lambda s: _in_ranges(s, bank_ranges)), "code"] = bank_code

    tb["header"] = tb["code"].map(lambda c: code_to_meta.get(c, {}).get("header"))
    tb["sheet_group"] = tb["code"].map(lambda c: code_to_meta.get(c, {}).get("sheet"))
    tb["sheet_group"] = tb["sheet_group"].fillna("Unmapped").astype(str).str.strip()
    return tb


# === STYLES / BORDERS / FORMATS ===
thin = Side(border_style="thin", color="000000")
thick = Side(border_style="medium", color="000000")

green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
header_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")  # header darker
entry_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")   # lighter
total_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")    # totals


def apply_borders(ws, top, bottom, left, right):
    """
    Draw neat rectangle with thick external lines and thin internal lines
    Ensures thick border above/below column A and last Amount column as requested.
    """
    # thin grid everywhere in the block
    for r in range(top, bottom + 1):
        for c in range(left, right + 1):
            ws.cell(r, c).border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # thick top and bottom across the full block
    for c in range(left, right + 1):
        top_cell = ws.cell(top, c)
        bottom_cell = ws.cell(bottom, c)
        top_cell.border = Border(top=thick, left=top_cell.border.left, right=top_cell.border.right, bottom=top_cell.border.bottom)
        bottom_cell.border = Border(bottom=thick, left=bottom_cell.border.left, right=bottom_cell.border.right, top=bottom_cell.border.top)

    # thick left & right sides
    for r in range(top, bottom + 1):
        left_cell = ws.cell(r, left)
        right_cell = ws.cell(r, right)
        left_cell.border = Border(left=thick, top=left_cell.border.top, bottom=left_cell.border.bottom, right=left_cell.border.right)
        right_cell.border = Border(right=thick, top=right_cell.border.top, bottom=right_cell.border.bottom, left=right_cell.border.left)

    # corners: ensure they have both thick sides
    ws.cell(top, left).border = Border(top=thick, left=thick, right=thin, bottom=thin)
    ws.cell(top, right).border = Border(top=thick, right=thick, left=thin, bottom=thin)
    ws.cell(bottom, left).border = Border(bottom=thick, left=thick, right=thin, top=thin)
    ws.cell(bottom, right).border = Border(bottom=thick, right=thick, left=thin, top=thin)


# === INTERNAL ZEROING ===
def remove_internal_zeroes(df, tol=TOLERANCE):
    """
    Remove internal cancelling entries when:
      - ICP CODE is the same (or empty/NaN in both lines), and
      - GAAP Code is the same (or empty/NaN in both lines), and
      - Amounts sum to ~0 within tolerance.

    Document No. is intentionally ignored.

    Then, within each (ICP, GAAP) bucket, perform cumulative zero-block
    trimming: drop entries up to the last point where that bucket's
    cumulative sum returns to (approx.) zero.
    """
    if df.empty:
        return df

    # Sort chronologically
    df = df.sort_values("Posting Date", ascending=True).reset_index(drop=True)

    def _norm_key(x):
        """Normalize ICP/GAAP for grouping & equality: treat NaN/empty as ''."""
        if pd.isna(x):
            return ""
        return str(x).strip()

    # Normalized keys used for both pairwise and cumulative logic
    df["_icp_norm"] = df["ICP CODE"].apply(_norm_key) if "ICP CODE" in df.columns else ""
    df["_gaap_norm"] = df["GAAP Code"].apply(_norm_key) if "GAAP Code" in df.columns else ""

    # ---------- 1) Pairwise zero removal within same ICP/GAAP ----------
    keep = [True] * len(df)
    for i in range(len(df)):
        if not keep[i]:
            continue
        for j in range(i + 1, len(df)):
            if not keep[j]:
                continue

            same_icp = (df.loc[i, "_icp_norm"] == df.loc[j, "_icp_norm"])
            same_gaap = (df.loc[i, "_gaap_norm"] == df.loc[j, "_gaap_norm"])

            if same_icp and same_gaap and abs(df.loc[i, "Amount (LCY)"] + df.loc[j, "Amount (LCY)"]) < tol:
                # Exact opposite pair within same ICP/GAAP bucket -> drop both
                keep[i] = keep[j] = False
                break

    df = df[keep].copy()

    # ---------- 2) Cumulative zero trimming per (ICP, GAAP) bucket ----------
    def _trim_group(g):
        if g.empty:
            return g
        g = g.copy()
        g["cum"] = g["Amount (LCY)"].cumsum().round(2)
        zero_idx = g.index[g["cum"].abs() < tol].tolist()
        if zero_idx:
            last_zero = zero_idx[-1]
            g = g.loc[g.index > last_zero]
        g = g.drop(columns=["cum"])
        return g

    df = df.groupby(["_icp_norm", "_gaap_norm"], group_keys=False).apply(_trim_group)

    # Clean up helper columns and reindex
    df = df.drop(columns=["_icp_norm", "_gaap_norm"], errors="ignore").reset_index(drop=True)

    return df



# === WORKBOOK BUILDING ===
def build_workbook(trial_balance_df, entries_df, map_dir, acct_to_code, code_to_meta, ICP, tolerance=TOLERANCE):
    """
    Returns: openpyxl.Workbook object, sheet_status dict, account_anchor dict, mismatch_accounts list
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # === TRIAL BALANCE TAB ===
    ws_tb = wb.create_sheet("Trial Balance (YTD)", 0)
    for c_idx, col in enumerate(trial_balance_df.columns, 1):
        ws_tb.cell(1, c_idx, col).font = Font(bold=True)
    for r_idx, row in enumerate(trial_balance_df.itertuples(index=False), 2):
        for c_idx, val in enumerate(row, 1):
            ws_tb.cell(r_idx, c_idx, val)

    # Autofit for trial balance
    for col in ws_tb.columns:
        max_len = max((len(str(c.value)) if c.value else 0 for c in col), default=0)
        ws_tb.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = max_len + 2

    # === ACCOUNT OVERVIEW (mapped groups) ===
    sheet_order = list(map_dir["sheet"].unique()) + ["Unmapped"]
    cols_needed = ["Posting Date", "Description", "Document No.", "ICP CODE", "GAAP Code", "Amount (LCY)"]
    cols_present = [c for c in cols_needed if c in entries_df.columns]

    sheet_status = {}
    account_anchor = {}
    mismatch_accounts = []  # NEW

    # iterate mapping order
    for sheet_name in sheet_order:
        subset = trial_balance_df[trial_balance_df["sheet_group"] == sheet_name]
        if subset.empty:
            continue

        ws = wb.create_sheet(sheet_name[:31])
        row_cursor = 3
        sheet_mismatch = False
        account_count = 0

        for _, tb in subset.sort_values("No.").iterrows():
            acc_no = str(tb["No."])
            acc_name = tb.get("Name", "")
            tb_bal = tb.get("Balance at Date", 0.0)

            acc_df = entries_df[entries_df["G/L Account No."] == acc_no].copy()
            if acc_df.empty:
                continue

            account_count += 1

            # Year trimming: drop prior years that net to 0 (keep latest year always)
            acc_df = acc_df.sort_values("Posting Date", ascending=True).reset_index(drop=True)
            acc_df["Year"] = pd.to_datetime(acc_df["Posting Date"], errors="coerce").dt.year.fillna(0).astype(int)
            if acc_df["Year"].nunique() > 1:
                years = sorted(acc_df["Year"].unique())
                for y in years[:-1]:
                    y_sum = round(acc_df.loc[acc_df["Year"] == y, "Amount (LCY)"].sum(), 2)
                    if abs(y_sum) < tolerance:
                        acc_df = acc_df[acc_df["Year"] != y]

            if acc_df.empty:
                continue

            # remove internal zeroes (only when Document No., ICP and GAAP match)
            acc_df = remove_internal_zeroes(acc_df)

            if acc_df.empty:
                continue

            # Special: account 731000 (show totals per ICP only)
            if acc_no == "731000":
                grouped = acc_df.groupby("ICP CODE", as_index=False)["Amount (LCY)"].sum()
                grouped["Description"] = grouped["ICP CODE"]
                grouped["Document No."] = ""
                grouped["GAAP Code"] = ""
                acc_view = grouped[["Description", "Document No.", "ICP CODE", "GAAP Code", "Amount (LCY)"]].copy()
                net_sum = round(acc_view["Amount (LCY)"].sum(), 2)

                header_cell = ws.cell(row=row_cursor, column=1, value=f"{acc_no} - {acc_name}")
                account_anchor[acc_no] = (ws.title, row_cursor)
                header_cell.fill = green_fill if abs(net_sum - tb_bal) <= tolerance else red_fill
                if abs(net_sum - tb_bal) > tolerance:
                    sheet_mismatch = True
                    mismatch_accounts.append({
                        "No": acc_no,
                        "Name": acc_name,
                        "tb_balance": tb_bal,
                        "entries_sum": net_sum,
                        "difference": round(net_sum - tb_bal, 2),
                    })

                row_cursor += 1
                block_start = row_cursor

                # Column headers
                for c_idx, col in enumerate(["Description", "Document No.", "ICP CODE", "GAAP Code", "Amount (LCY)"], 1):
                    ws.cell(row=row_cursor, column=c_idx, value=col).font = Font(bold=True)
                    ws.cell(row=row_cursor, column=c_idx).fill = header_fill
                row_cursor += 1

                # Rows
                for _, r in acc_view.iterrows():
                    for c_idx, col in enumerate(["Description", "Document No.", "ICP CODE", "GAAP Code", "Amount (LCY)"], 1):
                        cell = ws.cell(row=row_cursor, column=c_idx, value=r.get(col, ""))
                        cell.fill = entry_fill
                    row_cursor += 1

                # Totals row (Account Total above the number OR as requested)
                ws.cell(row=row_cursor, column=4, value="Account Total").font = Font(bold=True)
                vcell = ws.cell(row=row_cursor, column=5, value=net_sum)
                vcell.font = Font(bold=True)
                for c in range(1, 6):
                    ws.cell(row=row_cursor, column=c).fill = total_fill

                apply_borders(ws, block_start, row_cursor, 1, 5)
                row_cursor += 3
                continue

            # Special: 321000 same as 731000 (kept for compatibility)
            if acc_no == "321000":
                grouped = acc_df.groupby("ICP CODE", as_index=False)["Amount (LCY)"].sum()
                grouped["Description"] = grouped["ICP CODE"]
                grouped["Document No."] = ""
                grouped["GAAP Code"] = ""
                acc_view = grouped[["Description", "Document No.", "ICP CODE", "GAAP Code", "Amount (LCY)"]].copy()
                net_sum = round(acc_view["Amount (LCY)"].sum(), 2)

                header_cell = ws.cell(row=row_cursor, column=1, value=f"{acc_no} - {acc_name}")
                account_anchor[acc_no] = (ws.title, row_cursor)
                header_cell.fill = green_fill if abs(net_sum - tb_bal) <= tolerance else red_fill
                if abs(net_sum - tb_bal) > tolerance:
                    sheet_mismatch = True
                    mismatch_accounts.append({
                        "No": acc_no,
                        "Name": acc_name,
                        "tb_balance": tb_bal,
                        "entries_sum": net_sum,
                        "difference": round(net_sum - tb_bal, 2),
                    })

                row_cursor += 1
                block_start = row_cursor

                for c_idx, col in enumerate(["Description", "Document No.", "ICP CODE", "GAAP Code", "Amount (LCY)"], 1):
                    ws.cell(row=row_cursor, column=c_idx, value=col).font = Font(bold=True)
                    ws.cell(row=row_cursor, column=c_idx).fill = header_fill
                row_cursor += 1

                for _, r in acc_view.iterrows():
                    for c_idx, col in enumerate(["Description", "Document No.", "ICP CODE", "GAAP Code", "Amount (LCY)"], 1):
                        cell = ws.cell(row=row_cursor, column=c_idx, value=r.get(col, ""))
                        cell.fill = entry_fill
                    row_cursor += 1

                ws.cell(row=row_cursor, column=4, value="Account Total").font = Font(bold=True)
                vcell = ws.cell(row=row_cursor, column=5, value=net_sum)
                vcell.font = Font(bold=True)
                for c in range(1, 6):
                    ws.cell(row=row_cursor, column=c).fill = total_fill

                apply_borders(ws, block_start, row_cursor, 1, 5)
                row_cursor += 3
                continue

            # Special: bank/cash accounts 390000–399999 -> totals only with note (See documentation)
            if acc_no.isdigit() and 390000 <= int(acc_no) <= 399999:
                net_sum = round(acc_df["Amount (LCY)"].sum(), 2)

                header_cell = ws.cell(row=row_cursor, column=1, value=f"{acc_no} - {acc_name}")
                account_anchor[acc_no] = (ws.title, row_cursor)
                header_cell.fill = green_fill if abs(net_sum - tb_bal) <= tolerance else red_fill
                if abs(net_sum - tb_bal) > tolerance:
                    sheet_mismatch = True
                    mismatch_accounts.append({
                        "No": acc_no,
                        "Name": acc_name,
                        "tb_balance": tb_bal,
                        "entries_sum": net_sum,
                        "difference": round(net_sum - tb_bal, 2),
                    })

                row_cursor += 1
                block_start = row_cursor

                # First row: labels (Note header + Account Total label)
                ws.cell(row=row_cursor, column=1, value="Note").font = Font(bold=True)
                ws.cell(row=row_cursor, column=6, value="Account Total").font = Font(bold=True)
                for c in range(1, 7):
                    ws.cell(row=row_cursor, column=c).fill = total_fill
                row_cursor += 1

                # Second row: content - (See documentation) + total
                ws.cell(row=row_cursor, column=1, value="(See documentation)")
                vcell = ws.cell(row=row_cursor, column=6, value=net_sum)
                vcell.number_format = "#,##0.00"
                for c in range(1, 7):
                    ws.cell(row=row_cursor, column=c).fill = entry_fill

                apply_borders(ws, block_start, row_cursor, 1, 6)
                row_cursor += 3
                continue

            # Normal accounts: show full list newest -> oldest
            acc_df = acc_df.sort_values("Posting Date", ascending=False)
            net_sum = round(acc_df["Amount (LCY)"].sum(), 2)

            header_cell = ws.cell(row=row_cursor, column=1, value=f"{acc_no} - {acc_name}")
            account_anchor[acc_no] = (ws.title, row_cursor)
            header_cell.fill = green_fill if abs(net_sum - tb_bal) <= tolerance else red_fill
            if abs(net_sum - tb_bal) > tolerance:
                sheet_mismatch = True
                mismatch_accounts.append({
                    "No": acc_no,
                    "Name": acc_name,
                    "tb_balance": tb_bal,
                    "entries_sum": net_sum,
                    "difference": round(net_sum - tb_bal, 2),
                })

            row_cursor += 1
            block_start = row_cursor

            # Column headers
            for c_idx, col in enumerate(cols_present, 1):
                ws.cell(row=row_cursor, column=c_idx, value=col).font = Font(bold=True)
                ws.cell(row=row_cursor, column=c_idx).fill = header_fill
            row_cursor += 1

            # Rows: entries
            for _, e in acc_df.iterrows():
                for c_idx, col in enumerate(cols_present, 1):
                    val = e.get(col, "")
                    cell = ws.cell(row=row_cursor, column=c_idx, value=val)
                    cell.fill = entry_fill
                row_cursor += 1

            # Totals row
            ws.cell(row=row_cursor, column=len(cols_present) - 1, value="Account Total").font = Font(bold=True)
            total_cell = ws.cell(row=row_cursor, column=len(cols_present), value=net_sum)
            total_cell.font = Font(bold=True)
            for c in range(1, len(cols_present) + 1):
                ws.cell(row=row_cursor, column=c).fill = total_fill

            apply_borders(ws, block_start, row_cursor, 1, len(cols_present))
            row_cursor += 3

        sheet_status[sheet_name] = {"mismatches": int(sheet_mismatch), "accounts": account_count}

    return wb, sheet_status, account_anchor, mismatch_accounts


# === FINALIZE: front page, formatting, save to bytes ===
def finalize_workbook_to_bytes(
    wb,
    sheet_status,
    account_anchor,
    trial_balance_df,
    entries_df,
    ICP,
    plc_path=None,
    tolerance=TOLERANCE,
    mismatch_accounts=None,
):
    """
    Adds front page, number/date formatting, hides gridlines, autofit, and returns bytes buffer.
    """
    if mismatch_accounts is None:
        mismatch_accounts = []

    # FRONT PAGE
    from warnings import filterwarnings
    filterwarnings("ignore", message="Data Validation extension is not supported and will be removed")

    ws_front = wb.create_sheet("Frontpage", 0)
    ws_front["A1"] = "EE Reconciliation Overview"
    ws_front["A1"].font = Font(size=16, bold=True)

    # PLC (optional)
    plc_norm = None
    plc_path = Path(plc_path) if plc_path else DEFAULT_PLC
    try:
        if plc_path.exists():
            plc_df = pd.read_excel(plc_path, engine="openpyxl")
            plc_norm = plc_df.copy()
            plc_norm.columns = [c.strip() for c in plc_norm.columns]
            plc_norm["ICP code_norm"] = plc_norm["ICP code"].astype(str).str.strip().str.upper()
            plc_norm["Company name"] = plc_norm["Company name"].astype(str).str.strip()
            plc_norm["Accountant"] = plc_norm["Accountant"].astype(str).str.strip()
            plc_norm["Controller"] = plc_norm["Controller"].astype(str).str.strip()
    except Exception:
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
        ws_front.cell(row_ptr + 2, 1, "Accountant").font = Font(bold=True)
        ws_front.cell(row_ptr + 3, 1, "Controller").font = Font(bold=True)

        if not row.empty:
            ws_front.cell(row_ptr + 1, 2, row.iloc[0]["Company name"])
            ws_front.cell(row_ptr + 2, 2, row.iloc[0]["Accountant"])
            ws_front.cell(row_ptr + 3, 2, row.iloc[0]["Controller"])
        else:
            ws_front.cell(row_ptr + 1, 2, "Not found in PLC.xlsx")
            ws_front.cell(row_ptr + 2, 2, "—")
            ws_front.cell(row_ptr + 3, 2, "—")

        apply_borders(ws_front, top=row_ptr, bottom=row_ptr + 3, left=1, right=2)
        row_ptr += 6

    row_ptr += 1
    ws_front.cell(row_ptr, 1, "Automatically generated comments:").font = Font(bold=True, underline="single")
    row_ptr += 2

    # Quick checks for frontpage comments
    comments = []
    mask_200_399 = (
        trial_balance_df["No."].astype(str).str.isdigit() &
        trial_balance_df["No."].astype(int).between(200000, 399999)
    )
    negatives = trial_balance_df.loc[
        mask_200_399 & (trial_balance_df["Balance at Date"] < 0),
        ["No.", "Name", "Balance at Date"]
    ]

    mask_400_plus = (
        trial_balance_df["No."].astype(str).str.isdigit() &
        (trial_balance_df["No."].astype(int) >= 400000)
    )
    positives = trial_balance_df.loc[
        mask_400_plus & (trial_balance_df["Balance at Date"] > 0),
        ["No.", "Name", "Balance at Date"]
    ]

    if not negatives.empty:
        comments.append(f"{len(negatives)} account(s) in the 200000–399999 range have negative balances.")
    if not positives.empty:
        comments.append(f"{len(positives)} account(s) in the 400000+ range have positive balances.")
    if mismatch_accounts:
        comments.append(f"{len(mismatch_accounts)} account(s) have entry totals that do not match the trial balance.")

    mismatched_sheets = sum(v['mismatches'] for v in sheet_status.values()) if sheet_status else 0
    if mismatched_sheets > 0:
        comments.append(f"{mismatched_sheets} sheet(s) contain out-of-balance accounts exceeding tolerance {tolerance}.")
    else:
        comments.append("All sheets appear balanced within tolerance limits.")

    if not comments:
        comments.append("No issues detected based on configured checks.")

    # Write bullet comments
    start_row = row_ptr
    for i, comment in enumerate(comments, start=start_row):
        ws_front.cell(i, 1, f"• {comment}")
    row_ptr = start_row + len(comments) + 2

    # Detailed lists with hyperlinks to anchors (internal)
    def set_hyperlink(cell, acc_no):
        acc = str(acc_no)
        if acc in account_anchor:
            sheet_name, anchor_row = account_anchor[acc]
            sheet_ref = f"'{sheet_name}'" if not sheet_name.isalnum() else sheet_name
            cell.hyperlink = f"#{sheet_ref}!A{anchor_row}"
            cell.style = "Hyperlink"

    # 1) Accounts out of balance (TB vs entries) – FIRST
    if mismatch_accounts:
        ws_front.cell(row_ptr, 1, "Accounts out of balance (TB vs entries):").font = Font(bold=True)
        row_ptr += 1

        headers = ["Account", "Name", "TB balance", "Entries sum", "Difference"]
        for col_idx, h in enumerate(headers, start=1):
            cell = ws_front.cell(row_ptr, col_idx, h)
            cell.font = Font(bold=True)
        row_ptr += 1

        for m in mismatch_accounts:
            acc = str(m["No"])
            c = ws_front.cell(row_ptr, 1, acc)
            set_hyperlink(c, acc)
            ws_front.cell(row_ptr, 2, m.get("Name", ""))

            tb_cell = ws_front.cell(row_ptr, 3, m.get("tb_balance", 0.0))
            ent_cell = ws_front.cell(row_ptr, 4, m.get("entries_sum", 0.0))
            diff_cell = ws_front.cell(row_ptr, 5, m.get("difference", 0.0))

            tb_cell.number_format = "#,##0.00"
            ent_cell.number_format = "#,##0.00"
            diff_cell.number_format = "#,##0.00"

            row_ptr += 1

        row_ptr += 1  # spacing

    # 2) Negative balances
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

    # 3) Positive balances
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
  
    # === Documentation checklist (account-based comments) ===
    doc_rules = [
        # Tangible fixed assets
        (
            [
                (142110, 142120),
                (142210, 142220),
                (143110, 143120),
                (144110, 144120),
            ],
            "Add documentation for Depreciation fixed assets",
        ),
        # Long-term receivables
        (
            [
                (234110, 234120),
            ],
            "Add documentation for Long-term receivables",
        ),
        # Trade receivables
        (
            [
                (311000, 311020),
            ],
            "Add documentation for Trade receivables",
        ),
        # Amounts owed by affiliate companies
        (
            [
                (321000, 321100),
            ],
            "Add documentation for Amounts owed by affiliate companies",
        ),
        # Bank account
        (
            [
                (391010, 391070),
                (393005, 393998),
            ],
            "Add documentation for Bank account",
        ),
        # Other Liabilities
        (
            [
                (634010, 634011),
            ],
            "Add documentation for Other I/C Loans",
        ),
        # Trade payables
        (
            [
                (721000, 721001),
            ],
            "Add documentation for Trade payables",
        ),
        # Amounts owed to affiliated companies
        (
            [
                (731000, 731100),
            ],
            "Add documentation for Amounts owed to affiliated companies",
        ),
    ]

    def _in_any_range(acc_int, ranges):
        return any(lo <= acc_int <= hi for lo, hi in ranges)

    # Build list of documentation items based on TB balances
    doc_items = []
    for _, r in trial_balance_df.iterrows():
        acc_str = str(r["No."])
        if not acc_str.isdigit():
            continue

        acc_int = int(acc_str)
        bal = r.get("Balance at Date", 0.0) or 0.0
        if abs(bal) <= tolerance:
            continue  # no balance -> no documentation needed

        for ranges, message in doc_rules:
            if _in_any_range(acc_int, ranges):
                doc_items.append(
                    {
                        "No": acc_str,
                        "Name": r.get("Name", ""),
                        "Message": message,
                    }
                )
                break  # stop at first matching rule

    if doc_items:
        row_ptr += 2
        ws_front.cell(row_ptr, 1, "Documentation checklist:").font = Font(
            bold=True, underline="single"
        )
        row_ptr += 1

        # Header row
        headers = ["Account", "Name", "Comment", "Status"]
        for col_idx, h in enumerate(headers, start=1):
            cell = ws_front.cell(row_ptr, col_idx, h)
            cell.font = Font(bold=True)
        row_ptr += 1

        # Dropdown for "Done"
        dv = DataValidation(type="list", formula1='"Done"', allow_blank=True)
        ws_front.add_data_validation(dv)

        for item in doc_items:
            acc = item["No"]
            name = item["Name"]
            msg = item["Message"]

            # Account with hyperlink
            acc_cell = ws_front.cell(row_ptr, 1, acc)
            set_hyperlink(acc_cell, acc)

            # Name + comment
            ws_front.cell(row_ptr, 2, name)
            ws_front.cell(row_ptr, 3, msg)

            # Status dropdown
            status_cell = ws_front.cell(row_ptr, 4)
            dv.add(status_cell)

            # Conditional formatting: if Status == "Done", make the row green
            formula = f'$D{row_ptr}="Done"'
            rule = FormulaRule(formula=[formula], fill=green_fill)
            ws_front.conditional_formatting.add(f"A{row_ptr}:D{row_ptr}", rule)

            row_ptr += 1

    # Footer metadata
    row_ptr += 2
    ws_front.cell(row_ptr, 1, f"Generated on: {pd.Timestamp.now():%Y-%m-%d %H:%M}")
    ws_front.cell(row_ptr + 1, 1, f"Tolerance used: {tolerance}")
    ws_front.cell(row_ptr + 2, 1, f"Source files expected in: {STATIC_DIR}")

    # === FORMAT numbers & dates across all sheets ===
    amount_fmt = "#,##0.00"
    date_fmt = "yyyy-mm-dd"

    for ws in wb.worksheets:
        # Hide gridlines
        ws.sheet_view.showGridLines = False

        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                # Dates: if cell contains a date or pandas Timestamp
                if isinstance(cell.value, (pd.Timestamp,)) or hasattr(cell.value, "year") and not isinstance(cell.value, (int, float, str)):
                    # apply date format
                    try:
                        pd.to_datetime(cell.value)
                        cell.number_format = date_fmt
                    except Exception:
                        pass
                # Numeric formatting: numeric types -> amount format
                if isinstance(cell.value, (int, float)):
                    cell.number_format = amount_fmt

    # Auto-fit columns
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max((len(str(c.value)) if c.value else 0 for c in col), default=0)
            ws.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = max_len + 2

    # Save workbook to bytes buffer
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# === PUBLIC: generate_reconciliation_file ===
def generate_reconciliation_file(trial_balance_file, entries_file, icp_code, mapping_path=None, plc_path=None, tolerance=TOLERANCE):
    """
    Main entrypoint for Streamlit app.
    Inputs trial_balance_file and entries_file may be:
      - file-like objects (UploadedFile)
      - pathlib.Path or str paths

    Returns:
      - BytesIO (seeked to 0) containing the generated workbook .xlsx
    """
    # Load mapping from static (app-internal) unless explicit path provided
    mapping_path = Path(mapping_path) if mapping_path else DEFAULT_MAPPING
    plc_path = Path(plc_path) if plc_path else DEFAULT_PLC

    # Read user inputs (pandas handles file-like objects)
    trial_balance = pd.read_excel(trial_balance_file)
    entries = pd.read_excel(entries_file)

    # Load mapping tables
    acct_to_code, code_to_meta, map_dir = load_mapping(mapping_path)

    # Apply mapping to trial balance
    trial_balance = apply_mapping(trial_balance, acct_to_code, code_to_meta)

    # Normalise entries column names
    entries.rename(columns=lambda c: "Amount (LCY)" if str(c).strip().lower() in ["amount", "amount (lcy)"] else c, inplace=True)
    entries.rename(columns=lambda c: "ICP CODE" if str(c).strip().lower() == "icp code" else c, inplace=True)

    # Cast types & normalize
    trial_balance["Balance at Date"] = trial_balance["Balance at Date"].apply(to_float)
    entries["G/L Account No."] = entries["G/L Account No."].apply(normalize_account)
    entries["Amount (LCY)"] = entries["Amount (LCY)"].apply(to_float)
    # Remove timestamps: keep date only (pandas -> datetime.date)
    entries["Posting Date"] = pd.to_datetime(entries["Posting Date"], errors="coerce").dt.date

    # Build workbook
    wb, sheet_status, account_anchor, mismatch_accounts = build_workbook(
        trial_balance, entries, map_dir, acct_to_code, code_to_meta, icp_code, tolerance=tolerance
    )

    # Finalize & get bytes
    bio = finalize_workbook_to_bytes(
        wb,
        sheet_status,
        account_anchor,
        trial_balance,
        entries,
        icp_code,
        plc_path=plc_path,
        tolerance=tolerance,
        mismatch_accounts=mismatch_accounts,
    )
    return bio
