"""
SAP GUI Automation + Financial Analysis
----------------------------------------
Requires:
  pip install pywin32 openpyxl pandas

SAP GUI Scripting must be enabled in SAP Logon Options.
"""

import sys
import time
import os
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# USER CONFIGURATION — adjust before running
# ─────────────────────────────────────────────────────────────────────────────

COMPANY_CODE        = "1000"          # SAP company code for FEBAN
CASHPOOL_ACCOUNT    = "11300001"      # G/L account number for FBL3N check
FISCAL_YEAR         = str(datetime.now().year)
POSTING_DATE_FROM   = "01.01." + FISCAL_YEAR
POSTING_DATE_TO     = datetime.now().strftime("%d.%m.%Y")

# Local path where SAP will save the exported file
EXPORT_DIR          = os.path.join(os.path.expanduser("~"), "Desktop")
FEBAN_FILE          = os.path.join(EXPORT_DIR, "FEBAN_export.xlsx")
FBL3N_FILE          = os.path.join(EXPORT_DIR, "FBL3N_export.xlsx")

# SAP document types that indicate a payment run
PAYMENT_RUN_DOC_TYPES = {"ZP", "ZA", "ZB", "ZR", "ZV"}
PAYMENT_RUN_KEYWORDS  = {"payment", "zahlungslauf", "pay run", "payrun", "pmnt"}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _col(label: str) -> str:
    return label.strip().lower()


def _wait(session, ms: int = 500):
    session.utils.asyncWaitFinished()
    time.sleep(ms / 1000)


def _sap_connect():
    """Return an active SAP GUI session (first available logon)."""
    try:
        import win32com.client as win32
    except ImportError:
        sys.exit("pywin32 not installed. Run: pip install pywin32")

    try:
        rot = win32.GetObject("SAPGUI")
    except Exception:
        sys.exit(
            "SAP GUI not running or scripting disabled.\n"
            "Enable via: SAP Logon → Options → Accessibility & Scripting → Scripting."
        )

    app = rot.GetScriptingEngine
    if app.Connections.Count == 0:
        sys.exit("No active SAP connection found. Please log on to SAP first.")

    conn    = app.Connections(0)
    session = conn.Sessions(0)
    return session


# ─────────────────────────────────────────────────────────────────────────────
# SAP EXPORT HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _export_list_to_excel(session, filepath: str):
    """Use SAP's built-in 'Spreadsheet' export from any ALV list."""
    try:
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()  # System → List → Save → Local file
    except Exception:
        # Fallback: try toolbar export button
        try:
            session.findById("wnd[0]/tbar[1]/btn[43]").press()
        except Exception:
            raise RuntimeError("Cannot locate export menu entry. Check SAP layout.")

    _wait(session, 800)

    # 'Save list in file' dialog
    try:
        dlg = session.findById("wnd[1]")
        # Choose 'Spreadsheet' radio
        dlg.findById("usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        dlg.findById("usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()  # Continue
    except Exception:
        pass

    _wait(session, 600)

    # Enter filename
    try:
        fname_field = session.findById("wnd[1]/usr/ctxtDY_FILENAME")
        fname_field.text = filepath
        session.findById("wnd[1]/tbar[0]/btn[0]").press()  # Replace/OK
    except Exception:
        raise RuntimeError("Export filename dialog not found.")

    _wait(session, 1200)

    # Confirm overwrite if asked
    try:
        session.findById("wnd[2]/tbar[0]/btn[11]").press()
        _wait(session, 800)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# 1. FEBAN — Bank Statement Monitor
# ─────────────────────────────────────────────────────────────────────────────

def run_feban(session) -> str:
    """Navigate to FEBAN, enter parameters, execute, export to Excel."""
    print("\n[FEBAN] Starting transaction...")
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nFEBAN"
    session.findById("wnd[0]").sendVKey(0)  # Enter
    _wait(session, 1000)

    try:
        # Company code
        session.findById("wnd[0]/usr/ctxtRF05B-BUKRS").text = COMPANY_CODE
        # Posting date range
        session.findById("wnd[0]/usr/ctxtRF05B-BUDAT_FROM").text = POSTING_DATE_FROM
        session.findById("wnd[0]/usr/ctxtRF05B-BUDAT_TO").text   = POSTING_DATE_TO
    except Exception as e:
        print(f"  [WARN] Could not set some FEBAN fields: {e}")

    # Execute (F8)
    session.findById("wnd[0]").sendVKey(8)
    _wait(session, 2000)

    print(f"[FEBAN] Exporting to {FEBAN_FILE} ...")
    _export_list_to_excel(session, FEBAN_FILE)

    if not os.path.exists(FEBAN_FILE):
        sys.exit(f"[FEBAN] Export file not created: {FEBAN_FILE}")

    print(f"[FEBAN] Export OK → {FEBAN_FILE}")
    return FEBAN_FILE


# ─────────────────────────────────────────────────────────────────────────────
# 2. FBL3N — G/L Line Items (cashpool account)
# ─────────────────────────────────────────────────────────────────────────────

def run_fbl3n(session) -> str:
    """Navigate to FBL3N for the cashpool account, export to Excel."""
    print(f"\n[FBL3N] Starting transaction for account {CASHPOOL_ACCOUNT}...")
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL3N"
    session.findById("wnd[0]").sendVKey(0)
    _wait(session, 1000)

    try:
        session.findById("wnd[0]/usr/ctxtRF05B-AGKON").text = CASHPOOL_ACCOUNT
        session.findById("wnd[0]/usr/ctxtRF05B-BUKRS").text = COMPANY_CODE
        session.findById("wnd[0]/usr/radRF05B-XOPVW").select()   # Open items
        # Posting date
        session.findById("wnd[0]/usr/ctxtRF05B-BUDAT[0]").text = POSTING_DATE_FROM
        session.findById("wnd[0]/usr/ctxtRF05B-BUDAT[1]").text = POSTING_DATE_TO
    except Exception as e:
        print(f"  [WARN] Could not set some FBL3N fields: {e}")

    session.findById("wnd[0]").sendVKey(8)
    _wait(session, 2000)

    print(f"[FBL3N] Exporting to {FBL3N_FILE} ...")
    _export_list_to_excel(session, FBL3N_FILE)

    if not os.path.exists(FBL3N_FILE):
        sys.exit(f"[FBL3N] Export file not created: {FBL3N_FILE}")

    print(f"[FBL3N] Export OK → {FBL3N_FILE}")
    return FBL3N_FILE


# ─────────────────────────────────────────────────────────────────────────────
# 3. Load Excel exports into DataFrames
# ─────────────────────────────────────────────────────────────────────────────

def _normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [c.strip() for c in df.columns]
    return df


def _find_amount_column(df: pd.DataFrame) -> str:
    """Return the first column name that looks like an amount/value column."""
    candidates = ["amount", "betrag", "amount in lc", "betrag in hw", "wert", "value"]
    for col in df.columns:
        if col.strip().lower() in candidates:
            return col
    # Fallback: first numeric column
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            return col
    raise KeyError("No amount column detected. Check column names: " + str(list(df.columns)))


def _find_doctype_column(df: pd.DataFrame) -> str | None:
    candidates = ["doc. type", "document type", "belegtyp", "doc type", "blart"]
    for col in df.columns:
        if col.strip().lower() in candidates:
            return col
    return None


def _find_text_column(df: pd.DataFrame) -> str | None:
    candidates = ["text", "posting text", "buchungstext", "reference", "referenz", "description"]
    for col in df.columns:
        if col.strip().lower() in candidates:
            return col
    return None


def load_feban(filepath: str) -> pd.DataFrame:
    df = pd.read_excel(filepath, dtype=str)
    df = _normalise_columns(df)
    amt_col = _find_amount_column(df)
    df[amt_col] = pd.to_numeric(df[amt_col].str.replace(",", ".").str.replace(" ", ""), errors="coerce")
    df.dropna(subset=[amt_col], inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


def load_fbl3n(filepath: str) -> pd.DataFrame:
    df = pd.read_excel(filepath, dtype=str)
    df = _normalise_columns(df)
    amt_col = _find_amount_column(df)
    df[amt_col] = pd.to_numeric(df[amt_col].str.replace(",", ".").str.replace(" ", ""), errors="coerce")
    df.dropna(subset=[amt_col], inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


# ─────────────────────────────────────────────────────────────────────────────
# 4. Analysis
# ─────────────────────────────────────────────────────────────────────────────

def detect_clearable_pairs(df: pd.DataFrame, amt_col: str) -> pd.DataFrame:
    """
    Mark rows where the absolute amount appears at least once positive
    and at least once negative → potential clearing pair.
    Returns df with two new columns: 'clearable_pair' (bool) and 'pair_amount'.
    """
    amounts = df[amt_col]
    abs_amounts = amounts.abs()

    # Group by absolute value; a clearable pair needs both signs present
    sign_check = df.groupby(abs_amounts)[amt_col].apply(
        lambda s: (s > 0).any() and (s < 0).any()
    )
    clearable_abs = sign_check[sign_check].index

    df = df.copy()
    df["clearable_pair"] = abs_amounts.isin(clearable_abs)
    df["pair_amount"]    = df[amt_col].where(df["clearable_pair"])
    return df


def detect_payment_runs(df: pd.DataFrame) -> pd.DataFrame:
    """
    Flag rows as payment_run based on document type or text fields.
    Adds column 'is_payment_run' (bool) and 'payment_run_hint' (str).
    """
    df = df.copy()
    df["is_payment_run"]    = False
    df["payment_run_hint"]  = ""

    dt_col   = _find_doctype_column(df)
    txt_col  = _find_text_column(df)

    if dt_col:
        mask = df[dt_col].str.strip().str.upper().isin(PAYMENT_RUN_DOC_TYPES)
        df.loc[mask, "is_payment_run"]   = True
        df.loc[mask, "payment_run_hint"] = "Document type: " + df.loc[mask, dt_col].str.strip()

    if txt_col:
        pattern = "|".join(PAYMENT_RUN_KEYWORDS)
        mask2 = df[txt_col].str.lower().str.contains(pattern, na=False)
        mask2 &= ~df["is_payment_run"]   # avoid overwriting already-flagged
        df.loc[mask2, "is_payment_run"]   = True
        df.loc[mask2, "payment_run_hint"] = "Keyword in text: " + df.loc[mask2, txt_col]

    return df


def check_cashpool_match(feban_df: pd.DataFrame, fbl3n_df: pd.DataFrame,
                          feban_amt_col: str, fbl3n_amt_col: str) -> pd.DataFrame:
    """
    For every amount in FEBAN, check whether the same absolute value
    appears in the FBL3N cashpool export.
    Adds column 'cashpool_match' (bool).
    """
    cashpool_abs = fbl3n_df[fbl3n_amt_col].abs().unique()
    feban_df = feban_df.copy()
    feban_df["cashpool_match"] = feban_df[feban_amt_col].abs().isin(cashpool_abs)
    return feban_df


# ─────────────────────────────────────────────────────────────────────────────
# 5. Console report
# ─────────────────────────────────────────────────────────────────────────────

def _sep(char="─", width=80):
    print(char * width)


def print_console_report(df: pd.DataFrame, amt_col: str):
    total       = len(df)
    clearable   = df["clearable_pair"].sum()
    pay_runs    = df["is_payment_run"].sum()
    cashpool    = df["cashpool_match"].sum() if "cashpool_match" in df.columns else 0

    _sep("═")
    print(" SAP FEBAN — FINANCIAL ANALYSIS REPORT")
    print(f" Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f" Company code: {COMPANY_CODE}  |  Period: {POSTING_DATE_FROM} – {POSTING_DATE_TO}")
    _sep("═")

    print(f"\n{'SUMMARY':}")
    _sep()
    print(f"  Total line items       : {total:>8}")
    print(f"  Clearable pairs        : {clearable:>8}  "
          f"({clearable/total*100:.1f}% of total)" if total else "")
    print(f"  Payment run items      : {pay_runs:>8}")
    print(f"  Cashpool matches       : {cashpool:>8}")
    _sep()

    # Clearable pairs detail
    if clearable:
        print("\n CLEARABLE PAIRS (same amount with opposing signs):")
        _sep()
        pair_df = df[df["clearable_pair"]].copy()
        pair_df = pair_df.sort_values(by=amt_col)
        cols_to_show = [amt_col, "clearable_pair"]
        if _find_doctype_column(df):
            cols_to_show.insert(0, _find_doctype_column(df))
        if _find_text_column(df):
            cols_to_show.insert(0, _find_text_column(df))
        print(pair_df[cols_to_show].to_string(index=True, max_rows=40))
        if len(pair_df) > 40:
            print(f"  ... and {len(pair_df)-40} more rows (see Excel sheet)")
        _sep()

    # Payment runs
    if pay_runs:
        print("\n PAYMENT RUN ITEMS:")
        _sep()
        pr_df = df[df["is_payment_run"]][
            [c for c in [_find_doctype_column(df), _find_text_column(df),
                          amt_col, "payment_run_hint"] if c]
        ]
        print(pr_df.to_string(index=True, max_rows=30))
        if len(pr_df) > 30:
            print(f"  ... and {len(pr_df)-30} more rows (see Excel sheet)")
        _sep()

    # Cashpool matches
    if cashpool and "cashpool_match" in df.columns:
        print(f"\n CASHPOOL MATCHES (account {CASHPOOL_ACCOUNT}):")
        _sep()
        cm_df = df[df["cashpool_match"]][[amt_col, "cashpool_match"]]
        print(cm_df.to_string(index=True, max_rows=30))
        _sep()

    print("\n[OK] Console report complete.\n")


# ─────────────────────────────────────────────────────────────────────────────
# 6. Write analysis sheet back into FEBAN Excel
# ─────────────────────────────────────────────────────────────────────────────

_FILL_GREEN  = PatternFill("solid", fgColor="C6EFCE")
_FILL_ORANGE = PatternFill("solid", fgColor="FFEB9C")
_FILL_BLUE   = PatternFill("solid", fgColor="BDD7EE")
_FILL_HEADER = PatternFill("solid", fgColor="1F4E79")
_FONT_HEADER = Font(bold=True, color="FFFFFF", size=10)
_THIN        = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)


def _write_header(ws, row, values):
    for col_idx, val in enumerate(values, start=1):
        cell = ws.cell(row=row, column=col_idx, value=val)
        cell.fill   = _FILL_HEADER
        cell.font   = _FONT_HEADER
        cell.border = _THIN
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def write_analysis_sheet(df: pd.DataFrame, amt_col: str, filepath: str):
    """Append an 'Analysis' worksheet to the existing Excel workbook."""
    wb = load_workbook(filepath)

    sheet_name = "Analysis"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    # ── Title ──────────────────────────────────────────────────────────────
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value     = f"FEBAN Financial Analysis — {COMPANY_CODE} — {datetime.now().strftime('%Y-%m-%d')}"
    title_cell.font      = Font(bold=True, size=13, color="1F4E79")
    title_cell.alignment = Alignment(horizontal="center")

    # ── Column setup ───────────────────────────────────────────────────────
    export_cols = list(df.columns)
    _write_header(ws, row=3, values=export_cols)

    # Colour map per row
    for r_idx, (_, row) in enumerate(df.iterrows(), start=4):
        for c_idx, col in enumerate(export_cols, start=1):
            cell       = ws.cell(row=r_idx, column=c_idx, value=row[col])
            cell.border = _THIN
            cell.alignment = Alignment(wrap_text=False)

        # Apply row highlighting
        if "is_payment_run" in df.columns and row.get("is_payment_run"):
            fill = _FILL_ORANGE
        elif "clearable_pair" in df.columns and row.get("clearable_pair"):
            fill = _FILL_GREEN
        elif "cashpool_match" in df.columns and row.get("cashpool_match"):
            fill = _FILL_BLUE
        else:
            fill = None

        if fill:
            for c_idx in range(1, len(export_cols) + 1):
                ws.cell(row=r_idx, column=c_idx).fill = fill

    # Auto-fit columns (approximate)
    for col_idx, col_name in enumerate(export_cols, start=1):
        max_len = max(
            len(str(col_name)),
            df[col_name].astype(str).str.len().max() if len(df) else 0,
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 40)

    # ── Legend ─────────────────────────────────────────────────────────────
    legend_row = len(df) + 6
    ws.cell(row=legend_row,     column=1, value="LEGEND").font = Font(bold=True)
    legend_items = [
        (_FILL_GREEN,  "Clearable pair — same amount with opposing sign"),
        (_FILL_ORANGE, "Payment run item — document type or keyword match"),
        (_FILL_BLUE,   f"Cashpool match — found in FBL3N account {CASHPOOL_ACCOUNT}"),
    ]
    for i, (fill, label) in enumerate(legend_items, start=1):
        ws.cell(row=legend_row + i, column=1).fill  = fill
        ws.cell(row=legend_row + i, column=1).border = _THIN
        ws.cell(row=legend_row + i, column=2, value=label)

    wb.save(filepath)
    print(f"[Excel] Analysis sheet written → {filepath}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  SAP FEBAN Automation + Financial Analysis")
    print("=" * 60)
    print(f"  Company code  : {COMPANY_CODE}")
    print(f"  Cashpool acct : {CASHPOOL_ACCOUNT}")
    print(f"  Period        : {POSTING_DATE_FROM} – {POSTING_DATE_TO}")
    print(f"  Export dir    : {EXPORT_DIR}")
    print("=" * 60)

    # ── Step 1: SAP automation ─────────────────────────────────────────────
    session = _sap_connect()
    feban_path = run_feban(session)
    fbl3n_path = run_fbl3n(session)

    # ── Step 2: Load data ──────────────────────────────────────────────────
    print("\n[Analysis] Loading exported files...")
    feban_df  = load_feban(feban_path)
    fbl3n_df  = load_fbl3n(fbl3n_path)

    feban_amt = _find_amount_column(feban_df)
    fbl3n_amt = _find_amount_column(fbl3n_df)

    print(f"  FEBAN rows  : {len(feban_df)}  (amount column: '{feban_amt}')")
    print(f"  FBL3N rows  : {len(fbl3n_df)}  (amount column: '{fbl3n_amt}')")

    # ── Step 3: Analyses ───────────────────────────────────────────────────
    print("\n[Analysis] Detecting clearable pairs...")
    feban_df = detect_clearable_pairs(feban_df, feban_amt)

    print("[Analysis] Detecting payment runs...")
    feban_df = detect_payment_runs(feban_df)

    print("[Analysis] Cross-checking cashpool account (FBL3N)...")
    feban_df = check_cashpool_match(feban_df, fbl3n_df, feban_amt, fbl3n_amt)

    # ── Step 4: Console output ─────────────────────────────────────────────
    print_console_report(feban_df, feban_amt)

    # ── Step 5: Write analysis sheet ───────────────────────────────────────
    write_analysis_sheet(feban_df, feban_amt, feban_path)

    print("\n[DONE] All tasks completed successfully.")
    print(f"       Result file: {feban_path}")


if __name__ == "__main__":
    main()
