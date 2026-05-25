import win32com.client
from win32com.client import dynamic
import time
import os
import re
from datetime import datetime
from collections import defaultdict

try:
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = dynamic.Dispatch(SapGuiAuto.GetScriptingEngine)
except Exception:
    print("BLAD: Nie mozna polaczyc sie z SAP GUI.")
    print("Upewnij sie ze SAP GUI jest otwarty.")
    input("Nacisnij Enter aby zamknac...")
    raise SystemExit(1)

try:
    connection = application.Children(0)
except Exception:
    print("BLAD: Brak aktywnego polaczenia z SAP.")
    print("Zaloguj sie do SAP przed uruchomieniem skryptu.")
    input("Nacisnij Enter aby zamknac...")
    raise SystemExit(1)

try:
    session = connection.Children(0)
except Exception:
    print("BLAD: Brak aktywnej sesji SAP.")
    print("Zaloguj sie do SAP przed uruchomieniem skryptu.")
    input("Nacisnij Enter aby zamknac...")
    raise SystemExit(1)

SAVE_DIR = os.path.join(os.path.expanduser("~"), "Documents")
os.makedirs(SAVE_DIR, exist_ok=True)

company_code = input("Wpisz numer spolki: ").strip()
print(f"Polaczono z SAP - spolka {company_code}")

print("Otwieram FEBAN...")
session.findById("wnd[0]/tbar[0]/okcd").text = "/nFEBAN"
session.findById("wnd[0]").sendVKey(0)
time.sleep(2)

print("Wpisuje Company Code 2052...")
session.findById("wnd[1]/usr/ctxtSL_BUKRS-LOW").text = company_code

print("Wykonuje F8...")
session.findById("wnd[1]").sendVKey(8)
time.sleep(3)

shell = session.findById("wnd[0]/shellcont/shell")

print("Laduje wszystkie wiersze FEBAN...")
prev = -1
while True:
    current = shell.RowCount
    if current == prev:
        break
    prev = current
    try:
        shell.firstVisibleRow = current
    except:
        break
    time.sleep(0.3)
try:
    shell.firstVisibleRow = 0
except:
    pass
time.sleep(0.5)

row_count = shell.RowCount
col_count = shell.ColumnCount
print(f"Znaleziono {row_count} wierszy, {col_count} kolumn")

col_ids = []
try:
    for col in shell.ColumnOrder:
        col_ids.append(str(col))
except:
    for i in range(col_count):
        try:
            col_ids.append(shell.ColumnOrder(i))
        except:
            col_ids.append(f"Col{i}")


def parse_sap_amount(val):
    if not val or str(val).strip() == "":
        return 0
    s = str(val).strip()
    negative = s.endswith("-")
    if negative:
        s = s[:-1]
    s = s.replace(".", "").replace(",", ".")
    try:
        return -float(s) if negative else float(s)
    except:
        return val


NOTE_PATH = "wnd[0]/usr/ssubAREA_N2P:FEB_BSPROC_FE:0113/cntlAREA_N2P/shellcont/shell"


def note_to_value(note):
    if not note:
        return note
    s = str(note).strip()
    if re.match(r'^\d[\d.,]*$', s):
        try:
            f = float(s.replace(".", "").replace(",", "."))
            return int(f) if f == int(f) else f
        except:
            pass
    return note


def dismiss_popup():
    try:
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
        time.sleep(0.4)
        return True
    except:
        return False


def find_working_note_path():
    try:
        session.findById(NOTE_PATH)
        print(f"  Znaleziono Note to Payee pod: {NOTE_PATH}")
        return NOTE_PATH
    except:
        print("  UWAGA: Nie znaleziono pola Note to Payee - kolumna bedzie pusta")
        return None


kwbtr_idx = col_ids.index("KWBTR") if "KWBTR" in col_ids else None

# Faza 1: wczytaj dane z gridu FEBAN
print("Czytam dane z gridu FEBAN...")
grid_data = []
for row in range(row_count):
    try:
        shell.setCurrentCell(row, col_ids[0])
        if row % 10 == 0:
            time.sleep(0.3)
    except:
        pass
    row_data = {}
    for col_id in col_ids:
        try:
            val = shell.GetCellValue(row, col_id)
        except:
            val = ""
        if col_id == "KWBTR":
            val = parse_sap_amount(val)
        row_data[col_id] = val
    grid_data.append(row_data)
    print(f"  Grid wiersz {row+1}/{row_count}")

# Faza 2: klikaj kazdy wiersz i czytaj Note to Payee
print("Czytam Note to Payee (klikam kazdy wiersz)...")
shell.setCurrentCell(0, col_ids[0])
time.sleep(1)
working_path = find_working_note_path()

notes = []
for row in range(row_count):
    try:
        shell.setCurrentCell(row, col_ids[0])
        shell.doubleClickCurrentCell()
        time.sleep(0.5)
        dismiss_popup()
        time.sleep(0.6)
        if working_path:
            try:
                note = session.findById(working_path).text
            except:
                note = ""
        else:
            note = ""
    except Exception as e:
        note = ""
        print(f"  Wiersz {row+1}: blad - {e}")
    notes.append(note)
    preview = note[:60].replace("\n", " ") if note else "(brak)"
    print(f"  Note wiersz {row+1}/{row_count}: {preview}")

# === EXCEL - otwieramy i trzymamy otwarty az do konca ===
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
SAVE_PATH = os.path.join(SAVE_DIR, f"feban_raport_{company_code}_{timestamp}.xlsx")

print("Tworze plik Excel...")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Add()

ws = wb.Sheets(1)
ws.Name = "FEBAN"

all_cols = col_ids + ["Note to Payee"]
for i, col_id in enumerate(all_cols):
    ws.Cells(1, i + 1).Value = col_id

for row_idx, row_data in enumerate(grid_data):
    for col_idx, col_id in enumerate(col_ids):
        ws.Cells(row_idx + 2, col_idx + 1).Value = row_data[col_id]
    ws.Cells(row_idx + 2, len(col_ids) + 1).Value = note_to_value(notes[row_idx])

wb.SaveAs(SAVE_PATH)

try:
    if kwbtr_idx is not None:
        ws.Columns(kwbtr_idx + 1).NumberFormat = "#,##0.00"
    wb.Save()
except Exception as e:
    print(f"  (Formatowanie KWBTR pominiete: {e})")

try:
    note_col = len(col_ids) + 1
    last_row = row_count + 1
    ws.Range(ws.Cells(1, note_col), ws.Cells(last_row, note_col)).HorizontalAlignment = -4131
    ws.Range(ws.Cells(1, note_col), ws.Cells(last_row, note_col)).NumberFormat = "#,##0"
    wb.Save()
except Exception as e:
    print(f"  (Formatowanie Note to Payee pominiete: {e})")

LIGHT_GREEN = 144 + 238 * 256 + 144 * 65536  # RGB(144, 238, 144)
colored_rows = set()

# Pary +/- KWBTR w FEBAN
feban_amounts = defaultdict(list)  # amt -> [excel_rows]
if kwbtr_idx is not None:
    print("Szukam par +/- w kolumnie KWBTR (FEBAN)...")
    feban_positives = defaultdict(list)
    feban_negatives = defaultdict(list)
    for row in range(len(grid_data)):
        excel_row = row + 2
        val = ws.Cells(excel_row, kwbtr_idx + 1).Value
        if val is None:
            continue
        try:
            amt = round(float(val), 2)
        except:
            continue
        feban_amounts[amt].append(excel_row)
        if amt > 0:
            feban_positives[amt].append(excel_row)
        elif amt < 0:
            feban_negatives[round(abs(amt), 2)].append(excel_row)

    matched = set(feban_positives.keys()) & set(feban_negatives.keys())
    rows_to_color = []
    for amt in matched:
        rows_to_color.extend(feban_positives[amt])
        rows_to_color.extend(feban_negatives[amt])
    for excel_row in rows_to_color:
        ws.Rows(excel_row).Interior.Color = LIGHT_GREEN
        colored_rows.add(excel_row)
    wb.Save()
    print(f"Znaleziono {len(matched)} par, pokolorowano {len(rows_to_color)} wierszy")

# Payment run rozowy
note_col = len(col_ids) + 1
LIGHT_PINK = 255 + 220 * 256 + 220 * 65536
pink_count = 0
for row in range(len(grid_data)):
    excel_row = row + 2
    val = ws.Cells(excel_row, note_col).Value
    if val is not None and isinstance(val, (int, float)):
        ws.Rows(excel_row).Interior.Color = LIGHT_PINK
        pink_count += 1
wb.Save()
print(f"Pokolorowano {pink_count} wierszy na rozowy (payment run)")

# On-account posting zolty
SEARCH_TEXT = "@5D\\QPosting in Subledger Accounting Made as On Account Posting@"
LIGHT_YELLOW = 255 + 255 * 256 + 153 * 65536
yellow_count = 0
for row in range(len(grid_data)):
    excel_row = row + 2
    val = ws.Cells(excel_row, 2).Value
    if val and SEARCH_TEXT in str(val):
        ws.Rows(excel_row).Interior.Color = LIGHT_YELLOW
        colored_rows.add(excel_row)
        yellow_count += 1
wb.Save()
print(f"Pokolorowano {yellow_count} wierszy na zolty")

# CASHPOOLING
CASHPOOLING_KEYWORDS = ["sweep credit", "cash", "nazareth", "konsolidacja salda"]
cashpool_count = 0
for row in range(len(grid_data)):
    excel_row = row + 2
    note_val = ws.Cells(excel_row, note_col).Value
    if note_val:
        note_lower = str(note_val).lower()
        if any(kw in note_lower for kw in CASHPOOLING_KEYWORDS):
            cell = ws.Cells(excel_row, 24)
            cell.Value = "CASHPOOLING"
            cell.Font.Bold = True
            cashpool_count += 1
wb.Save()
print(f"Oznaczono {cashpool_count} wierszy jako CASHPOOLING")

# Return from vendor - dodatnia kwota bez zadnego koloru
retur_count = 0
if kwbtr_idx is not None:
    print("Szukam dodatnich kwot bez kategorii (Return from vendor?)...")
    for row in range(len(grid_data)):
        excel_row = row + 2
        if excel_row in colored_rows:
            continue
        val = ws.Cells(excel_row, kwbtr_idx + 1).Value
        if val is None:
            continue
        try:
            amt = round(float(val), 2)
        except:
            continue
        if amt > 0:
            cell = ws.Cells(excel_row, 24)
            existing = cell.Value
            if existing == "CASHPOOLING":
                cell.Value = "CASHPOOLING/Return from vendor?"
            else:
                cell.Value = "Return from vendor?"
            cell.Font.Bold = True
            retur_count += 1
wb.Save()
print(f"Oznaczono {retur_count} wierszy jako Return from vendor?")

# === FBL3N - nowe okno SAP ===
print("\nOtwieram nowe okno SAP dla FBL3N...")
session.createSession()
time.sleep(3)
session2 = connection.Children(1)

today_str = datetime.now().strftime("%d.%m.%Y")
print("Nawiguje do FBL3N...")
session2.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL3N"
session2.findById("wnd[0]").sendVKey(0)
time.sleep(2)

print(f"Wpisuje parametry FBL3N (konto 10441000, spolka {company_code}, data {today_str})...")
try:
    session2.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").text = "10441000"
except Exception as e:
    print(f"  UWAGA: G/L Account: {e}")
try:
    session2.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = company_code
except Exception as e:
    print(f"  UWAGA: Company Code: {e}")
try:
    session2.findById("wnd[0]/usr/radX_OPSEL").select()
    print("  Zaznaczono Open items")
except Exception as e:
    print(f"  UWAGA: Open items: {e}")
try:
    session2.findById("wnd[0]/usr/ctxtPA_STIDA").text = today_str
    print(f"  Data: {today_str}")
except Exception as e:
    print(f"  UWAGA: Data: {e}")
try:
    session2.findById("wnd[0]/usr/ctxtPA_VARI").text = "FEBAN2052MR"
    print("  Layout: FEBAN2052MR")
except Exception as e:
    print(f"  UWAGA: Layout: {e}")

print("Wykonuje F8 w FBL3N...")
session2.findById("wnd[0]").sendVKey(8)
time.sleep(4)

# Eksport: List -> Export -> Spreadsheet
FBL3N_TEMP_NAME   = f"fbl3n_temp_{timestamp}"
FBL3N_EXPORT_DIR  = SAVE_DIR + "\\"
FBL3N_EXPORT_XLSX = os.path.join(SAVE_DIR, FBL3N_TEMP_NAME + ".xlsx")
FBL3N_EXPORT_XLS  = os.path.join(SAVE_DIR, FBL3N_TEMP_NAME + ".xls")

fbl3n_data        = []
fbl3n_col_ids     = []
fbl3n_amounts     = defaultdict(list)
FBL3N_AMT_COL     = None
fbl3n_amt_col_idx = None
fbl3n_row_count   = 0

print("Eksportuje FBL3N (List -> Export -> Spreadsheet)...")
try:
    session2.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    time.sleep(2)

    # Krok 1: dialog wyboru formatu - XLSX juz zaznaczony, klikamy OK
    try:
        fmt_dlg = session2.findById("wnd[1]")
        print(f"  Dialog formatu: '{fmt_dlg.Text}' - potwierdzam...")
        fmt_dlg.sendVKey(0)
        time.sleep(3)
    except Exception as e:
        print(f"  (Brak dialogu formatu: {e})")

    # Krok 2: dialog zapisu pliku
    export_handled = False
    for save_wnd in ["wnd[1]", "wnd[2]"]:
        try:
            session2.findById(f"{save_wnd}/usr/ctxtDY_FILENAME")
            print(f"  Dialog zapisu w {save_wnd} - wpisuje sciezke...")
            try:
                session2.findById(f"{save_wnd}/usr/ctxtDY_PATH").text = FBL3N_EXPORT_DIR
            except:
                pass
            session2.findById(f"{save_wnd}/usr/ctxtDY_FILENAME").text = FBL3N_TEMP_NAME + ".xlsx"
            try:
                session2.findById(f"{save_wnd}/tbar[0]/btn[0]").press()
            except:
                session2.findById(save_wnd).sendVKey(0)
            time.sleep(3)
            export_handled = True
            print(f"  Zapisano: {FBL3N_EXPORT_XLSX}")
            break
        except:
            continue

    if not export_handled:
        print("  (Brak dialogu zapisu - SAP mogl otworzyc Excel bezposrednio)")

except Exception as e:
    print(f"  BLAD eksportu FBL3N: {e}")

# Wczytaj wyeksportowany plik
fbl3n_export_path = None
for candidate_path in [FBL3N_EXPORT_XLSX, FBL3N_EXPORT_XLS]:
    if os.path.exists(candidate_path):
        fbl3n_export_path = candidate_path
        break

if fbl3n_export_path:
    print(f"Wczytuje dane FBL3N z: {fbl3n_export_path}")
    try:
        wb_rd = excel.Workbooks.Open(fbl3n_export_path)
        ws_rd = wb_rd.Sheets(1)
        nr_rows = ws_rd.UsedRange.Rows.Count
        nr_cols = ws_rd.UsedRange.Columns.Count

        fbl3n_col_ids = []
        for c in range(1, nr_cols + 1):
            h = ws_rd.Cells(1, c).Value
            fbl3n_col_ids.append(str(h) if h is not None else f"Col{c}")

        for candidate in ["DMBTR", "WRBTR", "KWBTR", "HSL", "TSL",
                          "Amount in Doc. Curr.", "Amount in Local Currency"]:
            if candidate in fbl3n_col_ids:
                FBL3N_AMT_COL     = candidate
                fbl3n_amt_col_idx = fbl3n_col_ids.index(candidate) + 1
                print(f"  Kolumna kwot FBL3N: {FBL3N_AMT_COL}")
                break
        if not FBL3N_AMT_COL:
            print("  UWAGA: Nie rozpoznano kolumny kwot. Kolumny:", fbl3n_col_ids)

        fbl3n_row_count = nr_rows - 1
        print(f"  {fbl3n_row_count} wierszy, {nr_cols} kolumn")

        for r in range(2, nr_rows + 1):
            row_data = {}
            for c_idx, col_id in enumerate(fbl3n_col_ids, 1):
                val = ws_rd.Cells(r, c_idx).Value
                if col_id == FBL3N_AMT_COL and val is not None:
                    try:
                        val = round(float(val), 2)
                    except:
                        pass
                row_data[col_id] = val
            fbl3n_data.append(row_data)

        wb_rd.Close(False)
        print(f"  Wczytano {len(fbl3n_data)} wierszy")
    except Exception as e:
        print(f"  BLAD wczytywania pliku FBL3N: {e}")
else:
    print(f"  UWAGA: Nie znaleziono pliku eksportu FBL3N")
    print(f"  Oczekiwano: {FBL3N_EXPORT_XLSX}")

# Dodaj arkusz FBL3N do skoroszytu
print("Dodaje arkusz FBL3N do pliku Excel...")
ws_fbl3n = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
ws_fbl3n.Name = "FBL3N"

if fbl3n_col_ids:
    for i, col_id in enumerate(fbl3n_col_ids):
        ws_fbl3n.Cells(1, i + 1).Value = col_id
    for row_idx, row_data in enumerate(fbl3n_data):
        for col_idx, col_id in enumerate(fbl3n_col_ids):
            ws_fbl3n.Cells(row_idx + 2, col_idx + 1).Value = row_data.get(col_id, "")

    if fbl3n_amt_col_idx:
        try:
            ws_fbl3n.Columns(fbl3n_amt_col_idx).NumberFormat = "#,##0.00"
        except:
            pass

        print("Szukam par +/- w FBL3N...")
        fbl3n_positives = defaultdict(list)
        fbl3n_negatives = defaultdict(list)
        for row in range(len(fbl3n_data)):
            excel_row = row + 2
            val = ws_fbl3n.Cells(excel_row, fbl3n_amt_col_idx).Value
            if val is None:
                continue
            try:
                amt = round(float(val), 2)
            except:
                continue
            fbl3n_amounts[amt].append(excel_row)
            if amt > 0:
                fbl3n_positives[amt].append(excel_row)
            elif amt < 0:
                fbl3n_negatives[round(abs(amt), 2)].append(excel_row)

        matched_fbl3n = set(fbl3n_positives.keys()) & set(fbl3n_negatives.keys())
        rows_fbl3n_color = []
        for amt in matched_fbl3n:
            rows_fbl3n_color.extend(fbl3n_positives[amt])
            rows_fbl3n_color.extend(fbl3n_negatives[amt])
        for excel_row in rows_fbl3n_color:
            ws_fbl3n.Rows(excel_row).Interior.Color = LIGHT_GREEN
        print(f"FBL3N: {len(matched_fbl3n)} par, {len(rows_fbl3n_color)} wierszy pokolorowanych")

wb.Save()

# === POROWNANIE FEBAN <-> FBL3N ===
print("\nPorownuje kwoty FEBAN <-> FBL3N...")
common_amounts = set(feban_amounts.keys()) & set(fbl3n_amounts.keys())
match_count = 0
MATCH_TEXT = "Already in the Cashpool account"

for amt in common_amounts:
    if kwbtr_idx is not None:
        for excel_row in feban_amounts[amt]:
            cell = ws.Cells(excel_row, 24)
            existing = cell.Value
            if existing:
                cell.Value = str(existing) + " / " + MATCH_TEXT
            else:
                cell.Value = MATCH_TEXT
            cell.Font.Bold = True
            match_count += 1
    if fbl3n_amt_col_idx:
        fbl3n_note_col = len(fbl3n_col_ids) + 1
        for excel_row in fbl3n_amounts[amt]:
            cell = ws_fbl3n.Cells(excel_row, fbl3n_note_col)
            existing = cell.Value
            if existing:
                cell.Value = str(existing) + " / Already in FEBAN"
            else:
                cell.Value = "Already in FEBAN"
            cell.Font.Bold = True

wb.Save()
print(f"Oznaczono {match_count} wierszy FEBAN jako 'Already in the Cashpool account'")

wb.Close()
excel.Quit()

print(f"\nZapisano: {SAVE_PATH}")
os.startfile(SAVE_PATH)
