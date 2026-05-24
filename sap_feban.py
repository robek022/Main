import win32com.client
from win32com.client import dynamic
import time
import os
import re
from datetime import datetime
from collections import defaultdict

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = dynamic.Dispatch(SapGuiAuto.GetScriptingEngine)
connection = application.Children(0)
session = connection.Children(0)

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
SAVE_PATH = rf"C:\Users\mrobak\feban_raport_{company_code}_{timestamp}.xlsx"

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

print(f"Wpisuje parametry FBL3N (konto 10441000, bukrs 2052, data {today_str})...")

try:
    session2.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").text = "10441000"
except Exception as e:
    print(f"  UWAGA: Nie znaleziono pola G/L Account: {e}")

try:
    session2.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = "2052"
except Exception as e:
    print(f"  UWAGA: Nie znaleziono pola Company Code: {e}")

try:
    session2.findById("wnd[0]/usr/radX_OPSEL").select()
    print("  Zaznaczono Open items")
except Exception as e:
    print(f"  UWAGA: Nie znaleziono radio Open items: {e}")

try:
    session2.findById("wnd[0]/usr/ctxtPA_STIDA").text = today_str
    print(f"  Ustawiono date: {today_str}")
except Exception as e:
    print(f"  UWAGA: Nie znaleziono pola daty: {e}")

print("Wykonuje F8 w FBL3N...")
session2.findById("wnd[0]").sendVKey(8)
time.sleep(4)

# Ctrl+F9 = Select Layout (VKey 33)
print("Otwieram wybor layoutu (Ctrl+F9)...")
session2.findById("wnd[0]").sendVKey(33)
time.sleep(2)

print("Wybieram layout FEBAN2052MR...")
layout_found = False

try:
    table = session2.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")
    for r in range(table.RowCount):
        try:
            if str(table.GetCellValue(r, "VARIANT")).strip().upper() == "FEBAN2052MR":
                table.setCurrentCell(r, "VARIANT")
                table.doubleClickCurrentCell()
                layout_found = True
                print(f"  Layout znaleziony w wierszu {r} (ALV grid)")
                break
        except:
            continue
except Exception as e:
    print(f"  Proba 1 (ALV grid) nieudana: {e}")

if not layout_found:
    try:
        table = session2.findById("wnd[1]/usr/lsT_VARIANT")
        for r in range(table.RowCount):
            try:
                if str(table.GetCellValue(r, "VARIANT")).strip().upper() == "FEBAN2052MR":
                    table.setCurrentCell(r, "VARIANT")
                    table.doubleClickCurrentCell()
                    layout_found = True
                    print(f"  Layout znaleziony w wierszu {r} (lista)")
                    break
            except:
                continue
    except Exception as e:
        print(f"  Proba 2 (lista) nieudana: {e}")

if not layout_found:
    try:
        session2.findById("wnd[1]/usr/txtV-LOW").text = "FEBAN2052MR"
        session2.findById("wnd[1]").sendVKey(0)
        time.sleep(0.5)
        session2.findById("wnd[1]").sendVKey(2)
        layout_found = True
        print("  Layout wybrany przez pole filtra")
    except Exception as e:
        print(f"  Proba 3 (filtr) nieudana: {e}")

if not layout_found:
    print("  UWAGA: Nie udalo sie automatycznie wybrac layoutu - wybierz recznie FEBAN2052MR")
else:
    time.sleep(2)

# Czytaj dane FBL3N z gridu
fbl3n_data = []
fbl3n_col_ids = []
FBL3N_AMT_COL = None
fbl3n_amt_col_idx = None
fbl3n_row_count = 0

print("Czytam dane z gridu FBL3N...")
try:
    fbl3n_shell = session2.findById("wnd[0]/shellcont/shell")

    prev = -1
    while True:
        current = fbl3n_shell.RowCount
        if current == prev:
            break
        prev = current
        try:
            fbl3n_shell.firstVisibleRow = current
        except:
            break
        time.sleep(0.3)
    try:
        fbl3n_shell.firstVisibleRow = 0
    except:
        pass
    time.sleep(0.5)

    fbl3n_row_count = fbl3n_shell.RowCount
    print(f"FBL3N: {fbl3n_row_count} wierszy")

    try:
        for col in fbl3n_shell.ColumnOrder:
            fbl3n_col_ids.append(str(col))
    except:
        pass

    for candidate in ["DMBTR", "WRBTR", "KWBTR", "HSL", "TSL", "AMOUNT"]:
        if candidate in fbl3n_col_ids:
            FBL3N_AMT_COL = candidate
            fbl3n_amt_col_idx = fbl3n_col_ids.index(candidate) + 1
            print(f"  Kolumna kwot FBL3N: {FBL3N_AMT_COL}")
            break
    if not FBL3N_AMT_COL:
        print("  UWAGA: Nie rozpoznano kolumny kwot w FBL3N. Dostepne kolumny:", fbl3n_col_ids)

    for row in range(fbl3n_row_count):
        try:
            fbl3n_shell.setCurrentCell(row, fbl3n_col_ids[0])
            if row % 10 == 0:
                time.sleep(0.3)
        except:
            pass
        row_data = {}
        for col_id in fbl3n_col_ids:
            try:
                val = fbl3n_shell.GetCellValue(row, col_id)
            except:
                val = ""
            if col_id == FBL3N_AMT_COL:
                val = parse_sap_amount(val)
            row_data[col_id] = val
        fbl3n_data.append(row_data)
        print(f"  FBL3N wiersz {row+1}/{fbl3n_row_count}")

except Exception as e:
    print(f"  BLAD przy czytaniu gridu FBL3N: {e}")

# Dodaj arkusz FBL3N do skoroszytu
print("Dodaje arkusz FBL3N...")
ws_fbl3n = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
ws_fbl3n.Name = "FBL3N"

fbl3n_amounts = defaultdict(list)  # amt -> [excel_rows]

if fbl3n_col_ids:
    for i, col_id in enumerate(fbl3n_col_ids):
        ws_fbl3n.Cells(1, i + 1).Value = col_id

    for row_idx, row_data in enumerate(fbl3n_data):
        for col_idx, col_id in enumerate(fbl3n_col_ids):
            ws_fbl3n.Cells(row_idx + 2, col_idx + 1).Value = row_data[col_id]

    if fbl3n_amt_col_idx:
        try:
            ws_fbl3n.Columns(fbl3n_amt_col_idx).NumberFormat = "#,##0.00"
        except:
            pass

        # Pary +/- w FBL3N
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
comment_count = 0

for amt in common_amounts:
    amt_label = f"{amt:,.2f}".replace(",", " ")
    if kwbtr_idx is not None:
        for excel_row in feban_amounts[amt]:
            cell = ws.Cells(excel_row, kwbtr_idx + 1)
            try:
                cell.Comment.Delete()
            except:
                pass
            cell.AddComment(f"Kwota {amt_label} znaleziona rowniez w FBL3N")
            comment_count += 1
    if fbl3n_amt_col_idx:
        for excel_row in fbl3n_amounts[amt]:
            cell = ws_fbl3n.Cells(excel_row, fbl3n_amt_col_idx)
            try:
                cell.Comment.Delete()
            except:
                pass
            cell.AddComment(f"Kwota {amt_label} znaleziona rowniez w FEBAN")
            comment_count += 1

wb.Save()
print(f"Dodano {comment_count} komentarzy dla {len(common_amounts)} pasujacych kwot")

wb.Close()
excel.Quit()

print(f"\nZapisano: {SAVE_PATH}")
os.startfile(SAVE_PATH)
