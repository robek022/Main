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

print("Polaczono z SAP")

print("Otwieram FEBAN...")
session.findById("wnd[0]/tbar[0]/okcd").text = "/nFEBAN"
session.findById("wnd[0]").sendVKey(0)
time.sleep(2)

print("Wpisuje Company Code 2052...")
session.findById("wnd[1]/usr/ctxtSL_BUKRS-LOW").text = "2052"

print("Wykonuje F8...")
session.findById("wnd[1]").sendVKey(8)
time.sleep(3)

shell = session.findById("wnd[0]/shellcont/shell")
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
    """Jezeli nota to sama liczba (np. payment run ref), zamien na int."""
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

# Faza 1: wczytaj dane z gridu
print("Czytam dane z gridu FEBAN...")
grid_data = []
for row in range(row_count):
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

# Sprawdz sciezke na pierwszym wierszu
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

# Zapis do Excela
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
SAVE_PATH = rf"C:\Users\mrobak\feban_raport_{timestamp}.xlsx"

print("Zapisuje do Excela...")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Add()
ws = wb.Sheets(1)

# Naglowki - kolumny SAP + Note to Payee na koncu
all_cols = col_ids + ["Note to Payee"]
for i, col_id in enumerate(all_cols):
    ws.Cells(1, i + 1).Value = col_id

# Dane
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
    print(f"  (Formatowanie kolumny pominiete: {e})")

try:
    note_col = len(col_ids) + 1
    last_row = row_count + 1
    ws.Range(ws.Cells(1, note_col), ws.Cells(last_row, note_col)).HorizontalAlignment = -4131
    ws.Range(ws.Cells(1, note_col), ws.Cells(last_row, note_col)).NumberFormat = "#,##0"
    wb.Save()
except Exception as e:
    print(f"  (Formatowanie AA pominiete: {e})")

# Szukaj par +/- w kolumnie KWBTR i koloruj wiersze na jasnozielono
if kwbtr_idx is not None:
    print("Szukam par +/- w kolumnie KWBTR...")

    positives = defaultdict(list)
    negatives = defaultdict(list)

    for row in range(len(grid_data)):
        excel_row = row + 2
        val = ws.Cells(excel_row, kwbtr_idx + 1).Value
        if val is None:
            continue
        try:
            amt = round(float(val), 2)
        except:
            continue
        if amt > 0:
            positives[amt].append(excel_row)
        elif amt < 0:
            negatives[round(abs(amt), 2)].append(excel_row)

    matched = set(positives.keys()) & set(negatives.keys())
    LIGHT_GREEN = 144 + 238 * 256 + 144 * 65536  # RGB(144, 238, 144)

    rows_to_color = []
    for amt in matched:
        rows_to_color.extend(positives[amt])
        rows_to_color.extend(negatives[amt])

    for excel_row in rows_to_color:
        ws.Rows(excel_row).Interior.Color = LIGHT_GREEN

    wb.Save()
    print(f"Znaleziono {len(matched)} par, pokolorowano {len(rows_to_color)} wierszy")

# Koloruj wiersze gdzie Note to Payee to sama liczba (payment run) na jasnorozowy
note_col = len(col_ids) + 1
LIGHT_PINK = 255 + 220 * 256 + 220 * 65536  # RGB(255, 220, 220)
pink_count = 0
for row in range(len(grid_data)):
    excel_row = row + 2
    val = ws.Cells(excel_row, note_col).Value
    if val is not None and isinstance(val, (int, float)):
        ws.Rows(excel_row).Interior.Color = LIGHT_PINK
        pink_count += 1
wb.Save()
print(f"Pokolorowano {pink_count} wierszy na rozowy (payment run)")

# Szukaj tekstu w kolumnie B i koloruj wiersze na jasnozolty
SEARCH_TEXT = "@5D\\QPosting in Subledger Accounting Made as On Account Posting@"
LIGHT_YELLOW = 255 + 255 * 256 + 153 * 65536  # RGB(255, 255, 153)
print("Szukam tekstu w kolumnie B...")
yellow_count = 0
for row in range(len(grid_data)):
    excel_row = row + 2
    val = ws.Cells(excel_row, 2).Value
    if val and SEARCH_TEXT in str(val):
        ws.Rows(excel_row).Interior.Color = LIGHT_YELLOW
        yellow_count += 1
wb.Save()
print(f"Pokolorowano {yellow_count} wierszy na zolty")

wb.Close()
excel.Quit()

print(f"Zapisano: {SAVE_PATH}")
os.startfile(SAVE_PATH)
