import win32com.client
from win32com.client import dynamic
import time
import os
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

kwbtr_idx = col_ids.index("KWBTR") if "KWBTR" in col_ids else None

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
SAVE_PATH = rf"C:\Users\mrobak\feban_raport_{timestamp}.xlsx"

print("Czytam dane z FEBAN...")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Add()
ws = wb.Sheets(1)

for i, col_id in enumerate(col_ids):
    ws.Cells(1, i + 1).Value = col_id

for row in range(row_count):
    for col_idx, col_id in enumerate(col_ids):
        try:
            val = shell.GetCellValue(row, col_id)
        except:
            val = ""
        if col_idx == kwbtr_idx:
            val = parse_sap_amount(val)
        ws.Cells(row + 2, col_idx + 1).Value = val
    print(f"  Wiersz {row+1}/{row_count}")

print("Zapisuje plik...")
wb.SaveAs(SAVE_PATH)

try:
    if kwbtr_idx is not None:
        ws.Columns(kwbtr_idx + 1).NumberFormat = "#,##0.00"
    wb.Save()
except Exception as e:
    print(f"  (Formatowanie kolumny pominiete: {e})")

# Szukaj par +/- w kolumnie KWBTR i koloruj wiersze na jasnozielono
if kwbtr_idx is not None:
    print("Szukam par +/- w kolumnie KWBTR...")

    positives = defaultdict(list)   # kwota -> lista excel_row
    negatives = defaultdict(list)

    for row in range(row_count):
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

# Szukaj tekstu w kolumnie B i koloruj wiersze na jasnozolty
SEARCH_TEXT = "@5D\\QPosting in Subledger Accounting Made as On Account Posting@"
LIGHT_YELLOW = 255 + 255 * 256 + 153 * 65536  # RGB(255, 255, 153)
print("Szukam tekstu w kolumnie B...")
yellow_count = 0
for row in range(row_count):
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
