import win32com.client
import os
import math
from datetime import datetime
from collections import defaultdict
from tkinter import Tk, filedialog

# === WYBOR PLIKU ===
Tk().withdraw()
export_path = filedialog.askopenfilename(
    title="Wybierz plik eksportu FBL3N",
    filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
)

if not export_path:
    print("Nie wybrano pliku.")
    input("Nacisnij Enter aby zamknac...")
    raise SystemExit(1)

print(f"Wybrany plik: {export_path}")

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
SAVE_DIR  = os.path.join(os.path.expanduser("~"), "Documents")
os.makedirs(SAVE_DIR, exist_ok=True)
base_name = os.path.splitext(os.path.basename(export_path))[0]
SAVE_PATH = os.path.join(SAVE_DIR, f"{base_name}_clearing_{timestamp}.xlsx")

# === EXCEL ===
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb_src = excel.Workbooks.Open(export_path)
    ws_src = wb_src.Sheets(1)
    nr_rows = ws_src.UsedRange.Rows.Count
    nr_cols = ws_src.UsedRange.Columns.Count

    col_ids = []
    for c in range(1, nr_cols + 1):
        h = ws_src.Cells(1, c).Value
        col_ids.append(str(h) if h is not None else f"Col{c}")

    # Wykryj kolumne kwot
    amt_col_idx = None
    AMT_COL = None
    for candidate in ["Amount in Local Currency", "DMBTR", "HSL",
                      "Amount in Doc. Curr.", "WRBTR", "KWBTR", "TSL"]:
        if candidate in col_ids:
            AMT_COL     = candidate
            amt_col_idx = col_ids.index(candidate) + 1
            print(f"Kolumna kwot: {AMT_COL} (kolumna {amt_col_idx})")
            break
    if not AMT_COL:
        print("UWAGA: Nie rozpoznano kolumny kwot.")
        print("Dostepne kolumny:", col_ids)
        amt_col_name = input("Wpisz nazwe kolumny z kwotami: ").strip()
        if amt_col_name in col_ids:
            AMT_COL     = amt_col_name
            amt_col_idx = col_ids.index(amt_col_name) + 1
        else:
            print("Nie znaleziono podanej kolumny.")
            input("Nacisnij Enter aby zamknac...")
            raise SystemExit(1)

    row_count = nr_rows - 1
    print(f"{row_count} wierszy, {nr_cols} kolumn")

    # Skopiuj dane do nowego skoroszytu
    wb = excel.Workbooks.Add()
    ws = wb.Sheets(1)
    ws.Name = "FBL3N"

    for c in range(1, nr_cols + 1):
        ws.Cells(1, c).Value = ws_src.Cells(1, c).Value
    for r in range(2, nr_rows + 1):
        for c in range(1, nr_cols + 1):
            ws.Cells(r, c).Value = ws_src.Cells(r, c).Value

    wb_src.Close(False)
    wb.SaveAs(SAVE_PATH)
    print(f"Zapisano jako: {SAVE_PATH}")

    # Formatowanie kolumny kwot
    try:
        ws.Columns(amt_col_idx).NumberFormat = "#,##0.00"
    except:
        pass

    LIGHT_GREEN = 144 + 238 * 256 + 144 * 65536
    LIGHT_BLUE  = 173 + 216 * 256 + 230 * 65536

    group_col = nr_cols + 1
    ws.Cells(1, group_col).Value = "Clearing Group"

    # === PARY +/- ===
    print("\nSzukam par +/- do clearowania...")
    positives = defaultdict(list)
    negatives = defaultdict(list)

    for r in range(2, nr_rows + 1):
        val = ws.Cells(r, amt_col_idx).Value
        if val is None:
            continue
        try:
            amt = round(float(val), 2)
        except:
            continue
        if amt > 0:
            positives[amt].append(r)
        elif amt < 0:
            negatives[round(abs(amt), 2)].append(r)

    matched   = set(positives.keys()) & set(negatives.keys())
    green_rows = set()
    group_num  = 1

    for amt in matched:
        for r in positives[amt] + negatives[amt]:
            ws.Rows(r).Interior.Color = LIGHT_GREEN
            ws.Cells(r, group_col).Value = group_num
            green_rows.add(r)
        group_num += 1

    wb.Save()
    print(f"Znaleziono {len(matched)} par, pokolorowano {len(green_rows)} wierszy na zielono")

    # === TROJKI SUMUJACE SIE DO 0 ===
    print("\nSzukam trojek sumujacych sie do 0...")
    unmatched = []
    for r in range(2, nr_rows + 1):
        if r in green_rows:
            continue
        val = ws.Cells(r, amt_col_idx).Value
        if val is None:
            continue
        try:
            amt = round(float(val), 2)
            if amt != 0:
                unmatched.append((amt, r))
        except:
            continue

    print(f"Niezmatchowanych wierszy do sprawdzenia: {len(unmatched)}")

    blue_rows   = set()
    int_to_rows = defaultdict(list)
    for amt, r in unmatched:
        int_to_rows[round(amt * 100)].append(r)

    n = len(unmatched)
    for i in range(n):
        ai_int = round(unmatched[i][0] * 100)
        ri     = unmatched[i][1]
        if ri in blue_rows:
            continue
        for j in range(i + 1, n):
            aj_int = round(unmatched[j][0] * 100)
            rj     = unmatched[j][1]
            if rj in blue_rows:
                continue
            needed = -(ai_int + aj_int)
            if needed in int_to_rows:
                for rk in int_to_rows[needed]:
                    if rk not in blue_rows and rk != ri and rk != rj:
                        for r in [ri, rj, rk]:
                            ws.Rows(r).Interior.Color = LIGHT_BLUE
                            ws.Cells(r, group_col).Value = group_num
                            blue_rows.add(r)
                        group_num += 1
                        break

    wb.Save()
    blue_groups = group_num - 1 - len(matched)
    remaining   = row_count - len(green_rows) - len(blue_rows)
    print(f"Znaleziono {blue_groups} trojek = 0, pokolorowano {len(blue_rows)} wierszy na niebiesko")
    print(f"Bez dopasowania: {remaining} wierszy (biale)")
    print(f"Lacznie grup clearowania: {group_num - 1}")

    wb.Save()

except Exception as e:
    print(f"BLAD: {e}")
    try:
        excel.Quit()
    except:
        pass
    input("Nacisnij Enter aby zamknac...")
    raise SystemExit(1)

excel.Visible = True
wb.Activate()

print(f"\nGotowe! Zapisano: {SAVE_PATH}")
