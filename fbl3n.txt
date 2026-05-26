import win32com.client
from win32com.client import dynamic
import time
import os
from datetime import datetime
from collections import defaultdict

# === POLACZENIE Z SAP ===
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

# === PARAMETRY ===
account      = input("Wpisz numer konta G/L: ").strip()
company_code = input("Wpisz numer spolki: ").strip()
date_from    = input("Posting date OD (dd.mm.rrrr): ").strip()
date_to      = input("Posting date DO (dd.mm.rrrr): ").strip()
timestamp    = datetime.now().strftime("%Y%m%d_%H%M%S")

SAVE_DIR = os.path.join(os.path.expanduser("~"), "Documents")
os.makedirs(SAVE_DIR, exist_ok=True)

TEMP_NAME   = f"fbl3n_{account}_{company_code}_{timestamp}"
EXPORT_DIR  = SAVE_DIR + "\\"
EXPORT_XLSX = os.path.join(SAVE_DIR, TEMP_NAME + ".xlsx")
EXPORT_XLS  = os.path.join(SAVE_DIR, TEMP_NAME + ".xls")
SAVE_PATH   = os.path.join(SAVE_DIR, f"fbl3n_raport_{account}_{company_code}_{timestamp}.xlsx")

print(f"\nKonto: {account} | Spolka: {company_code} | Posting date: {date_from} - {date_to}")

# === FBL3N ===
print("Nawiguje do FBL3N...")
session.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL3N"
session.findById("wnd[0]").sendVKey(0)
time.sleep(2)

print("Wpisuje parametry...")
try:
    session.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").text = account
except Exception as e:
    print(f"  UWAGA: G/L Account: {e}")
try:
    session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = company_code
except Exception as e:
    print(f"  UWAGA: Company Code: {e}")
try:
    session.findById("wnd[0]/usr/radX_AISEL").select()
    print("  All items zaznaczone")
except Exception as e:
    print(f"  UWAGA: All items: {e}")
try:
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = date_from
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = date_to
    print(f"  Posting date: {date_from} - {date_to}")
except Exception as e:
    print(f"  UWAGA: Posting date: {e}")
try:
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/FEBANMR"
    print("  Layout: /FEBANMR")
except Exception as e:
    print(f"  UWAGA: Layout: {e}")

print("Wykonuje F8...")
session.findById("wnd[0]").sendVKey(8)
time.sleep(4)

# === EKSPORT DO EXCELA ===
print("Eksportuje: List -> Export -> Spreadsheet...")
try:
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    time.sleep(2)

    try:
        fmt_dlg = session.findById("wnd[1]")
        print(f"  Dialog formatu: '{fmt_dlg.Text}' - potwierdzam...")
        fmt_dlg.sendVKey(0)
        time.sleep(3)
    except Exception as e:
        print(f"  (Brak dialogu formatu: {e})")

    export_handled = False
    for save_wnd in ["wnd[1]", "wnd[2]"]:
        try:
            session.findById(f"{save_wnd}/usr/ctxtDY_FILENAME")
            print(f"  Dialog zapisu w {save_wnd} - wpisuje sciezke...")
            try:
                session.findById(f"{save_wnd}/usr/ctxtDY_PATH").text = EXPORT_DIR
            except:
                pass
            session.findById(f"{save_wnd}/usr/ctxtDY_FILENAME").text = TEMP_NAME + ".xlsx"
            try:
                session.findById(f"{save_wnd}/tbar[0]/btn[0]").press()
            except:
                session.findById(save_wnd).sendVKey(0)
            time.sleep(3)
            export_handled = True
            print(f"  Zapisano: {EXPORT_XLSX}")
            break
        except:
            continue

    if not export_handled:
        print("  (Brak dialogu zapisu)")

except Exception as e:
    print(f"  BLAD eksportu: {e}")

# === WCZYTAJ PLIK I PRZETWORZ ===
export_path = None
for candidate in [EXPORT_XLSX, EXPORT_XLS]:
    if os.path.exists(candidate):
        export_path = candidate
        break

if not export_path:
    print(f"\nUWAGA: Nie znaleziono pliku eksportu: {EXPORT_XLSX}")
    input("Nacisnij Enter aby zamknac...")
    raise SystemExit(1)

print(f"\nWczytuje dane z: {export_path}")
excel = win32com.client.Dispatch("Excel.Application")
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
    for candidate in ["DMBTR", "WRBTR", "KWBTR", "HSL", "TSL",
                      "Amount in Doc. Curr.", "Amount in Local Currency"]:
        if candidate in col_ids:
            AMT_COL     = candidate
            amt_col_idx = col_ids.index(candidate) + 1
            print(f"  Kolumna kwot: {AMT_COL} (kolumna {amt_col_idx})")
            break
    if not AMT_COL:
        print("  UWAGA: Nie rozpoznano kolumny kwot. Kolumny:", col_ids)

    row_count = nr_rows - 1
    print(f"  {row_count} wierszy, {nr_cols} kolumn")

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
    print(f"  Zapisano jako: {SAVE_PATH}")

    # Formatowanie kolumny kwot
    if amt_col_idx:
        try:
            ws.Columns(amt_col_idx).NumberFormat = "#,##0.00"
        except:
            pass

        # === PARY +/- (do clearowania) ===
        print("Szukam par +/- do clearowania...")
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

        matched = set(positives.keys()) & set(negatives.keys())
        LIGHT_GREEN = 144 + 238 * 256 + 144 * 65536

        colored = 0
        for amt in matched:
            for r in positives[amt] + negatives[amt]:
                ws.Rows(r).Interior.Color = LIGHT_GREEN
                colored += 1

        wb.Save()
        print(f"  Znaleziono {len(matched)} par, pokolorowano {colored} wierszy na zielono")

        # Wiersze bez pary - brak koloru (zostaja biale)
        print(f"  Wiersze bez pary: {row_count - colored} (biale)")

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
