import win32com.client
from win32com.client import dynamic
import time
import os
from datetime import datetime

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = dynamic.Dispatch(SapGuiAuto.GetScriptingEngine)
connection = application.Children(0)
session = connection.Children(0)

company_code = input("Wpisz numer spolki: ").strip()
today_str = datetime.now().strftime("%d.%m.%Y")
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

FBL3N_TEMP_NAME   = f"fbl3n_export_{timestamp}"
FBL3N_EXPORT_DIR  = r"C:\Users\mrobak\\"
FBL3N_EXPORT_XLSX = rf"C:\Users\mrobak\{FBL3N_TEMP_NAME}.xlsx"
FBL3N_EXPORT_XLS  = rf"C:\Users\mrobak\{FBL3N_TEMP_NAME}.xls"

print("Otwieram nowe okno SAP dla FBL3N...")
session.createSession()
time.sleep(3)
session2 = connection.Children(1)

print("Nawiguje do FBL3N...")
session2.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL3N"
session2.findById("wnd[0]").sendVKey(0)
time.sleep(2)

print(f"Wpisuje parametry (konto 10441000, spolka {company_code}, data {today_str})...")
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
    print("  Open items zaznaczone")
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

print("Wykonuje F8...")
session2.findById("wnd[0]").sendVKey(8)
time.sleep(4)

print("Eksportuje: List -> Export -> Spreadsheet...")
try:
    session2.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    time.sleep(2)

    # Krok 1: dialog wyboru formatu (Select Spreadsheet) - XLSX juz zaznaczony, klikamy OK
    try:
        fmt_dlg = session2.findById("wnd[1]")
        print(f"  Dialog formatu: '{fmt_dlg.Text}' - potwierdzam Enter...")
        fmt_dlg.sendVKey(0)
        time.sleep(3)
    except Exception as e:
        print(f"  (Brak dialogu formatu: {e})")

    # Krok 2: sprawdz czy pojawil sie dialog zapisu pliku
    file_saved = False
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
            file_saved = True
            print(f"  Zapisano jako: {FBL3N_EXPORT_XLSX}")
            break
        except:
            continue

    if not file_saved:
        print("  (Brak dialogu zapisu - SAP mogl otworzyc Excel bezposrednio)")

except Exception as e:
    print(f"  BLAD eksportu: {e}")

# Sprawdz czy plik istnieje
fbl3n_export_path = None
for candidate in [FBL3N_EXPORT_XLSX, FBL3N_EXPORT_XLS]:
    if os.path.exists(candidate):
        fbl3n_export_path = candidate
        print(f"\nPlik znaleziony: {fbl3n_export_path}")
        break

if not fbl3n_export_path:
    print(f"\nUWAGA: Nie znaleziono pliku eksportu")
    print(f"Oczekiwano: {FBL3N_EXPORT_XLSX}")
    print("Skrypt zakonczony.")
else:
    print("Otwieram plik Excel...")
    os.startfile(fbl3n_export_path)
    print("Gotowe!")
