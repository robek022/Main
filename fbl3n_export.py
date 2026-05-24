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

    def try_fill_file_dialog(wnd_id):
        try:
            session2.findById(f"{wnd_id}/usr/ctxtDY_PATH").text = FBL3N_EXPORT_DIR
        except:
            pass
        try:
            session2.findById(f"{wnd_id}/usr/ctxtDY_FILENAME").text = FBL3N_TEMP_NAME
        except:
            pass
        try:
            session2.findById(f"{wnd_id}/tbar[0]/btn[0]").press()
            return True
        except:
            try:
                session2.findById(wnd_id).sendVKey(0)
                return True
            except:
                return False

    export_handled = False
    for first_wnd in ["wnd[1]", "wnd[2]"]:
        try:
            session2.findById(first_wnd)
        except:
            continue

        has_filename = False
        try:
            session2.findById(f"{first_wnd}/usr/ctxtDY_FILENAME")
            has_filename = True
        except:
            pass

        if has_filename:
            print(f"  Dialog zapisu w {first_wnd} - wpisuje sciezke...")
            try_fill_file_dialog(first_wnd)
            time.sleep(2)
            export_handled = True
            break
        else:
            print(f"  Dialog potwierdzenia w {first_wnd} - klikam OK...")
            try:
                session2.findById(f"{first_wnd}/tbar[0]/btn[0]").press()
            except:
                try:
                    session2.findById(first_wnd).sendVKey(0)
                except:
                    pass
            time.sleep(2)
            for save_wnd in ["wnd[1]", "wnd[2]"]:
                try:
                    session2.findById(f"{save_wnd}/usr/ctxtDY_FILENAME")
                    print(f"  Dialog zapisu w {save_wnd} - wpisuje sciezke...")
                    try_fill_file_dialog(save_wnd)
                    time.sleep(2)
                    export_handled = True
                    break
                except:
                    continue
            if export_handled:
                break

    if not export_handled:
        print("  UWAGA: Nie udalo sie automatycznie obsluzyc dialogu")
        print("  Jesli pojawil sie dialog, zapisz plik recznie i kontynuuj")

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
