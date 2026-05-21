import win32com.client
from win32com.client import dynamic
import time
import os

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

print("Otwieram menu Export...")
shell = session.findById("wnd[0]/shellcont/shell")
shell.pressToolbarButton("&MB_EXPORT")
time.sleep(1)

print("Wybieram Spreadsheet...")
shell.selectContextMenuItem("&SPREADSHEET")
time.sleep(2)

print("Klikam OK w popup formatu...")
# "Select from All Available Formats" jest juz zaznaczone - klikamy OK
session.findById("wnd[1]/tbar[0]/btn[0]").press()
time.sleep(2)

print("Podaje sciezke zapisu...")
SAVE_PATH = r"C:\Users\mrobak\feban_raport.xlsx"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.path.dirname(SAVE_PATH) + "\\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = os.path.basename(SAVE_PATH)
session.findById("wnd[1]/tbar[0]/btn[0]").press()
time.sleep(2)

print(f"Zapisano plik: {SAVE_PATH}")
