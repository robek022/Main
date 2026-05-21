import win32com.client
import time
import os

# --- Połączenie z SAP ---
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# --- Wejście do FEBAN ---
session.findById("wnd[0]/tbar[0]/okcd").text = "/nfeban"
session.findById("wnd[0]").sendVKey(0)
time.sleep(1)

# --- Wpisanie Company Code 2052 ---
session.findById("wnd[1]/usr/ctxtSL_BUKRS-LOW").text = "2052"

# --- Execute (F8) ---
session.findById("wnd[1]").sendVKey(8)
time.sleep(2)

# --- Export do Excel przez przycisk Spreadsheet w ALV grid ---
shell = session.findById(
    "wnd[0]/usr/ssubAREA_N2P:FEB_BSPROC_FE:0113/cntlAREA_N2P/shellcont/shell"
)
shell.pressToolbarButton("&SPREADSHEET")
time.sleep(1)

# --- Popup wyboru formatu - wybierz XLSX (pozycja 2) ---
session.findById(
    "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]"
).select()
session.findById("wnd[1]/tbar[0]/btn[0]").press()  # OK
time.sleep(1)

# --- Dialog zapisu pliku ---
SAVE_PATH = r"C:\Users\Public\feban_raport.xlsx"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.path.dirname(SAVE_PATH) + "\\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = os.path.basename(SAVE_PATH)
session.findById("wnd[1]/tbar[0]/btn[0]").press()  # Generuj / OK
time.sleep(2)

print(f"Zapisano plik: {SAVE_PATH}")
