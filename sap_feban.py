import win32com.client
from win32com.client import dynamic
import time
import os

# dynamic.Dispatch omija blad ISapComponentTarget z makepy
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = dynamic.Dispatch(SapGuiAuto.GetScriptingEngine)
connection = application.Children(0)
session = connection.Children(0)

print("Polaczono z SAP")

session.findById("wnd[0]").sendVKey(12)
time.sleep(0.5)
session.findById("wnd[0]").sendVKey(12)
time.sleep(0.5)

print("Otwieram FEBAN...")
session.findById("wnd[0]/tbar[0]/okcd").text = "/nFEBAN"
session.findById("wnd[0]").sendVKey(0)
time.sleep(2)

print("Wpisuje Company Code 2052...")
session.findById("wnd[1]/usr/ctxtSL_BUKRS-LOW").text = "2052"

print("Wykonuje F8...")
session.findById("wnd[1]").sendVKey(8)
time.sleep(3)

print("Klikam Spreadsheet...")
shell_raw = session.findById(
    "wnd[0]/usr/ssubAREA_N2P:FEB_BSPROC_FE:0113/cntlAREA_N2P/shellcont/shell"
)
shell = win32com.client.gencache.EnsureDispatch(shell_raw)
shell.pressToolbarButton("&SPREADSHEET")
time.sleep(1)

session.findById(
    "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]"
).select()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
time.sleep(1)

SAVE_PATH = r"C:\Users\mrobak\feban_raport.xlsx"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.path.dirname(SAVE_PATH) + "\\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = os.path.basename(SAVE_PATH)
session.findById("wnd[1]/tbar[0]/btn[0]").press()
time.sleep(2)

print(f"Zapisano plik: {SAVE_PATH}")
