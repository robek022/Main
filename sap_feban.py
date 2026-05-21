import win32com.client
import time

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
