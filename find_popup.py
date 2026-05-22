import win32com.client
from win32com.client import dynamic
import time

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = dynamic.Dispatch(SapGuiAuto.GetScriptingEngine)
connection = application.Children(0)
session = connection.Children(0)

shell = session.findById("wnd[0]/shellcont/shell")
col_ids = []
for col in shell.ColumnOrder:
    col_ids.append(str(col))

print("Klikam wiersz 0...")
shell.setCurrentCell(0, col_ids[0])
shell.doubleClickCurrentCell()
time.sleep(1.5)

print("Klikam wiersz 1 zeby wywolac popup...")
shell.setCurrentCell(1, col_ids[0])
shell.doubleClickCurrentCell()
time.sleep(1.5)

try:
    wnd1 = session.findById("wnd[1]")
    print("Znaleziono popup wnd[1] - dumpuje strukture:\n")

    def dump(obj, depth=0):
        if depth > 6:
            return
        prefix = "  " * depth
        try:
            obj_id = obj.Id
            obj_type = obj.Type
            extra = ""
            try:
                t = obj.text
                if t and str(t).strip():
                    extra = f"  >>> text='{str(t)[:80]}'"
            except:
                pass
            print(f"{prefix}{obj_id}  [{obj_type}]{extra}")
            try:
                count = obj.Children.Count
                for i in range(count):
                    dump(obj.Children(i), depth + 1)
            except:
                pass
        except:
            pass

    dump(wnd1)

except Exception as e:
    print(f"Nie znaleziono wnd[1]: {e}")
    print("Popup moze jeszcze nie wyszedl albo juz zniknal")
