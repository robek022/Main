import win32com.client
from win32com.client import dynamic
import time

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = dynamic.Dispatch(SapGuiAuto.GetScriptingEngine)
connection = application.Children(0)
session = connection.Children(0)

print("Otwieram FEBAN...")
session.findById("wnd[0]/tbar[0]/okcd").text = "/nFEBAN"
session.findById("wnd[0]").sendVKey(0)
time.sleep(2)
session.findById("wnd[1]/usr/ctxtSL_BUKRS-LOW").text = "2052"
session.findById("wnd[1]").sendVKey(8)
time.sleep(3)

shell = session.findById("wnd[0]/shellcont/shell")
print("Klikam pierwszy wiersz...")
shell.setCurrentCell(0, shell.ColumnOrder[0] if hasattr(shell.ColumnOrder, '__iter__') else "HKTID")
time.sleep(1.5)

lines = []

def dump(obj, depth=0):
    if depth > 10:
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
        lines.append(f"{prefix}{obj_id}  [{obj_type}]{extra}")
        try:
            count = obj.Children.Count
            for i in range(count):
                dump(obj.Children(i), depth + 1)
        except:
            pass
    except:
        pass

dump(session.findById("wnd[0]"))

OUTPUT = r"C:\Users\mrobak\feban_note_tree.txt"
with open(OUTPUT, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"Pelne drzewo zapisane do: {OUTPUT}")
print()
print("=== Elementy zawierajace 'note', 'payee', 'n2p', 'sgtxt', 'txt' ===")
for line in lines:
    low = line.lower()
    if any(x in low for x in ["note", "payee", "n2p", "sgtxt", "txt"]):
        print(line)
