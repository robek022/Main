import win32com.client
from win32com.client import dynamic
import time

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = dynamic.Dispatch(SapGuiAuto.GetScriptingEngine)
connection = application.Children(0)
session = connection.Children(0)

print("Otwieram FBL3N...")
session.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL3N"
session.findById("wnd[0]").sendVKey(0)
time.sleep(2)

lines = []

def dump(obj, depth=0):
    if depth > 8:
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

OUTPUT = r"C:\Users\mrobak\fbl3n_fields.txt"
with open(OUTPUT, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"Pelne drzewo zapisane do: {OUTPUT}")
print()
print("=== Elementy zawierajace 'rad', 'date', 'stida', 'opvw', 'saknr', 'bukrs' ===")
for line in lines:
    low = line.lower()
    if any(x in low for x in ["rad", "date", "stida", "opvw", "saknr", "bukrs", "ctxt"]):
        print(line)
