import win32com.client
from win32com.client import dynamic

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = dynamic.Dispatch(SapGuiAuto.GetScriptingEngine)
connection = application.Children(0)
session = connection.Children(0)

def find_shells(obj, depth=0):
    if depth > 6:
        return
    prefix = "  " * depth
    try:
        obj_id = obj.Id
        obj_type = obj.Type
        print(f"{prefix}{obj_id}  [{obj_type}]")
        try:
            count = obj.Children.Count
            for i in range(count):
                find_shells(obj.Children(i), depth + 1)
        except:
            pass
    except:
        pass

print("=== Struktura wnd[0]/usr ===")
find_shells(session.findById("wnd[0]/usr"))
