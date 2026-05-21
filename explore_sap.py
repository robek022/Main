import win32com.client

SapGuiAuto = win32com.client.GetObject("SAPGUI")
session = SapGuiAuto.GetScriptingEngine.Children(0).Children(0)

print("=== CHILDREN OF wnd[0]/usr ===")
usr = session.findById("wnd[0]/usr")
for i in range(usr.Children.Count):
    ch = usr.Children(i)
    print(f"  [{i}] {ch.Id}  type={ch.Type}")

print("\n=== MENU BAR ===")
mbar = session.findById("wnd[0]/mbar")
for i in range(mbar.Children.Count):
    m = mbar.Children(i)
    print(f"  menu[{i}]: {m.Text}")
    for j in range(m.Children.Count):
        print(f"    menu[{i}][{j}]: {m.Children(j).Text}")
