"""
Diagnostic – run while the SMARTset Scheme Manager is open (no other dialogs).
Prints the full control tree of the main window so we can find the correct
identifiers for the tree (left panel) and list (right panel) used in Step 4.

Usage:
    python inspect_scheme_manager.py
"""
import re
import time
from pywinauto import Application, Desktop

print("Searching for SMARTset Scheme Manager window...")

found = None
deadline = time.time() + 30
while time.time() < deadline:
    for w in Desktop(backend="win32").windows():
        if re.search(r"Scheme Manager", w.window_text(), re.IGNORECASE):
            found = w
            break
    if found:
        break
    time.sleep(1)

if not found:
    print("ERROR: Scheme Manager window not found.")
    exit(1)

handle = found.handle
print(f"\nFound: '{found.window_text()}' (handle={handle})\n")

# win32
print("=" * 70)
print("win32 backend:")
print("=" * 70)
app32 = Application(backend="win32").connect(handle=handle)
win32 = app32.window(handle=handle)
win32.print_control_identifiers()

# uia full dump
print("\n" + "=" * 70)
print("uia backend – all descendants:")
print("=" * 70)
app_uia = Application(backend="uia").connect(handle=handle)
win_uia = app_uia.window(handle=handle)

def dump(ctrl, depth=0):
    indent = "  " * depth
    try: ct = ctrl.element_info.control_type
    except: ct = "?"
    try: txt = ctrl.window_text()
    except: txt = ""
    try: aid = ctrl.element_info.automation_id
    except: aid = ""
    try: cls = ctrl.element_info.class_name
    except: cls = ""
    print(f"{indent}[{ct}] auto_id='{aid}' class='{cls}' text='{txt}'")
    try:
        for child in ctrl.children():
            dump(child, depth + 1)
    except: pass

dump(win_uia)
