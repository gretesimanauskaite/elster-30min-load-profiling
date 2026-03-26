"""
Diagnostic script - run this while the Amend a Connection dialog is open.

Usage:
    1. Open SMARTset manually.
    2. Go to System -> Maintain Connections...
    3. Select the 'New Connection' (FLAG Network) row.
    4. Click Amend.
    5. The 'Amend a Connection' dialog should be open.
    6. In the VS Code terminal run:  python inspect_amend_connection.py
"""

import re
import time
from pywinauto import Application, Desktop

print("Searching for Amend a Connection window...")

found = None
deadline = time.time() + 30
while time.time() < deadline:
    for w in Desktop(backend="win32").windows():
        t = w.window_text()
        if re.search(r"Amend", t, re.IGNORECASE) and re.search(r"Connection", t, re.IGNORECASE):
            found = w
            break
    if found:
        break
    time.sleep(1)

if not found:
    print("ERROR: 'Amend a Connection' window not found. Is it open?")
    exit(1)

handle = found.handle
print(f"\nFound window: '{found.window_text()}' (handle={handle})\n")

print("=" * 70)
print("win32 backend - print_control_identifiers():")
print("=" * 70)
app32 = Application(backend="win32").connect(handle=handle)
win32 = app32.top_window()
win32.print_control_identifiers()

print("\n" + "=" * 70)
print("uia backend - all descendants with text:")
print("=" * 70)
app_uia = Application(backend="uia").connect(handle=handle)
win_uia = app_uia.top_window()

def dump(ctrl, depth=0):
    indent = "  " * depth
    try:
        ct = ctrl.element_info.control_type
    except Exception:
        ct = "?"
    try:
        txt = ctrl.window_text()
    except Exception:
        txt = ""
    try:
        aid = ctrl.element_info.automation_id
    except Exception:
        aid = ""
    try:
        cls = ctrl.element_info.class_name
    except Exception:
        cls = ""
    print(f"{indent}[{ct}] auto_id='{aid}' class='{cls}' text='{txt}'")
    try:
        for child in ctrl.children():
            dump(child, depth + 1)
    except Exception:
        pass

dump(win_uia)