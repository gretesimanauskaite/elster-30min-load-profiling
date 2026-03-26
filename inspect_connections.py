"""
Diagnostic script – run this while the Browse Connections dialog is open.
Prints the full control tree so we can find the correct class/type for the
connection list rows.

Usage:
    1. Open SMARTset manually.
    2. Go to System -> Maintain Connections...
    3. The 'SMARTset - Browse Connections' dialog should be open.
    4. In the VS Code terminal run:  python inspect_connections.py
"""

import re
import time
from pywinauto import Application, Desktop

print("Searching for Browse Connections window...")

found = None
deadline = time.time() + 30
while time.time() < deadline:
    for w in Desktop(backend="win32").windows():
        if re.search(r"Browse Connections", w.window_text(), re.IGNORECASE):
            found = w
            break
    if found:
        break
    time.sleep(1)

if not found:
    print("ERROR: Browse Connections window not found. Is it open?")
    exit(1)

handle = found.handle
print(f"\nFound window: '{found.window_text()}' (handle={handle})\n")

# ── win32 backend dump ───────────────────────────────────────────────────────
print("=" * 70)
print("win32 backend – print_control_identifiers():")
print("=" * 70)
app32 = Application(backend="win32").connect(handle=handle)
win32 = app32.top_window()
win32.print_control_identifiers()

# ── uia backend dump ─────────────────────────────────────────────────────────
print("\n" + "=" * 70)
print("uia backend – print_control_identifiers():")
print("=" * 70)
app_uia = Application(backend="uia").connect(handle=handle)
win_uia = app_uia.top_window()
win_uia.print_control_identifiers()

# ── uia: enumerate every child with its text ─────────────────────────────────
print("\n" + "=" * 70)
print("uia backend – all descendants with window_text():")
print("=" * 70)
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
        cls = ctrl.element_info.class_name
    except Exception:
        cls = ""
    print(f"{indent}[{ct}] class='{cls}' text='{txt}'")
    try:
        for child in ctrl.children():
            dump(child, depth + 1)
    except Exception:
        pass

dump(win_uia)
