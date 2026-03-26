"""
Diagnostic – run while the Scheme Editor is open (after double-clicking A1140 nonTOU).

Navigates to Load Profiling (row 9 in dgvPages) and dumps the full control tree
of the right panel (pnlEditor) so we can identify checkboxes, comboboxes, etc.

Usage:
    1. Open SMARTset, navigate to Imports, double-click A1140 nonTOU manually.
    2. The 'SMARTset - Scheme Editor - A1140 nonTOU' window should be open.
    3. Run:  python inspect_scheme_editor.py
"""
import re
import time
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys

print("Searching for Scheme Editor window...")

found = None
deadline = time.time() + 30
while time.time() < deadline:
    for w in Desktop(backend="win32").windows():
        if re.search(r"Scheme Editor", w.window_text(), re.IGNORECASE):
            found = w
            break
    if found:
        break
    time.sleep(1)

if not found:
    print("ERROR: Scheme Editor window not found.")
    exit(1)

handle = found.handle
print(f"\nFound: '{found.window_text()}' (handle={handle})\n")

app_uia = Application(backend="uia").connect(handle=handle)
win_uia = app_uia.window(handle=handle)

# Navigate to Load Profiling (row index 9 in dgvPages).
print("Navigating to Load Profiling (row 9)...")
dgv = win_uia.child_window(auto_id="dgvPages")
dgv.click_input()
time.sleep(0.3)
send_keys("^{HOME}")   # row 0 = Summary Page
time.sleep(0.3)
for _ in range(9):
    send_keys("{DOWN}")
    time.sleep(0.1)
time.sleep(0.8)   # let the right panel update
print("Done. Dumping right panel (pnlEditor)...\n")


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


# Dump pnlEditor (right panel) only.
print("=" * 70)
print("Right panel (pnlEditor) when Load Profiling is selected:")
print("=" * 70)
try:
    pnl = win_uia.child_window(auto_id="pnlEditor")
    dump(pnl)
except Exception as e:
    print(f"pnlEditor not found ({e}), dumping full window instead:")
    dump(win_uia)

# Also dump the full window for completeness.
print("\n" + "=" * 70)
print("Full window dump:")
print("=" * 70)
dump(win_uia)
