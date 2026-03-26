"""
Diagnostic script – run this while the CommsApp Configuration dialog is open.
It prints every control inside the window so we can find the correct class name
for the port list.

Usage:
    1. Open SMARTset manually.
    2. Go to System -> FLAG Communications Server Setup...
    3. The CommsApp Configuration dialog should be open.
    4. In the VS Code terminal run:  python inspect_commsapp.py
"""

import time
import re
from pywinauto import Desktop, Application

print("Searching for CommsApp Configuration window...")

found = None
deadline = time.time() + 30
while time.time() < deadline:
    for w in Desktop(backend="win32").windows():
        if re.search(r"CommsApp", w.window_text(), re.IGNORECASE):
            found = w
            break
    if found:
        break
    time.sleep(1)

if not found:
    print("ERROR: CommsApp Configuration window not found. Is it open?")
    exit(1)

print(f"\nFound window: '{found.window_text()}'\n")
print("=" * 60)
print("All controls inside CommsApp Configuration:")
print("=" * 60)

app = Application(backend="win32").connect(handle=found.handle)
win = app.top_window()

# Print a full dump of every control
win.print_control_identifiers()
