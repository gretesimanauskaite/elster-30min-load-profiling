"""
SMARTset Automated Configuration Script
========================================
Automates the following tasks in SMARTset:
  1. Launch SMARTset and log in.
  2. System > FLAG Communications Server Setup
       - Select NET port, update IP, Save.
  3. System > Maintain Connections
       - Find / create "New Connection" (FLAG Network), set Host IP + Outstation.
  4. Imports > A1140 nonTOU  [skipped when --no-scheme is passed]
       - Untick everything except Load Profiling.
       - In Load Profiling: enable Demand Period (30 min), tick Import W / Export W /
         Q1 Inductive Import / VA 1.

RUN COMMANDS (copy/paste into the VS Code terminal)
====================================================

  --- Normal use ---

  Full run – all 4 steps (login, CommsApp IP, connection, scheme):
    python smartset_configure.py --ip 10.0.19.37 --serial 38110126

  Steps 1-3 only – skip the scheme editor entirely:
    python smartset_configure.py --ip 10.0.19.37 --serial 38110126 --no-scheme

  --- Safety / review ---

  Dry run – read current values, print what would change, touch NOTHING:
    python smartset_configure.py --ip 10.0.19.37 --serial 38110126 --dry-run

  Dry run, steps 1-3 only:
    python smartset_configure.py --ip 10.0.19.37 --serial 38110126 --dry-run --no-scheme

  --- Scripted / unattended ---

  Skip the Y/N confirmation prompts:
    python smartset_configure.py --ip 10.0.19.37 --serial 38110126 --yes

  Steps 1-3, no prompts:
    python smartset_configure.py --ip 10.0.19.37 --serial 38110126 --no-scheme --yes

  --- Diagnostics (run while the relevant dialog is open in SMARTset) ---

  Inspect CommsApp Configuration controls:
    python inspect_commsapp.py

  Inspect Browse Connections controls:
    python inspect_connections.py

  Inspect Amend a Connection controls:
    python inspect_amend_connection.py

  Inspect Scheme Manager (tree + list) controls:
    python inspect_scheme_manager.py

  Inspect Scheme Editor (navigates to Load Profiling, dumps right panel):
    python inspect_scheme_editor.py

  --- Install dependencies (run once) ---
    pip install pywinauto pyautogui

Arguments:
  --ip          IP address of the FLAG Communications Server (e.g. 10.0.19.37)
  --serial      Full meter serial number; last 3 digits become Outstation (e.g. 38110126 → 126)
  --no-scheme   Skip Step 4 (scheme editor) – only update CommsApp + connection settings
  --dry-run     Read and report current values without making any changes
  --yes         Auto-confirm all Y/N prompts (use with care on production meters)
"""

import argparse
import csv
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

try:
    import pyautogui
    from pywinauto import Application, Desktop
    from pywinauto.keyboard import send_keys
except ImportError:
    print("ERROR: Required packages missing. Run:  pip install pywinauto pyautogui")
    sys.exit(1)

# --------------------------------------------------------------------------- #
# Runtime flag – set once in main() from --dry-run CLI argument.
# When True, every GUI mutation is logged but never executed.
# --------------------------------------------------------------------------- #
DRY_RUN: bool = False

# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #
# Full path to the SMARTset shortcut / executable on this machine.
SMARTSET_EXE = (
    r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
    r"\ELSTER\SMARTset V1.1 (Standard)\SMARTset V1.1 (Standard).lnk"
)
USERNAME = "Elster"
PASSWORD = "Elster"

# How long (seconds) to wait after launching SMARTset before interacting with it.
STARTUP_WAIT = 15   # seconds – SMARTset can be slow to show the login dialog

# How long to wait for a dialog/window to appear after triggering a menu action.
DIALOG_WAIT  = 1.5

# Short pause injected between individual GUI actions to let the UI settle.
ACTION_PAUSE = 0.4

# pyautogui safety: moving the mouse to the top-left corner will abort the script.
pyautogui.FAILSAFE = True
pyautogui.PAUSE    = ACTION_PAUSE

# --------------------------------------------------------------------------- #
# Scheme page lists
# --------------------------------------------------------------------------- #

# Pages in the Scheme Editor that must have NO green tick when done.
PAGES_TO_CLEAR = [
    "Meter Identifiers",
    "Passwords",
    "Tariff/Display",
    "Deferred Tariff/Display",
    "Deferred Setup",
    "Billing",
    "Meter Constants",
    "Relay Setup",
    "Meter Options",
    "Meter UI Options",
    "Time and Date",
]

# The only Load Profiling channel checkboxes that should be TICKED.
# All others in ALL_CHANNELS will be explicitly unticked.
LOAD_PROFILE_CHANNELS = ["Import W", "Export W", "Q1 Inductive Import", "VA 1"]

# Every channel checkbox that exists in the Load Profiling panel.
ALL_CHANNELS = [
    "Import W",
    "Export W",
    "Q1 Inductive Import",
    "Q2 Capacitive Import",
    "Q3 Inductive Export",
    "Q4 Capacitive Export",
    "VA 1",
    "VA 2",
    "Customer Defined 1",
    "Customer Defined 2",
]

# Fixed row indices for the dgvPages DataGridView in the Scheme Editor.
# Row 0 = Summary Page (never cleared); ordering confirmed by diagnostic.
PAGE_ROW: dict[str, int] = {
    "Summary Page":            0,
    "Meter Identifiers":       1,
    "Passwords":               2,
    "Tariff/Display":          3,
    "Deferred Tariff/Display": 4,
    "Deferred Setup":          5,
    "Billing":                 6,
    "Meter Constants":         7,
    "Relay Setup":             8,
    "Load Profiling":          9,
    "Meter Options":          10,
    "Meter UI Options":       11,
    "Time and Date":          12,
}


# --------------------------------------------------------------------------- #
# Audit log
# Rows are collected in memory during the run, then written to a CSV at the end.
# The CSV is named  audit_<serial>_<timestamp>.csv  and saved next to this script.
# --------------------------------------------------------------------------- #
_AUDIT_ROWS: list[dict] = []


def _audit(step: str, field: str, before: str, after: str, applied: bool):
    """Append one row to the in-memory audit log."""
    _AUDIT_ROWS.append({
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "step":      step,
        "field":     field,
        "before":    before,
        "after":     after,
        # Mark whether the change was actually applied or was a dry-run record.
        "applied":   "DRY RUN" if not applied else "YES",
    })


def _write_audit_log(serial: str) -> Path:
    """Write all collected audit rows to a timestamped CSV file."""
    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = Path(__file__).parent / f"audit_{serial}_{ts}.csv"
    with path.open("w", newline="") as fh:
        writer = csv.DictWriter(
            fh,
            fieldnames=["timestamp", "step", "field", "before", "after", "applied"],
        )
        writer.writeheader()
        writer.writerows(_AUDIT_ROWS)
    print(f"\nAudit log written to: {path}")
    return path


# --------------------------------------------------------------------------- #
# GUI action wrappers
#
# Every action that changes a GUI control goes through one of these functions.
# This means --dry-run is enforced in a single place, and every action is
# automatically recorded in the audit log whether it ran or not.
# --------------------------------------------------------------------------- #

def _gui_click(control, description: str, step: str):
    """Click a button or control, or log the intended click in dry-run mode."""
    if DRY_RUN:
        print(f"  [DRY RUN] Would click: {description}")
        _audit(step, description, "", "CLICK", applied=False)
    else:
        print(f"  Clicking: {description}")
        control.click()
        _audit(step, description, "", "CLICK", applied=True)
    time.sleep(ACTION_PAUSE)


def _gui_set_text(control, new_value: str, field_name: str, step: str):
    """Type a value into a text field, replacing whatever was there."""
    # window_text() works for both Delphi TEdit (win32) and WinForms (uia).
    try:
        before = control.window_text().strip()
    except Exception:
        before = "<unknown>"

    if DRY_RUN:
        print(f"  [DRY RUN] Would set '{field_name}': '{before}' → '{new_value}'")
        _audit(step, field_name, before, new_value, applied=False)
    else:
        print(f"  Setting '{field_name}': '{before}' → '{new_value}'")
        # click_input() + Home→Shift+End→Delete selects the whole line and
        # clears it reliably for both Delphi TEdit and WinForms TextBox.
        # (Ctrl+A is ignored by Delphi TEdit with no explicit handler.)
        control.click_input()
        time.sleep(0.1)
        send_keys("{HOME}+{END}{DELETE}")
        time.sleep(0.1)
        control.type_keys(new_value, with_spaces=False)
        _audit(step, field_name, before, new_value, applied=True)
    time.sleep(ACTION_PAUSE)


def _gui_select(control, value, field_name: str, step: str):
    """Select an item in a ComboBox by text or index."""
    try:
        before = control.selected_text() if hasattr(control, "selected_text") else "<unknown>"
    except Exception:
        before = "<unknown>"

    if DRY_RUN:
        print(f"  [DRY RUN] Would select '{field_name}': '{before}' → '{value}'")
        _audit(step, field_name, before, str(value), applied=False)
    else:
        print(f"  Selecting '{field_name}': '{before}' → '{value}'")
        control.select(value)
        _audit(step, field_name, before, str(value), applied=True)
    time.sleep(ACTION_PAUSE)


def _gui_checkbox(control, desired: bool, label: str, step: str):
    """Tick or untick a checkbox, but only if it isn't already in the desired state."""
    try:
        current_state = control.get_toggle_state() == 1
    except Exception:
        current_state = None

    # No-op if already in the right state – avoids unnecessary clicks.
    if current_state == desired:
        print(f"  Checkbox '{label}' already {'ticked' if desired else 'unticked'} – skipping.")
        return

    before = str(current_state)
    after  = str(desired)

    if DRY_RUN:
        print(f"  [DRY RUN] Would {'tick' if desired else 'untick'} checkbox '{label}'")
        _audit(step, f"checkbox:{label}", before, after, applied=False)
    else:
        print(f"  {'Ticking' if desired else 'Unticking'} checkbox '{label}'")
        control.click()
        _audit(step, f"checkbox:{label}", before, after, applied=True)
    time.sleep(ACTION_PAUSE)


def _keyboard_menu(win, menu_path: str):
    """
    Navigate a menu using Alt + arrow keys.

    SMARTset uses Delphi VCL menus: pywinauto menu_select() raises
    'There is no menu', and letter-jump shortcuts are not supported.
    Arrow-key counting is reliable because SMARTset is a fixed application
    whose menu structure does not change.

    Steps:
      1. Press Alt  → activates menu bar, File is highlighted.
      2. Press Right N times → reach the target top-level menu (System = 2).
      3. Press Down → opens the dropdown; first item is highlighted.
      4. Press Down (position-1) times → reach the target item.
      5. Press Enter → activate it.

    System menu item positions (1-based, from SMARTset screenshots):
      1  Configure Auto-Import...
      2  Configure Time Set Mode...
      3  View Audit Trail...
      4  Change Password...
      5  Application Settings...
      6  Configure Database Locations...
      7  FLAG Communications Server Setup...
      8  View FLAG Communications Log...
      9  Launch FLAG Communications Server...
      10 Maintain Connections...
      11 Backup Databases
      12 Language
    """
    ITEM_POSITIONS = {
        "System->FLAG Communications Server Setup...": 7,
        "System->Maintain Connections...":             10,
    }

    TOP_MENU_RIGHTS = {
        "File":   0,
        "Window": 1,
        "System": 2,
        "Help":   3,
    }

    parts    = [p.strip() for p in menu_path.split("->")]
    top_menu = parts[0]
    rights   = TOP_MENU_RIGHTS.get(top_menu, 0)
    position = ITEM_POSITIONS.get(menu_path, 1)

    win.set_focus()
    time.sleep(ACTION_PAUSE)

    # Activate menu bar — lands on File.
    send_keys("%")
    time.sleep(0.4)

    # Move right to the target top-level menu.
    for _ in range(rights):
        send_keys("{RIGHT}")
        time.sleep(0.2)

    # Open the dropdown.
    send_keys("{DOWN}")
    time.sleep(0.4)

    # Navigate down to the target item (item 1 already selected after open).
    for _ in range(position - 1):
        send_keys("{DOWN}")
        time.sleep(0.15)

    send_keys("{ENTER}")
    time.sleep(DIALOG_WAIT)




# --------------------------------------------------------------------------- #
# Window helpers
# --------------------------------------------------------------------------- #

def _wait_for_window(title_re: str, timeout: int = 15) -> Application:
    """
    Scan all open windows on the desktop until one whose title matches
    title_re appears, then return an Application connected to it by handle.
    Using Desktop().windows() is more reliable than Application.connect()
    because it doesn't depend on process ownership or title encoding quirks.
    """
    import re
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            for w in Desktop(backend="win32").windows():
                if re.search(title_re, w.window_text(), re.IGNORECASE):
                    app = Application(backend="win32").connect(handle=w.handle)
                    return app
        except Exception:
            pass
        time.sleep(1)
    raise TimeoutError(f"Window matching '{title_re}' did not appear within {timeout}s")


def _get_main_app(timeout: int = 30) -> Application:
    """
    Wait for the SMARTset main window (shown after login).
    Its title is 'SMARTset - [Scheme Manager]', distinct from the login dialog.
    """
    return _wait_for_window(r"SMARTset.*Scheme Manager", timeout=timeout)


def _wait_gone(title_re: str, timeout: float = 3.0):
    """
    Poll Desktop windows until none match title_re, or timeout expires.
    Title-based scanning is more reliable than holding a stale window handle.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        if not any(
            re.search(title_re, w.window_text(), re.IGNORECASE)
            for w in Desktop(backend="win32").windows()
        ):
            break
        time.sleep(0.1)


def _close_window(win):
    """
    Close a dialog without saving anything.
    Tries Cancel first (safest), then Close, then the window's X button.
    """
    for btn_title in ("Cancel", "Close"):
        for search_kwargs in (
            {"title": btn_title,        "control_type": "Button"},
            {"title": "&" + btn_title,  "control_type": "Button"},
            {"title": btn_title,        "class_name": "TButton"},
            {"title": btn_title,        "class_name": "Button"},
        ):
            try:
                win.child_window(**search_kwargs).click()
                time.sleep(0.2)
                return
            except Exception:
                pass
    try:
        win.close()
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# Step 1 – Launch SMARTset and log in
# --------------------------------------------------------------------------- #

def _smartset_already_running() -> bool:
    """Return True if the SMARTset main window (Scheme Manager) is already open."""
    import re
    try:
        for w in Desktop(backend="win32").windows():
            if re.search(r"SMARTset.*Scheme Manager", w.window_text(), re.IGNORECASE):
                return True
    except Exception:
        pass
    return False


def launch_and_login() -> Application:
    """
    Launch the SMARTset executable and log in.
    If the main Scheme Manager window is already open, skip launch/login entirely
    and ask the user whether to proceed directly to Step 2.
    """
    import re

    # ── Already running? ─────────────────────────────────────────────
    if _smartset_already_running():
        print("\n── Step 1: SMARTset is already running ───────────────────────")
        print("  The Scheme Manager window is open — no login needed.")
        return _get_main_app()

    # ── Not running — launch it ───────────────────────────────────────
    print("\n── Step 1: Launch SMARTset and log in ────────────────────────")
    os.startfile(SMARTSET_EXE)
    time.sleep(STARTUP_WAIT)

    # Wait for the login dialog (title "SMARTset", distinct from main window).
    print("   Waiting for login dialog…")
    login_dlg = None
    deadline  = time.time() + 60

    while time.time() < deadline:
        try:
            for w in Desktop(backend="win32").windows():
                title = w.window_text()
                if re.search(r"SMARTset", title, re.IGNORECASE) and "Scheme Manager" not in title:
                    login_dlg = w
                    break
        except Exception:
            pass
        if login_dlg:
            break
        time.sleep(1)

    if login_dlg is None:
        print("   No login dialog found – assuming already authenticated.")
        return _get_main_app()

    print(f"   Login dialog found. Entering credentials…")
    login_dlg.set_focus()
    time.sleep(ACTION_PAUSE)

    # Tab order: User Name (ComboBox) → Password (Edit) → OK → Cancel
    send_keys("^a")
    send_keys(USERNAME)
    send_keys("{TAB}")
    send_keys("^a")
    send_keys(PASSWORD)
    send_keys("{ENTER}")
    _audit("Login", "credentials", "", "submitted", applied=True)

    print("   Credentials submitted. Waiting for main window…")
    time.sleep(DIALOG_WAIT)
    print("   Login complete.")
    return _get_main_app()


# --------------------------------------------------------------------------- #
# Snapshot helpers
#
# These functions open each dialog, read the current values, then close
# WITHOUT saving.  They run even in --dry-run mode so you always get a
# before-picture to compare against the planned changes.
# --------------------------------------------------------------------------- #

def _snapshot_comms_server(app) -> dict:
    """
    Read the current NET IP from CommsApp Configuration.
    Opens the dialog, reads the value, then cancels – no changes made.
    """
    snapshot = {"net_ip": "<could not read>"}
    try:
        main_win = app.top_window()
        main_win.set_focus()
        _keyboard_menu(main_win, "System->FLAG Communications Server Setup...")
        time.sleep(DIALOG_WAIT)

        comms_app = _wait_for_window(r"(?i)CommsApp Configuration", timeout=10)
        comms_win = comms_app.top_window()

        # The port list may be a standard ListView or a Delphi TListView.
        try:
            port_list = comms_win.child_window(control_type="List")
        except Exception:
            port_list = comms_win.child_window(class_name="TListView")

        # Scroll to the bottom so the NET row is visible, then click it.
        port_list.set_focus()
        send_keys("{END}")
        time.sleep(0.5)

        items = port_list.items() if hasattr(port_list, "items") else []
        for item in items:
            if item.text().strip().upper() == "NET":
                item.click()
                time.sleep(ACTION_PAUSE)
                break

        # Read the IP address field that appears on the right after selecting NET.
        try:
            ip_field = comms_win.child_window(control_type="Edit", found_index=0)
            snapshot["net_ip"] = ip_field.get_value()
        except Exception:
            pass

        # Cancel – we are only reading, not saving.
        _close_window(comms_win)
    except Exception as exc:
        snapshot["error"] = str(exc)
    return snapshot


def _snapshot_connection(app) -> dict:
    """
    Read the Host and Outstation of the FLAG Network connection.
    Opens Amend dialog in read mode, reads values, then cancels.
    """
    snapshot = {"host": "<could not read>", "outstation": "<could not read>"}
    try:
        main_win = app.top_window()
        main_win.set_focus()
        _keyboard_menu(main_win, "System->Maintain Connections...")
        time.sleep(DIALOG_WAIT)

        _wait_for_window(r"(?i)(Browse Connections|Maintain Connections)", timeout=10)
        _browse_handle = next(
            (w.handle for w in Desktop(backend="win32").windows()
             if re.search(r"(?i)(Browse Connections|Maintain Connections)", w.window_text())),
            None,
        )
        if not _browse_handle:
            raise RuntimeError("Browse Connections window not found.")
        browse_app_win32 = Application(backend="win32").connect(handle=_browse_handle)
        browse_w32 = browse_app_win32.window(handle=_browse_handle)

        # Navigate to the FLAG Network row using keyboard (same approach as CommsApp).
        # Ctrl+End moves to the last row which is "New Connection" (FLAG Network).
        conn_grid_w32 = browse_w32.child_window(auto_id="dgvResults")
        conn_grid_w32.click_input()
        time.sleep(0.3)
        send_keys("^{END}")
        time.sleep(0.4)

        # Click Amend via win32 backend.
        amend_clicked = False
        try:
            browse_w32.child_window(auto_id="btnAction2").click()
            amend_clicked = True
        except Exception:
            pass
        if not amend_clicked:
            for search_kwargs in (
                {"title": "&Amend", "class_name": "Button"},
                {"title": "Amend",  "class_name": "Button"},
            ):
                try:
                    browse_w32.child_window(**search_kwargs).click()
                    amend_clicked = True
                    break
                except Exception:
                    continue

        time.sleep(DIALOG_WAIT)
        _wait_for_window(r"(?i)Amend a Connection", timeout=10)
        _amend_handle = next(
            (w.handle for w in Desktop(backend="win32").windows()
             if re.search(r"(?i)Amend a Connection", w.window_text())),
            None,
        )
        # Use win32 backend: window_text() returns actual values ('10.0.19.37', '126').
        # uia window_text() returns only the accessible label ('Host', 'Outstation').
        amend_win = (Application(backend="win32").connect(handle=_amend_handle).window(handle=_amend_handle)
                     if _amend_handle else Application(backend="win32").connect(process=browse_app_win32.process).top_window())

        try:
            snapshot["host"] = amend_win.child_window(
                auto_id="pnlHost"
            ).child_window(best_match="HostEdit").window_text()
            snapshot["outstation"] = amend_win.child_window(
                auto_id="pnlOutstation"
            ).child_window(best_match="OutstationEdit").window_text()
        except Exception:
            try:
                snapshot["host"]       = amend_win.child_window(best_match="HostEdit").window_text()
                snapshot["outstation"] = amend_win.child_window(best_match="OutstationEdit").window_text()
            except Exception:
                pass

        # Cancel both dialogs – we are only reading.
        _close_window(amend_win)
        _close_window(browse_w32)
    except Exception as exc:
        snapshot["error"] = str(exc)
    return snapshot


def _snapshot_scheme(app) -> dict:
    """
    Read current page assignments and Load Profiling channel states from the
    Scheme Editor.  Opens and cancels without saving anything.
    Only called when --no-scheme is NOT set.
    """
    snapshot = {"pages": {}, "channels": {}}
    try:
        main_win = app.top_window()
        main_win.set_focus()

        # Navigate to the scheme: expand Imports, double-click A1140 nonTOU.
        tree = main_win.child_window(control_type="Tree")
        imports_node = tree.get_item(["Imports"])
        imports_node.expand()
        time.sleep(ACTION_PAUSE)
        tree.get_item(["Imports", "A1140 nonTOU"]).double_click()
        time.sleep(DIALOG_WAIT)

        editor_app = _wait_for_window(r"(?i)Scheme Editor.*A1140", timeout=15)
        editor_win = editor_app.top_window()

        try:
            page_tree = editor_win.child_window(control_type="Tree")
        except Exception:
            page_tree = editor_win.child_window(class_name="TTreeView")

        # Read the assigned page name for every page we care about.
        for page_name in PAGES_TO_CLEAR + ["Load Profiling"]:
            try:
                page_tree.get_item([page_name]).click()
                time.sleep(ACTION_PAUSE)
                combos = editor_win.children(control_type="ComboBox")
                val    = combos[0].selected_text() if combos else "<unknown>"
                snapshot["pages"][page_name] = val
            except Exception:
                snapshot["pages"][page_name] = "<could not read>"

        # Read which Load Profiling channel checkboxes are currently ticked.
        try:
            page_tree.get_item(["Load Profiling"]).click()
            time.sleep(ACTION_PAUSE)
            for ch in ALL_CHANNELS:
                for title_variant in (ch, ch.replace(" ", "")):
                    try:
                        cb = editor_win.child_window(
                            title=title_variant, control_type="CheckBox"
                        )
                        snapshot["channels"][ch] = cb.get_toggle_state() == 1
                        break
                    except Exception:
                        continue
                else:
                    snapshot["channels"][ch] = "<not found>"
        except Exception:
            pass

        # Cancel – we have only been reading.
        _close_window(editor_win)
    except Exception as exc:
        snapshot["error"] = str(exc)
    return snapshot


def snapshot_all_settings(app, include_scheme: bool = True) -> dict:
    """
    Read the current state of all relevant SMARTset settings.
    Pass include_scheme=False when running with --no-scheme to avoid
    opening the Scheme Editor unnecessarily.
    """
    print("\n--- Reading current SMARTset settings (no changes made) ---")
    result = {}
    result["comms"]      = _snapshot_comms_server(app)
    result["connection"] = _snapshot_connection(app)

    if include_scheme:
        result["scheme"] = _snapshot_scheme(app)
    else:
        # Scheme snapshot is intentionally skipped – return empty placeholders
        # so that print_snapshot / print_plan don't crash.
        result["scheme"] = {"pages": {}, "channels": {}}
        print("   Scheme snapshot skipped (--no-scheme).")

    return result


# --------------------------------------------------------------------------- #
# Display helpers – print current state and planned changes side-by-side
# --------------------------------------------------------------------------- #

def print_snapshot(snapshot: dict):
    """Print a summary of the values read from SMARTset before any changes."""
    print("\n┌─ CURRENT STATE (before any changes) ─────────────────────┐")

    c = snapshot.get("comms", {})
    print(f"│  CommsApp NET IP       : {c.get('net_ip', '<unknown>')}")

    cn = snapshot.get("connection", {})
    print(f"│  Connection Host       : {cn.get('host', '<unknown>')}")
    print(f"│  Connection Outstation : {cn.get('outstation', '<unknown>')}")

    s = snapshot.get("scheme", {})
    if s.get("pages"):
        print("│  Scheme page assignments:")
        for page, val in s["pages"].items():
            # Mark pages that are already assigned (non-None) with an asterisk.
            marker = "  " if val.lower() in ("<none>", "none", "") else "* "
            print(f"│    {marker}{page:<30} {val}")

    if s.get("channels"):
        print("│  Load Profiling channels:")
        for ch, state in s["channels"].items():
            tick = "✓" if state is True else ("✗" if state is False else "?")
            print(f"│    [{tick}] {ch}")

    print("└───────────────────────────────────────────────────────────┘")


def print_plan(new_ip: str, outstation: str, snapshot: dict):
    """
    Show exactly what will be changed – a diff between current state and
    what the script intends to set.  Lines beginning with ~ are changes;
    lines beginning with = are no-ops.
    """
    print("\n┌─ PLANNED CHANGES ─────────────────────────────────────────┐")

    c   = snapshot.get("comms", {})
    cn  = snapshot.get("connection", {})
    s   = snapshot.get("scheme", {})

    _diff("CommsApp NET IP",       c.get("net_ip", "?"),          new_ip)
    _diff("Connection Host",       cn.get("host", "?"),           new_ip)
    _diff("Connection Outstation", cn.get("outstation", "?"),     outstation)

    # Only show scheme diffs when the scheme snapshot was collected.
    if s.get("pages"):
        for page in PAGES_TO_CLEAR:
            old_val = s["pages"].get(page, "?")
            if old_val.lower() not in ("<none>", "none", ""):
                _diff(f"Scheme page '{page}'", old_val, "<None>")

        lp_old = s["pages"].get("Load Profiling", "?")
        if lp_old.lower() in ("<none>", "none", ""):
            print("│  + Scheme page 'Load Profiling' will be assigned")

    if s.get("channels"):
        for ch in ALL_CHANNELS:
            desired = ch in LOAD_PROFILE_CHANNELS
            current = s["channels"].get(ch)
            if current != desired:
                action = "TICK" if desired else "UNTICK"
                print(f"│  ~ Channel '{ch}': {action}")

    print("└───────────────────────────────────────────────────────────┘")


def _diff(label: str, old: str, new: str):
    """Print one line of the plan – = for no-op, ~ for a real change."""
    if str(old).strip() == str(new).strip():
        print(f"│  = {label:<32} (no change: {old})")
    else:
        print(f"│  ~ {label:<32} '{old}' → '{new}'")


# --------------------------------------------------------------------------- #
# Step 2 – FLAG Communications Server Setup
# Opens dialog → reads current IP → asks Y/N → updates + saves → closes
# --------------------------------------------------------------------------- #

def _open_comms_app(app):
    """Open CommsApp, navigate to NET row, return (comms_win, ip_field)."""
    main_win = app.top_window()
    main_win.set_focus()
    _keyboard_menu(main_win, "System->FLAG Communications Server Setup...")
    comms_app = _wait_for_window(r"(?i)CommsApp Configuration", timeout=10)
    comms_win = comms_app.top_window()

    port_list = comms_win.child_window(class_name="TStringGrid")

    # First click to focus the grid.
    port_list.click_input()
    time.sleep(0.3)

    # Ctrl+End moves keyboard selection to the last row (NET) and scrolls it
    # into view, but does NOT fire Delphi's OnClick event which populates the
    # Port configuration panel on the right.  We must also physically click
    # the NET row to trigger that event.
    send_keys("^{END}")
    time.sleep(0.6)

    # Click near the bottom of the grid where NET is now visible.
    # Using relative coords (x=centre of grid, y=near bottom) so it works
    # regardless of where the window is on screen.
    rect = port_list.rectangle()
    rel_x = (rect.right - rect.left) // 2
    rel_y = (rect.bottom - rect.top) - 12   # 12 px above the bottom edge
    port_list.click_input(coords=(rel_x, rel_y))
    time.sleep(0.5)

    port_config = comms_win.child_window(
        title="Port configuration", class_name="TGroupBox"
    )
    ip_field = port_config.child_window(class_name="TEdit", found_index=0)
    return comms_win, ip_field


def _close_comms_app(comms_win):
    """Click Cancel to close CommsApp and wait until the window is gone."""
    for btn_class in ("TButton", "Button"):
        try:
            comms_win.child_window(title="Cancel", class_name=btn_class).click()
            break
        except Exception:
            pass
    else:
        try:
            comms_win.close()
        except Exception:
            pass
    # Wait for the window to actually disappear (max 5 s).
    # exists() returns False (not an exception) when the handle is gone,
    # so we must check the return value, not catch an exception.
    deadline = time.time() + 5
    while time.time() < deadline:
        try:
            if not comms_win.exists():
                break
        except Exception:
            break
        time.sleep(0.3)


def configure_comms_server(app, new_ip: str):
    """
    Step 2 – FLAG Communications Server Setup.
    Opens CommsApp, reads the current NET IP, shows it, asks the user
    whether to apply the change, then saves and closes (or just closes
    if the user says no).
    """
    print("\n── Step 2: FLAG Communications Server Setup ──────────────────")
    comms_win, ip_field = _open_comms_app(app)

    # Read current value.  window_text() is reliable for Delphi TEdit;
    # get_value() sends EM_GETLINE which some Delphi controls ignore.
    try:
        current_ip = ip_field.window_text().strip()
    except Exception:
        current_ip = "<could not read>"

    print(f"  Current NET IP : {current_ip}")
    print(f"  New NET IP     : {new_ip}")

    if current_ip == new_ip:
        print("  No change needed — values already match.")
        _close_comms_app(comms_win)
        return

    answer = input("\n  Apply this change? [y/N]: ").strip().lower()
    if answer not in ("y", "yes"):
        print("  Skipped by user.")
        _close_comms_app(comms_win)
        return

    # Update IP.
    # Delphi TEdit ignores Ctrl+A (no built-in "select all"), so we use
    # Home → Shift+End to select the full line, then Delete to clear it.
    ip_field.click_input()
    time.sleep(0.2)
    send_keys("{HOME}+{END}{DELETE}")
    time.sleep(0.1)
    ip_field.type_keys(new_ip, with_spaces=False)
    _audit("CommsApp", "NET IP address", current_ip, new_ip, applied=True)

    # Save.
    saved = False
    for btn_class in ("TButton", "Button"):
        try:
            comms_win.child_window(title="Save", class_name=btn_class).click()
            saved = True
            break
        except Exception:
            pass
    if not saved:
        send_keys("%s")   # Alt+S fallback
    time.sleep(DIALOG_WAIT)
    print("  Saved.")

    # Close.
    _close_comms_app(comms_win)
    print("  CommsApp closed. Ready for next step.")


# --------------------------------------------------------------------------- #
# Step 3 – Maintain Connections
# Opens dialog → reads current values → asks Y/N → updates + saves → closes
# --------------------------------------------------------------------------- #

def _open_maintain_connections(app):
    """
    Open Maintain Connections, find the FLAG Network row, click Amend,
    and return (browse_win, amend_win, host_field, outstation_field).
    """
    main_win = app.top_window()
    main_win.set_focus()
    _keyboard_menu(main_win, "System->Maintain Connections...")

    # Scan Desktop for the exact Browse Connections handle.
    # top_window() is unreliable — it can return the main SMARTset frame.
    _wait_for_window(r"(?i)(Browse Connections|Maintain Connections)", timeout=10)
    _browse_handle = next(
        (w.handle for w in Desktop(backend="win32").windows()
         if re.search(r"(?i)(Browse Connections|Maintain Connections)", w.window_text())),
        None,
    )
    if not _browse_handle:
        raise RuntimeError("Browse Connections window not found.")
    browse_app_win32 = Application(backend="win32").connect(handle=_browse_handle)
    browse_win32 = browse_app_win32.window(handle=_browse_handle)
    browse_app = Application(backend="uia").connect(handle=_browse_handle)
    browse_win = browse_app.window(handle=_browse_handle)

    # The connection list is a WinForms DataGridView (auto_id="dgvResults").
    # UIA does not reliably expose its rows as children in this application.
    # Instead, use keyboard navigation: click the grid to focus it, then
    # Ctrl+End jumps to the last row which is always "New Connection" (FLAG Network).
    conn_grid_win32 = browse_win32.child_window(auto_id="dgvResults")
    conn_grid_win32.click_input()
    time.sleep(0.3)
    send_keys("^{END}")   # move selection to last row = New Connection
    time.sleep(0.4)
    print("  Navigated to last row (FLAG Network / New Connection).")

    # Click Amend using the stable auto_id from the win32-backed browse window.
    amend_opened = False
    try:
        browse_win32.child_window(auto_id="btnAction2").click()
        amend_opened = True
    except Exception:
        pass
    if not amend_opened:
        for search_kwargs in (
            {"title": "&Amend", "class_name": "Button"},
            {"title": "Amend",  "class_name": "Button"},
        ):
            try:
                browse_win32.child_window(**search_kwargs).click()
                amend_opened = True
                break
            except Exception:
                pass
    if not amend_opened:
        conn_grid_win32.double_click_input()

    time.sleep(DIALOG_WAIT)

    amend_app_win32 = _wait_for_window(r"(?i)Amend a Connection", timeout=10)
    # top_window() can return the main SMARTset frame rather than the Amend dialog
    # when other windows are open.  Scan Desktop to get the exact dialog handle.
    _amend_handle = next(
        (w.handle for w in Desktop(backend="win32").windows()
         if re.search(r"(?i)Amend a Connection", w.window_text())),
        None,
    )
    amend_win = (amend_app_win32.window(handle=_amend_handle)
                 if _amend_handle else amend_app_win32.top_window())

    # Find fields via their panel auto_ids (stable) then by named alias.
    # pnlHost contains Port (Edit2) and Host (Edit3 / alias 'HostEdit').
    # pnlOutstation contains Outstation (Edit0 / alias 'OutstationEdit').
    try:
        host_field = amend_win.child_window(auto_id="pnlHost").child_window(best_match="HostEdit")
    except Exception:
        host_field = amend_win.child_window(best_match="HostEdit")

    try:
        outstation_field = amend_win.child_window(auto_id="pnlOutstation").child_window(best_match="OutstationEdit")
    except Exception:
        outstation_field = amend_win.child_window(best_match="OutstationEdit")

    return browse_win, amend_win, host_field, outstation_field


def configure_connection(app, new_ip: str, outstation: str):
    """
    Step 3 – Maintain Connections.
    Opens the Amend dialog, reads the current Host and Outstation, shows them,
    asks the user whether to apply changes, then saves and closes (or just
    cancels if the user says no).
    """
    print("\n── Step 3: Maintain Connections ──────────────────────────────")
    browse_win, amend_win, host_field, outstation_field = \
        _open_maintain_connections(app)

    # Read current values using window_text() — reliable for Delphi TEdit.
    try:
        current_host = host_field.window_text().strip()
    except Exception:
        current_host = "<could not read>"
    try:
        current_outstation = outstation_field.window_text().strip()
    except Exception:
        current_outstation = "<could not read>"

    print(f"  Current Host       : {current_host}  →  New: {new_ip}")
    print(f"  Current Outstation : {current_outstation}  →  New: {outstation}")

    no_change = (current_host == new_ip and current_outstation == outstation)
    if no_change:
        print("  No change needed — values already match.")
        _close_window(amend_win)
        _wait_gone(r"Amend a Connection")
        _close_window(browse_win)
        return

    answer = input("\n  Apply these changes? [y/N]: ").strip().lower()
    if answer not in ("y", "yes"):
        print("  Skipped by user.")
        _close_window(amend_win)
        _wait_gone(r"Amend a Connection")
        _close_window(browse_win)
        return

    # Update fields.
    _gui_set_text(host_field,       new_ip,     "Host",       step="Connections")
    _gui_set_text(outstation_field, outstation, "Outstation", step="Connections")

    # Click OK — auto_id="btnOK" confirmed by diagnostic.
    ok_clicked = False
    for search_kwargs in (
        {"auto_id": "btnOK"},
        {"title": "&OK", "class_name": "Button"},
        {"title": "OK",  "class_name": "Button"},
    ):
        try:
            amend_win.child_window(**search_kwargs).click()
            ok_clicked = True
            break
        except Exception:
            pass
    if not ok_clicked:
        send_keys("{ENTER}")
    _wait_gone(r"Amend a Connection")   # exits as soon as dialog closes, max 3s
    print("  Saved.")

    # Close Browse Connections — auto_id="btnCancel" is the "&Close" button
    # (confirmed by inspect_connections.py diagnostic).
    try:
        browse_win.child_window(auto_id="btnCancel").click()
    except Exception:
        _close_window(browse_win)
    time.sleep(ACTION_PAUSE)
    print("  Browse Connections closed.")


# --------------------------------------------------------------------------- #
# Step 4 – Imports > A1140 nonTOU scheme
# --------------------------------------------------------------------------- #

def configure_scheme(app):
    """
    Open the A1140 nonTOU Scheme Editor, clear all page assignments except
    Load Profiling, configure the Load Profiling channels, then Save and OK.
    This step is skipped entirely when --no-scheme is passed.
    """
    print("\n[4/4] Configuring A1140 nonTOU scheme…")
    main_win = app.top_window()
    main_win.set_focus()

    # In dry-run mode, just log what would happen and return.
    if DRY_RUN:
        print("  [DRY RUN] Would open Scheme Editor for A1140 nonTOU.")
        print("            Would clear all page assignments except Load Profiling.")
        print(f"            Would enable Demand Period=30 min,")
        print(f"            channels: {', '.join(LOAD_PROFILE_CHANNELS)}.")
        print("            Would click Save then OK.")
        for page in PAGES_TO_CLEAR:
            _audit("Scheme", f"page:{page}", "<current>", "<None>", applied=False)
        for ch in ALL_CHANNELS:
            desired = ch in LOAD_PROFILE_CHANNELS
            _audit("Scheme", f"channel:{ch}", "<current>", str(desired), applied=False)
        return

    # Step 4 navigation (confirmed from uia diagnostic):
    #   - Tree (auto_id="trvScheme"): top-level nodes are Schemes, Readings, Imports, Trash Can
    #   - Clicking "Imports" populates the right-hand list (auto_id="lvScheme")
    #   - "A1140 nonTOU" is a ListItem in lvScheme — double-clicking it opens Scheme Editor

    # 1. Click "Imports" in the left tree.
    try:
        tree = main_win.child_window(auto_id="trvScheme")
    except Exception:
        tree = main_win.child_window(control_type="Tree")
    try:
        tree.get_item(["Imports"]).click()
    except Exception:
        for item in tree.items():
            if item.text().lower() == "imports":
                item.click()
                break
    time.sleep(ACTION_PAUSE)

    # 2. Double-click "A1140 nonTOU" in the right-hand list.
    try:
        scheme_list = main_win.child_window(auto_id="lvScheme")
    except Exception:
        scheme_list = main_win.child_window(control_type="List")
    opened = False
    for item in scheme_list.items():
        if "a1140" in item.text().lower():
            try:
                item.click(double=True)       # _listview_item API
            except Exception:
                item.select()
                send_keys("{ENTER}{ENTER}")
            opened = True
            break
    if not opened:
        raise RuntimeError("Could not find 'A1140 nonTOU' in the Imports list.")

    time.sleep(DIALOG_WAIT)

    _wait_for_window(r"(?i)Scheme Editor.*A1140", timeout=15)
    _editor_handle = next(
        (w.handle for w in Desktop(backend="win32").windows()
         if re.search(r"(?i)Scheme Editor.*A1140", w.window_text())),
        None,
    )
    if not _editor_handle:
        raise RuntimeError("Scheme Editor window disappeared after detection.")
    editor_win_uia = Application(backend="uia").connect(handle=_editor_handle).window(handle=_editor_handle)

    # Clear every page that should not be assigned.
    for page_name in PAGES_TO_CLEAR:
        print(f"   Clearing: {page_name}")
        _clear_scheme_page(editor_win_uia, page_name)

    # Configure the Load Profiling page.
    print("   Configuring Load Profiling…")
    _configure_load_profiling(editor_win_uia)

    # Save via ToolBar button (auto_id="tlsEditor", text="Save").
    saved = False
    try:
        editor_win_uia.child_window(auto_id="tlsEditor").child_window(
            title="Save", control_type="Button"
        ).click()
        saved = True
    except Exception:
        pass
    if not saved:
        try:
            editor_win_uia.child_window(title="Save", control_type="Button").click()
            saved = True
        except Exception:
            pass
    if not saved:
        send_keys("^s")
    print("   Scheme saved.")

    # Close the editor with OK.
    try:
        editor_win_uia.child_window(auto_id="btnOK").click()
    except Exception:
        try:
            editor_win_uia.child_window(title="OK", control_type="Button").click()
        except Exception:
            send_keys("{ENTER}")
    time.sleep(DIALOG_WAIT)
    print("   Scheme editor closed.")


def _click_scheme_page(editor_win_uia, page_name: str):
    """
    Navigate to a named page in the Scheme Editor by row index in dgvPages
    (confirmed by diagnostic: left panel is a DataGridView, NOT a tree).
    """
    row = PAGE_ROW.get(page_name, -1)
    if row < 0:
        raise RuntimeError(f"Unknown scheme page: {page_name!r}")
    dgv = editor_win_uia.child_window(auto_id="dgvPages")
    dgv.click_input()
    time.sleep(0.2)
    send_keys("^{HOME}")      # row 0 = Summary Page
    time.sleep(0.2)
    for _ in range(row):
        send_keys("{DOWN}")
        time.sleep(0.05)
    time.sleep(ACTION_PAUSE)


def _clear_scheme_page(editor_win_uia, page_name: str):
    """
    Navigate to page_name and set its assignment ComboBox to <None>,
    effectively removing the green tick for that page.
    """
    _click_scheme_page(editor_win_uia, page_name)
    try:
        for combo in editor_win_uia.descendants(control_type="ComboBox"):
            try:
                val = combo.selected_text()
            except Exception:
                val = ""
            # Only act if something is currently assigned (i.e. not already <None>).
            if val and val.strip().lower() not in ("<none>", "none", ""):
                _gui_select(combo, "<None>", f"page assignment: {page_name}", step="Scheme")
    except Exception:
        pass  # Page may already be <None> – safe to ignore.


def _find_checkbox(editor_win_uia, label: str):
    """
    Locate a checkbox in the editor window by its label text.
    Tries the exact label, then label without spaces, then a substring scan.
    Returns None if not found so callers can warn gracefully.
    """
    for title_variant in (label, label.replace(" ", ""), f"&{label}"):
        try:
            return editor_win_uia.child_window(title=title_variant, control_type="CheckBox")
        except Exception:
            continue
    # Substring scan across all checkboxes recursively.
    for cb in editor_win_uia.descendants(control_type="CheckBox"):
        if label.lower() in cb.window_text().lower():
            return cb
    return None


def _enable_led_section(group_pane, label: str):
    """
    Enable a collapsible LED group (cgbChannels / cgbDemandPeriod) if its
    first interactive child is currently disabled.  The LED itself is a Pane
    with auto_id='ledIcon' — clicking it toggles the section on/off.
    """
    try:
        # Sample any direct child that has an 'is_enabled' method.
        for child in group_pane.children():
            ct = child.element_info.control_type
            if ct in ("CheckBox", "ComboBox", "Edit"):
                if not child.is_enabled():
                    print(f"   Enabling '{label}' section…")
                    group_pane.child_window(auto_id="ledIcon").click_input()
                    time.sleep(0.3)
                return  # state confirmed, nothing more to do
    except Exception:
        pass


def _configure_load_profiling(editor_win_uia):
    """
    Configure the Load Profiling page (confirmed structure from diagnostic):
      ECLoadProfileConfigEditor
        pnlChannels → cgbChannels → ledIcon + channel CheckBoxes
        pnlDemandPeriod → cgbDemandPeriod → ledIcon + ComboBox
    """
    _click_scheme_page(editor_win_uia, "Load Profiling")

    lp = editor_win_uia.child_window(auto_id="ECLoadProfileConfigEditor")

    # ── Demand Period ─────────────────────────────────────────────────────────
    demand_pnl = lp.child_window(auto_id="pnlDemandPeriod")
    demand_grp = demand_pnl.child_window(auto_id="cgbDemandPeriod")
    _enable_led_section(demand_grp, "Demand Period")

    # Set the ComboBox to 30 Minutes (only one ComboBox in pnlDemandPeriod).
    try:
        combo = demand_pnl.child_window(control_type="ComboBox")
        items = combo.item_texts()
        target = next((t for t in items if "30" in t), None)
        if target:
            _gui_select(combo, target, "Demand Period", step="Scheme")
        else:
            print(f"   WARNING: no '30' item in Demand Period dropdown: {items}")
    except Exception as exc:
        print(f"   WARNING: Demand Period ComboBox error: {exc}")

    # ── Load Profile Definition (channels) ───────────────────────────────────
    channels_pnl = lp.child_window(auto_id="pnlChannels")
    channels_grp = channels_pnl.child_window(auto_id="cgbChannels")
    _enable_led_section(channels_grp, "Load Profile Definition")

    # Set each channel checkbox to the correct state.
    for ch in ALL_CHANNELS:
        desired = ch in LOAD_PROFILE_CHANNELS
        cb = _find_checkbox(editor_win_uia, ch)
        if cb:
            _gui_checkbox(cb, desired, ch, step="Scheme")
        else:
            print(f"   WARNING: checkbox '{ch}' not found.")

    time.sleep(ACTION_PAUSE)


# --------------------------------------------------------------------------- #
# Pre-step cleanup
# --------------------------------------------------------------------------- #

# Titles of known SMARTset child dialogs that must be closed before
# beginning a new automation step.  Matched as case-insensitive substrings.
_DIALOG_TITLES = [
    "CommsApp Configuration",
    "Browse Connections",
    "Maintain Connections",
    "Amend a Connection",
    "Add a Connection",
    "Scheme Editor",
]


def _close_open_dialogs():
    """
    Close any open SMARTset child dialogs before starting an automation step.
    Called when SMARTset is already running so we start from a clean state.
    Each dialog is dismissed without saving — tries button click, then ESC,
    then Alt+F4 as a last resort.
    """
    import re
    closed_any = False
    for w in Desktop(backend="win32").windows():
        title = w.window_text()
        if any(re.search(kw, title, re.IGNORECASE) for kw in _DIALOG_TITLES):
            print(f"  Closing open dialog: '{title}'")
            dismissed = False

            # Use uia + window(handle=) to target the exact dialog.
            # auto_id="btnCancel" covers both Browse Connections ("&Close")
            # and Amend a Connection ("&Cancel") — confirmed by diagnostic.
            try:
                app_uia = Application(backend="uia").connect(handle=w.handle)
                dlg_uia = app_uia.window(handle=w.handle)
                dlg_uia.child_window(auto_id="btnCancel").click()
                dismissed = True
            except Exception:
                pass

            # Fallback for Delphi dialogs (CommsApp, Scheme Editor): win32 TButton.
            if not dismissed:
                try:
                    app_w32 = Application(backend="win32").connect(handle=w.handle)
                    dlg_w32 = app_w32.window(handle=w.handle)
                    for search_kwargs in (
                        {"title": "Cancel", "class_name": "TButton"},
                        {"title": "Close",  "class_name": "TButton"},
                    ):
                        try:
                            dlg_w32.child_window(**search_kwargs).click()
                            dismissed = True
                            break
                        except Exception:
                            pass
                except Exception:
                    pass

            closed_any = True
            time.sleep(0.5)

    if closed_any:
        time.sleep(ACTION_PAUSE)


# --------------------------------------------------------------------------- #
# Confirmation prompt
# --------------------------------------------------------------------------- #

def confirm(prompt: str) -> bool:
    """Ask the user to confirm before applying changes. Returns True for yes."""
    try:
        answer = input(f"\n{prompt} [y/N]: ").strip().lower()
        return answer in ("y", "yes")
    except EOFError:
        return False


# --------------------------------------------------------------------------- #
# Argument parsing
# --------------------------------------------------------------------------- #

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Automate SMARTset meter scheme configuration.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  # Full run:\n"
            "  python smartset_configure.py --ip 10.0.19.37 --serial 38110126\n\n"
            "  # Steps 1-3 only (skip scheme editor):\n"
            "  python smartset_configure.py --ip 10.0.19.37 --serial 38110126 --no-scheme\n\n"
            "  # Dry run (read + plan, no changes):\n"
            "  python smartset_configure.py --ip 10.0.19.37 --serial 38110126 --dry-run\n\n"
            "  # Skip confirmation prompt:\n"
            "  python smartset_configure.py --ip 10.0.19.37 --serial 38110126 --yes\n"
        ),
    )
    parser.add_argument(
        "--ip",
        required=True,
        help="IP address of the FLAG Communications Server (e.g. 10.0.19.37)",
    )
    parser.add_argument(
        "--serial",
        required=True,
        help="Full meter serial number – last 3 digits become Outstation (e.g. 38110126 → 126)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        default=False,
        help="Read current state and show what would change, but touch nothing.",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        default=False,
        help="Skip the confirmation prompt (for scripted / unattended runs).",
    )
    parser.add_argument(
        "--no-scheme",
        action="store_true",
        default=False,
        help=(
            "Run steps 1-3 only (login, CommsApp IP, connection host/outstation). "
            "Skip the Scheme Editor entirely – safe to use when you only want to "
            "update connectivity settings without touching meter configuration."
        ),
    )
    return parser.parse_args()


# --------------------------------------------------------------------------- #
# Main entry point
# --------------------------------------------------------------------------- #

def _ask_proceed(step_name: str) -> bool:
    """Ask the user whether to proceed to the next step. Returns True to continue."""
    try:
        answer = input(f"\nProceed to {step_name}? [y/N]: ").strip().lower()
        return answer in ("y", "yes")
    except EOFError:
        return False


def main():
    global DRY_RUN
    args       = parse_args()
    DRY_RUN    = args.dry_run
    outstation = args.serial.strip()[-3:]  # last 3 digits of the serial number

    print("=" * 62)
    print("SMARTset Automated Configuration")
    print(f"  IP address  : {args.ip}")
    print(f"  Serial      : {args.serial}  →  Outstation: {outstation}")
    print(f"  Scheme step : {'SKIPPED (--no-scheme)' if args.no_scheme else 'included'}")
    print("=" * 62)

    # ── Step 1 – Launch SMARTset and log in ───────────────────────────
    already_running = _smartset_already_running()
    if already_running:
        print("\n── Step 1: SMARTset already running ──────────────────────────")
        print("  Scheme Manager window is open — skipping launch and login.")
        if not _ask_proceed("Step 2 directly (FLAG Communications Server Setup)"):
            print("Exited at user request.")
            return
        app = _get_main_app()
    else:
        app = launch_and_login()

    # ── Step 2 – CommsApp ─────────────────────────────────────────────
    # Close any leftover SMARTset dialogs before we open new ones.
    _close_open_dialogs()

    if not already_running and not _ask_proceed("Step 2 (FLAG Communications Server Setup)"):
        print("Stopped after Step 1.")
        _write_audit_log(args.serial)
        return

    configure_comms_server(app, args.ip)

    # ── Step 3 – Maintain Connections ─────────────────────────────────
    if not _ask_proceed("Step 3 (Maintain Connections)"):
        print("Stopped after Step 2.")
        _write_audit_log(args.serial)
        return

    configure_connection(app, args.ip, outstation)

    # ── Step 4 – Scheme Editor (skipped with --no-scheme) ─────────────
    if not args.no_scheme:
        if not _ask_proceed("Step 4 (Scheme Editor – A1140 nonTOU)"):
            print("Stopped after Step 3.")
            _write_audit_log(args.serial)
            return
        configure_scheme(app)
    else:
        print("\nScheme configuration skipped (--no-scheme).")

    _write_audit_log(args.serial)
    print("\nAll steps completed successfully.")


if __name__ == "__main__":
    main()
