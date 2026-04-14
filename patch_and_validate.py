"""
patch_and_validate.py
Automates the full D2 validation cycle via Excel COM (pywin32):
  1. Re-imports LylesWIPData.bas and BatchValidate.bas into the active workbook
  2. Saves the workbook
  3. Runs BatchValidateAll (24 combos)
  4. Copies updated snapshots to repo vm/validate-d3/
  5. Runs validate_wip.py and prints results

Run from E:\\Auto-Wip\\:
    python patch_and_validate.py

Flags:
    --import-only    Re-import modules and save, but skip BatchValidateAll
    --run-only       Skip module import, go straight to BatchValidateAll
    --copy-only      Just copy C:\\Trusted\\validate-d3\\ to vm\\validate-d3\\
    --validate-only  Just run validate_wip.py
"""

import argparse
import os
import shutil
import subprocess
import sys
import threading
import time

import win32com.client
import win32con
import win32gui

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
WORKBOOK_PATH   = r"E:\Auto-Wip\vm\WIPSchedule -Rev 5.73p.xltm"
VBA_SOURCE_DIR  = r"E:\Auto-Wip\vba_source"
SNAPSHOT_SOURCE = r"C:\Trusted\validate-d3"
SNAPSHOT_DEST   = r"E:\Auto-Wip\vm\validate-d3"
VALIDATE_SCRIPT = r"E:\Auto-Wip\validate_wip.py"

# Modules to re-import (module_name_in_workbook, source_bas_file)
MODULES_TO_REIMPORT = [
    ("LylesWIPData", "LylesWIPData.bas"),
    ("VistaData", "VistaData.bas"),
    ("BatchValidate", "BatchValidate.bas"),
]

# ---------------------------------------------------------------------------
# Background thread: auto-dismiss Excel MsgBox dialogs
# ---------------------------------------------------------------------------
def dismiss_excel_dialogs(stop_event):
    """
    Watches for Excel MsgBox dialogs and clicks OK/Yes automatically.
    Needed because BatchValidateAll ends with MsgBox which blocks excel.Run().
    """
    while not stop_event.is_set():
        time.sleep(0.75)
        try:
            def check_window(hwnd, _ctx):
                if not win32gui.IsWindowVisible(hwnd):
                    return
                if win32gui.GetClassName(hwnd) != "#32770":   # dialog class
                    return
                title = win32gui.GetWindowText(hwnd)
                if title not in ("Microsoft Excel", "WIP Validation", "LylesWIP Connection Error",
                                 "LylesWIP Error", "Vista Connection Error",
                                 "LylesWIP Connection Test"):
                    return
                # Enumerate child buttons and click the first OK/Yes
                def find_button(child, buttons):
                    if win32gui.GetClassName(child) == "Button":
                        txt = win32gui.GetWindowText(child)
                        if txt.replace("&", "") in ("OK", "Yes", "End"):
                            buttons.append(child)
                buttons = []
                win32gui.EnumChildWindows(hwnd, find_button, buttons)
                if buttons:
                    win32gui.PostMessage(buttons[0], win32con.BM_CLICK, 0, 0)
                    print(f"  [auto-dismiss] Clicked OK on '{title}' dialog")

            win32gui.EnumWindows(check_window, None)
        except Exception:
            pass   # EnumWindows can raise if a window is destroyed mid-enum


# ---------------------------------------------------------------------------
# Step 1 — Re-import VBA modules
# ---------------------------------------------------------------------------
def reimport_modules(wb):
    vbp = wb.VBProject
    # If VBProject is inaccessible, COM raises an error with a specific message
    # about "Programmatic access to Visual Basic Project is not trusted"
    for module_name, bas_file in MODULES_TO_REIMPORT:
        bas_path = os.path.join(VBA_SOURCE_DIR, bas_file)
        if not os.path.exists(bas_path):
            print(f"  ERROR: {bas_path} not found — skipping {module_name}")
            continue

        # Remove existing module
        try:
            existing = vbp.VBComponents(module_name)
            vbp.VBComponents.Remove(existing)
            print(f"  Removed existing {module_name}")
        except Exception:
            print(f"  {module_name} not in workbook — importing fresh")

        # Import new version
        vbp.VBComponents.Import(bas_path)
        print(f"  Imported {bas_file}")

    print("Module import complete.")


# ---------------------------------------------------------------------------
# Step 2 — Run BatchValidateAll
# ---------------------------------------------------------------------------
def run_batch_validate(excel):
    os.makedirs(SNAPSHOT_SOURCE, exist_ok=True)

    stop_event = threading.Event()
    dismiss_thread = threading.Thread(
        target=dismiss_excel_dialogs, args=(stop_event,), daemon=True
    )
    dismiss_thread.start()

    print("Running BatchValidateAll (24 combos — allow 5-15 min) ...")
    start = time.time()
    try:
        excel.Run("BatchValidate.BatchValidateAll")
    finally:
        stop_event.set()

    elapsed = time.time() - start
    print(f"BatchValidateAll finished in {elapsed:.0f}s")


# ---------------------------------------------------------------------------
# Step 3 — Copy snapshots to repo
# ---------------------------------------------------------------------------
def copy_snapshots():
    if not os.path.isdir(SNAPSHOT_SOURCE):
        print(f"  Snapshot source not found: {SNAPSHOT_SOURCE}")
        return 0

    copied = 0
    for fname in sorted(os.listdir(SNAPSHOT_SOURCE)):
        if not fname.endswith(".xltm"):
            continue
        src = os.path.join(SNAPSHOT_SOURCE, fname)
        dst = os.path.join(SNAPSHOT_DEST, fname)
        shutil.copy2(src, dst)
        print(f"  Copied {fname}")
        copied += 1

    print(f"Copied {copied} snapshot(s) to {SNAPSHOT_DEST}")
    return copied


# ---------------------------------------------------------------------------
# Step 4 — Run validate_wip.py
# ---------------------------------------------------------------------------
def run_validator():
    env = os.environ.copy()
    env["WIP_SNAPSHOT_DIR"] = SNAPSHOT_SOURCE      # point at C:\Trusted\validate-d3
    env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(
        [sys.executable, VALIDATE_SCRIPT],
        cwd=os.path.dirname(VALIDATE_SCRIPT),
        env=env,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return result.returncode


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Patch workbook + run validation")
    parser.add_argument("--import-only", action="store_true")
    parser.add_argument("--run-only",    action="store_true")
    parser.add_argument("--copy-only",   action="store_true")
    parser.add_argument("--validate-only", action="store_true")
    args = parser.parse_args()

    do_import   = not any([args.run_only, args.copy_only, args.validate_only])
    do_run      = not any([args.import_only, args.copy_only, args.validate_only])
    do_copy     = not any([args.import_only, args.run_only, args.validate_only])
    do_validate = not any([args.import_only, args.run_only, args.copy_only])

    if args.import_only:  do_import   = True
    if args.run_only:     do_run      = True
    if args.copy_only:    do_copy     = True
    if args.validate_only: do_validate = True

    # ---- Excel COM needed for import and/or run ----
    if do_import or do_run:
        if not os.path.exists(WORKBOOK_PATH):
            print(f"ERROR: Workbook not found: {WORKBOOK_PATH}")
            sys.exit(1)

        print("Connecting to Excel COM ...")
        excel = None
        we_launched_excel = False

        # Try to attach to an already-running Excel instance first.
        # On some Windows / M365 installs, CoCreateInstance fails from a
        # non-interactive session but GetActiveObject works if Excel is open.
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            print("  Attached to existing Excel instance.")
        except Exception:
            pass

        if excel is None:
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                we_launched_excel = True
                print("  Launched new Excel instance.")
            except Exception as e:
                print(f"\nERROR: Could not start Excel COM server: {e}")
                print("\nFix options:")
                print("  1. Open Excel manually first, then re-run this script")
                print("     (script will attach to the running instance)")
                print("  2. Run this script from the same user account Excel runs under")
                sys.exit(1)

        excel.Visible = True        # visible so you can see BatchValidate progress
        excel.DisplayAlerts = False

        try:
            print(f"Opening {os.path.basename(WORKBOOK_PATH)} ...")
            wb = excel.Workbooks.Open(WORKBOOK_PATH)

            if do_import:
                print("\n--- Step 1: Re-import VBA modules ---")
                try:
                    reimport_modules(wb)
                except Exception as e:
                    if "Programmatic access" in str(e) or "800a03ec" in str(e).lower():
                        print("\nERROR: VBA project access is blocked.")
                        print("Fix: Excel -> File -> Options -> Trust Center -> Trust Center Settings")
                        print("       -> Macro Settings -> check 'Trust access to the VBA project object model'")
                        excel.Quit()
                        sys.exit(1)
                    raise
                # .xltm = xlOpenXMLTemplateMacroEnabled (FileFormat 53)
                excel.DisplayAlerts = False
                wb.SaveAs(WORKBOOK_PATH, FileFormat=53)
                print("Workbook saved (.xltm).")

            if do_run:
                print("\n--- Step 2: Run BatchValidateAll ---")
                run_batch_validate(excel)
                excel.DisplayAlerts = False
                wb.SaveAs(WORKBOOK_PATH, FileFormat=53)

            excel.EnableEvents = False
            wb.Close(SaveChanges=False)
            excel.EnableEvents = True

        finally:
            if we_launched_excel:
                excel.Quit()
                print("Excel closed.")
            else:
                print("Excel left open (was already running).")

    if do_copy:
        print("\n--- Step 3: Copy snapshots to repo ---")
        copy_snapshots()

    if do_validate:
        print("\n--- Step 4: Run validate_wip.py ---")
        rc = run_validator()
        sys.exit(rc)


if __name__ == "__main__":
    main()
