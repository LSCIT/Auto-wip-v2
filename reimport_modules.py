"""
reimport_modules.py
Re-imports VBA modules into the active WIP workbook via COM.
"""

import os
import sys
import time
import threading
import win32com.client
import win32con
import win32gui

WORKBOOK_PATH = r"E:\Auto-Wip\vm\WIPSchedule -Rev 5.68p.xltm"
VBA_SOURCE = r"E:\Auto-Wip\vba_source"

# Document modules: replace code in-place (cannot use Import)
DOCUMENT_MODULES = [
    ("Sheet11", "Sheet11.cls"),
    ("Sheet12", "Sheet12.cls"),
]

# Standard modules: remove + import
STANDARD_MODULES = [
    ("LylesWIPData", "LylesWIPData.bas"),
    ("FormButtons", "FormButtons.bas"),
    ("Module6", "Module6_Phase1_Complete.bas"),
    ("Permissions", "Permissions_Modified.bas"),
]


def dismiss_dialogs(stop_event):
    """Auto-dismiss Excel dialogs in background."""
    while not stop_event.is_set():
        time.sleep(0.5)
        try:
            def check(hwnd, _):
                if not win32gui.IsWindowVisible(hwnd):
                    return
                if win32gui.GetClassName(hwnd) != "#32770":
                    return
                def find_btn(child, btns):
                    if win32gui.GetClassName(child) == "Button":
                        txt = win32gui.GetWindowText(child).replace("&", "")
                        if txt in ("OK", "Yes", "End"):
                            btns.append(child)
                btns = []
                win32gui.EnumChildWindows(hwnd, find_btn, btns)
                if btns:
                    win32gui.PostMessage(btns[0], win32con.BM_CLICK, 0, 0)
                    title = win32gui.GetWindowText(hwnd)
                    print(f"  [auto-dismiss] Clicked button on '{title}' dialog")
            win32gui.EnumWindows(check, None)
        except Exception:
            pass


def replace_document_module(vbp, comp_name, cls_file):
    """Replace code in a document module (Sheet code-behind)."""
    cls_path = os.path.join(VBA_SOURCE, cls_file)
    if not os.path.exists(cls_path):
        print(f"  ERROR: {cls_path} not found -- skipping {comp_name}")
        return False

    try:
        sheet_comp = vbp.VBComponents(comp_name)
    except Exception as e:
        print(f"  ERROR: Component '{comp_name}' not found in workbook: {e}")
        return False

    # Clear existing code
    cm = sheet_comp.CodeModule
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)

    # Read .cls file, skip Attribute lines and leading blanks
    with open(cls_path, "r", encoding="utf-8", errors="replace") as f:
        lines = f.readlines()

    code_start = 0
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("Attribute") or stripped == "":
            code_start = i + 1
        else:
            break

    code = "".join(lines[code_start:])
    cm.AddFromString(code)
    print(f"  Replaced code in {comp_name} from {cls_file} ({len(lines) - code_start} lines)")
    return True


def reimport_standard_module(vbp, module_name, bas_file):
    """Remove and re-import a standard .bas module."""
    bas_path = os.path.join(VBA_SOURCE, bas_file)
    if not os.path.exists(bas_path):
        print(f"  ERROR: {bas_path} not found -- skipping {module_name}")
        return False

    # Remove existing
    try:
        existing = vbp.VBComponents(module_name)
        vbp.VBComponents.Remove(existing)
        print(f"  Removed existing {module_name}")
    except Exception:
        print(f"  {module_name} not in workbook -- importing fresh")

    # Import
    vbp.VBComponents.Import(bas_path)
    print(f"  Imported {bas_file} (as {module_name})")
    return True


def main():
    print("=== Re-import VBA modules into workbook ===\n")

    # Connect to running Excel
    print("Connecting to Excel...")
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        print("  Attached to running Excel instance.")
    except Exception:
        print("ERROR: Excel is not running. Open Excel first.")
        sys.exit(1)

    excel.Visible = True
    excel.DisplayAlerts = False
    excel.EnableEvents = False

    stop_event = threading.Event()
    t = threading.Thread(target=dismiss_dialogs, args=(stop_event,), daemon=True)
    t.start()

    errors = []

    try:
        # Close all open workbooks
        print("\nClosing all open workbooks...")
        while excel.Workbooks.Count > 0:
            wb_name = excel.Workbooks(1).Name
            excel.Workbooks(1).Close(SaveChanges=False)
            print(f"  Closed: {wb_name}")
        print("  All workbooks closed.")

        # Open the workbook fresh
        print(f"\nOpening {os.path.basename(WORKBOOK_PATH)}...")
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        print(f"  Opened: {wb.Name}")
        time.sleep(2)

        vbp = wb.VBProject

        # Import document modules (Sheet11, Sheet12)
        print("\n--- Document modules (code replacement) ---")
        for comp_name, cls_file in DOCUMENT_MODULES:
            try:
                if not replace_document_module(vbp, comp_name, cls_file):
                    errors.append(f"Failed: {comp_name}")
            except Exception as e:
                print(f"  ERROR on {comp_name}: {e}")
                errors.append(f"Exception: {comp_name}: {e}")

        # Import standard modules
        print("\n--- Standard modules (remove + import) ---")
        for module_name, bas_file in STANDARD_MODULES:
            try:
                if not reimport_standard_module(vbp, module_name, bas_file):
                    errors.append(f"Failed: {module_name}")
            except Exception as e:
                print(f"  ERROR on {module_name}: {e}")
                errors.append(f"Exception: {module_name}: {e}")

        # Save as .xltm (FileFormat=53)
        print(f"\nSaving workbook as .xltm...")
        excel.DisplayAlerts = False
        wb.SaveAs(WORKBOOK_PATH, FileFormat=53)
        print("  Saved.")

        # Close workbook
        print("Closing workbook...")
        excel.EnableEvents = False
        wb.Close(SaveChanges=False)
        excel.EnableEvents = True
        print("  Closed.")

    except Exception as e:
        errors.append(f"Top-level exception: {e}")
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
    finally:
        stop_event.set()
        excel.EnableEvents = True
        excel.DisplayAlerts = True

    # Summary
    print(f"\n{'='*50}")
    if errors:
        print(f"COMPLETED WITH {len(errors)} ERROR(S):")
        for err in errors:
            print(f"  - {err}")
        sys.exit(1)
    else:
        print("ALL MODULES IMPORTED SUCCESSFULLY")
        modules_list = [c for c, _ in DOCUMENT_MODULES] + [m for m, _ in STANDARD_MODULES]
        print(f"  Imported: {', '.join(modules_list)}")
        sys.exit(0)


if __name__ == "__main__":
    main()
