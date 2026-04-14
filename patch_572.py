"""
patch_572.py
Creates Rev 5.72p from 5.71p by importing the updated VBA modules.

Changes in 5.72p (Nicole feedback 2026-04-08):
  1. VistaData.bas — ContractStatus mapped to 1/2, HardClosed filter, zero-activity exclusion
  2. GetWIPDetailData_Modified.bas — Override column defaulting, bonus calculation

Run from E:\Auto-Wip\:
    python patch_572.py
"""

import os
import shutil
import sys
import time

import win32com.client

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SOURCE_WORKBOOK = r"E:\Auto-Wip\vm\WIPSchedule -Rev 5.71p.xltm"
TARGET_WORKBOOK = r"E:\Auto-Wip\vm\WIPSchedule -Rev 5.72p.xltm"
VBA_SOURCE_DIR  = r"E:\Auto-Wip\vba_source"

# Modules to re-import (module_name_in_workbook, source_bas_file)
MODULES_TO_REIMPORT = [
    ("VistaData",         "VistaData.bas"),
    ("GetWIPDetailData",  "GetWIPDetailData_Modified.bas"),
]


def main():
    # Copy 5.71p → 5.72p
    if not os.path.exists(SOURCE_WORKBOOK):
        print(f"ERROR: Source workbook not found: {SOURCE_WORKBOOK}")
        sys.exit(1)

    if os.path.exists(TARGET_WORKBOOK):
        print(f"Target already exists: {TARGET_WORKBOOK}")
        resp = input("Overwrite? [y/N] ").strip().lower()
        if resp != "y":
            print("Aborted.")
            sys.exit(0)

    print(f"Copying {os.path.basename(SOURCE_WORKBOOK)} -> {os.path.basename(TARGET_WORKBOOK)} ...")
    shutil.copy2(SOURCE_WORKBOOK, TARGET_WORKBOOK)

    # Connect to Excel
    print("Connecting to Excel COM ...")
    excel = None
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        print("  Attached to existing Excel instance.")
    except Exception:
        pass

    if excel is None:
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            print("  Launched new Excel instance.")
        except Exception as e:
            print(f"\nERROR: Could not start Excel COM server: {e}")
            sys.exit(1)

    excel.Visible = True
    excel.DisplayAlerts = False

    try:
        print(f"Opening {os.path.basename(TARGET_WORKBOOK)} ...")
        wb = excel.Workbooks.Open(TARGET_WORKBOOK)

        # Re-import modules
        vbp = wb.VBProject
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
            print(f"  Imported {bas_file} as {module_name}")

        # Save
        wb.Save()
        print(f"\nSaved {os.path.basename(TARGET_WORKBOOK)}")
        print("\nDone! The workbook is open in Excel.")
        print("To test: load Co 15, Div 51, Dec 2025 from the Start sheet.")

    except Exception as e:
        print(f"\nERROR: {e}")
        if "Programmatic access" in str(e) or "800a03ec" in str(e).lower():
            print("\nFix: Excel -> File -> Options -> Trust Center -> Trust Center Settings")
            print("       -> Macro Settings -> check 'Trust access to the VBA project object model'")
        sys.exit(1)


if __name__ == "__main__":
    main()
