"""
Re-import VBA modules into WIPSchedule workbook and reset DB rows.
Steps:
  1. Close all open workbooks (DisplayAlerts=False, EnableEvents=False)
  2. Open WIPSchedule - Rev 5.68p.xltm
  3. Remove & re-import LylesWIPData and Module3
  4. Save as .xltm (FileFormat=53)
  5. Close workbook
  6. Reset DB rows via pyodbc
"""

import sys
import os
import traceback

# ── Paths ──────────────────────────────────────────────────────────────
WORKBOOK_PATH = r"E:\Auto-Wip\vm\WIPSchedule -Rev 5.68p.xltm"
MODULES_TO_IMPORT = {
    "LylesWIPData": r"E:\Auto-Wip\vba_source\LylesWIPData.bas",
    "Module3":      r"E:\Auto-Wip\vba_source\Module3.bas",
}

results = {}

# ══════════════════════════════════════════════════════════════════════
# STEP 1-5: Excel / VBA operations via win32com
# ══════════════════════════════════════════════════════════════════════
try:
    import win32com.client

    # Attach to running Excel
    xl = win32com.client.GetObject(Class="Excel.Application")
    print("[OK] Attached to running Excel instance")

    # Step 1: Close all open workbooks
    try:
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        count = xl.Workbooks.Count
        while xl.Workbooks.Count > 0:
            xl.Workbooks(1).Close(SaveChanges=False)
        print(f"[STEP 1 OK] Closed {count} open workbook(s)")
        results["1_close_workbooks"] = "OK"
    except Exception as e:
        print(f"[STEP 1 FAIL] {e}")
        results["1_close_workbooks"] = f"FAIL: {e}"

    # Step 2: Open the workbook
    try:
        wb = xl.Workbooks.Open(WORKBOOK_PATH)
        print(f"[STEP 2 OK] Opened {wb.Name}")
        results["2_open_workbook"] = "OK"
    except Exception as e:
        print(f"[STEP 2 FAIL] {e}")
        results["2_open_workbook"] = f"FAIL: {e}"
        raise  # can't continue without the workbook

    # Step 3: Remove old modules and import new ones
    try:
        vb_proj = wb.VBProject
        imported = []
        for mod_name, bas_path in MODULES_TO_IMPORT.items():
            # Remove existing module if present
            try:
                old_mod = vb_proj.VBComponents(mod_name)
                vb_proj.VBComponents.Remove(old_mod)
                print(f"  Removed existing module: {mod_name}")
            except Exception:
                print(f"  Module {mod_name} not found (will import fresh)")

            # Import the .bas file
            if not os.path.isfile(bas_path):
                raise FileNotFoundError(f"{bas_path} does not exist")
            vb_proj.VBComponents.Import(bas_path)
            imported.append(mod_name)
            print(f"  Imported: {mod_name} from {bas_path}")

        print(f"[STEP 3 OK] Re-imported modules: {', '.join(imported)}")
        results["3_reimport_modules"] = "OK"
    except Exception as e:
        print(f"[STEP 3 FAIL] {e}")
        results["3_reimport_modules"] = f"FAIL: {e}"

    # Step 4: Save as .xltm (FileFormat=53, ConflictResolution=2)
    try:
        wb.SaveAs(WORKBOOK_PATH, FileFormat=53, ConflictResolution=2)
        print(f"[STEP 4 OK] Saved workbook as .xltm")
        results["4_save_workbook"] = "OK"
    except Exception as e:
        print(f"[STEP 4 FAIL] {e}")
        results["4_save_workbook"] = f"FAIL: {e}"

    # Step 5: Close workbook
    try:
        xl.EnableEvents = False
        wb.Close(SaveChanges=False)
        print(f"[STEP 5 OK] Closed workbook")
        results["5_close_workbook"] = "OK"
    except Exception as e:
        print(f"[STEP 5 FAIL] {e}")
        results["5_close_workbook"] = f"FAIL: {e}"

    # Restore Excel state
    xl.DisplayAlerts = True
    xl.EnableEvents = True

except Exception as e:
    print(f"[EXCEL FATAL] {e}")
    traceback.print_exc()
    results["excel_fatal"] = str(e)

# ══════════════════════════════════════════════════════════════════════
# STEP 6: Reset DB via pyodbc
# ══════════════════════════════════════════════════════════════════════
try:
    import pyodbc

    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 18 for SQL Server};'
        f"SERVER={os.environ.get('LYLESWIP_SERVER', '10.103.30.11')};"
        f"DATABASE={os.environ.get('LYLESWIP_DATABASE', 'LylesWIP')};"
        f"UID={os.environ.get('LYLESWIP_UID', 'wip.excel.sql')};"
        f"PWD={os.environ.get('LYLESWIP_PWD', '')};"
        'TrustServerCertificate=yes;Encrypt=no'
    )
    cursor = conn.cursor()

    # 6a: Reset WipJobData
    cursor.execute(
        "UPDATE WipJobData SET IsOpsDone=0, IsGAAPDone=0, IsClosed=0, "
        "GAAPRevOverride=NULL, GAAPRevPlugged=0, GAAPCostOverride=NULL, "
        "GAAPCostPlugged=0, Source='ExcelImport' "
        "WHERE JCCo=15 AND WipMonth='2025-12-01'"
    )
    print(f"[STEP 6a OK] Reset {cursor.rowcount} job rows (including GAAP overrides)")
    results["6a_reset_wipjobdata"] = f"OK ({cursor.rowcount} rows)"

    # 6b: Delete Dept 51 batch
    cursor.execute(
        "DELETE FROM WipBatches WHERE JCCo=15 AND WipMonth='2025-12-01' AND Department='51'"
    )
    print(f"[STEP 6b OK] Deleted {cursor.rowcount} Dept 51 batch(es)")
    results["6b_delete_batch"] = f"OK ({cursor.rowcount} deleted)"

    # 6c: Clear snapshots
    cursor.execute(
        "DELETE FROM WipYearEndSnapshot WHERE JCCo=15 AND SnapshotYear=2025"
    )
    print(f"[STEP 6c OK] Cleared {cursor.rowcount} snapshot(s)")
    results["6c_clear_snapshots"] = f"OK ({cursor.rowcount} deleted)"

    conn.commit()
    conn.close()
    print("[STEP 6 OK] DB reset complete")
    results["6_db_connection"] = "OK"

except Exception as e:
    print(f"[STEP 6 FAIL] {e}")
    traceback.print_exc()
    results["6_db_reset"] = f"FAIL: {e}"

# ══════════════════════════════════════════════════════════════════════
# Summary
# ══════════════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("SUMMARY")
print("=" * 60)
all_ok = True
for step, status in results.items():
    flag = "OK" if status.startswith("OK") else "FAIL"
    if flag == "FAIL":
        all_ok = False
    print(f"  {step}: {status}")

print("=" * 60)
if all_ok:
    print("ALL STEPS SUCCEEDED")
else:
    print("SOME STEPS FAILED — review output above")
    sys.exit(1)
