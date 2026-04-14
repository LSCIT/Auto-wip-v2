"""
test_workflow_e2e.py
End-to-end test of the 3-stage WIP workflow via COM + direct DB verification.

Uses WML Co15 / Div51 / Dec 2025 as the test case (28 jobs, well-validated).
Exercises: load -> batch create -> RFO -> ops done -> OFA -> GAAP done -> AFA

Run from E:\\Auto-Wip\\ with Excel already open:
    python test_workflow_e2e.py
"""

import os
import sys
import time
import threading
import pyodbc
import win32com.client
import win32con
import win32gui

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
WORKBOOK_PATH = r"E:\Auto-Wip\vm\WIPSchedule -Rev 5.68p.xltm"
TEST_CO       = 15       # WML
TEST_MONTH    = "12/1/2025"
TEST_DEPT     = "51"     # 28 jobs, well-validated

DB_CONN_STR = (
    "DRIVER={ODBC Driver 18 for SQL Server};"
    "SERVER=10.103.30.11;DATABASE=LylesWIP;"
    "UID=wip.excel.sql;PWD=WES@2024;"
    "TrustServerCertificate=yes;Encrypt=no"
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def db_query(sql, params=None):
    conn = pyodbc.connect(DB_CONN_STR)
    cursor = conn.cursor()
    cursor.execute(sql, params or ())
    cols = [d[0] for d in cursor.description] if cursor.description else []
    rows = [dict(zip(cols, row)) for row in cursor.fetchall()]
    conn.close()
    return rows

def db_exec(sql, params=None):
    conn = pyodbc.connect(DB_CONN_STR)
    cursor = conn.cursor()
    cursor.execute(sql, params or ())
    conn.commit()
    conn.close()

def dismiss_dialogs(stop_event):
    while not stop_event.is_set():
        time.sleep(0.5)
        try:
            def check(hwnd, _):
                if not win32gui.IsWindowVisible(hwnd): return
                if win32gui.GetClassName(hwnd) != "#32770": return
                def find_ok(child, btns):
                    if win32gui.GetClassName(child) == "Button":
                        txt = win32gui.GetWindowText(child).replace("&", "")
                        if txt in ("OK", "Yes", "End"):
                            btns.append(child)
                btns = []
                win32gui.EnumChildWindows(hwnd, find_ok, btns)
                if btns:
                    win32gui.PostMessage(btns[0], win32con.BM_CLICK, 0, 0)
            win32gui.EnumWindows(check, None)
        except Exception:
            pass

PASS = 0
FAIL = 0

def check(label, condition, detail=""):
    global PASS, FAIL
    if condition:
        PASS += 1
        print(f"  PASS: {label}")
    else:
        FAIL += 1
        print(f"  FAIL: {label}  -- {detail}")


def get_batch_state():
    """Read current batch state from DB."""
    rows = db_query("""
        SELECT BatchState FROM WipBatches
        WHERE JCCo = ? AND WipMonth = '2025-12-01' AND Department = ?
    """, (TEST_CO, TEST_DEPT))
    return rows[0]['BatchState'] if rows else None


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    global PASS, FAIL

    # Clean up prior test batch (leave WipJobData — that's Nicole's data)
    print("--- Cleanup: removing prior test batch for Co15/Dec2025/Dept51 ---")
    db_exec("""
        DELETE FROM WipBatches
        WHERE JCCo = ? AND WipMonth = '2025-12-01' AND Department = ?
    """, (TEST_CO, TEST_DEPT))
    # Reset IsOpsDone/IsGAAPDone flags from any prior test run
    db_exec("""
        UPDATE WipJobData SET IsOpsDone = 0, IsGAAPDone = 0
        WHERE JCCo = ? AND WipMonth = '2025-12-01'
        AND Job IN (SELECT Job FROM WipJobData
                    WHERE JCCo = ? AND WipMonth = '2025-12-01')
    """, (TEST_CO, TEST_CO))
    print("  Done.")

    # Connect to Excel
    print("\n--- Connecting to Excel ---")
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        print("ERROR: Open Excel first, then re-run.")
        sys.exit(1)
    print("  Attached.")

    excel.Visible = True
    excel.DisplayAlerts = False

    stop_event = threading.Event()
    t = threading.Thread(target=dismiss_dialogs, args=(stop_event,), daemon=True)
    t.start()

    try:
        # Reuse workbook if already open; otherwise open fresh
        wb_base = os.path.splitext(os.path.basename(WORKBOOK_PATH))[0]
        wb = None
        for i in range(1, excel.Workbooks.Count + 1):
            if excel.Workbooks(i).Name.startswith(wb_base):
                wb = excel.Workbooks(i)
                print(f"  Reusing: {wb.Name}")
                break
        if wb is None:
            wb = excel.Workbooks.Open(WORKBOOK_PATH)
            print(f"  Opened: {wb.Name}")
        time.sleep(2)

        # =================================================================
        # STAGE 1: Load data
        # =================================================================
        print("\n========== STAGE 1: Accounting Initial Load ==========")

        excel.Run("ResetWorkbook")
        time.sleep(2)

        # MUST disable events AFTER ResetWorkbook (it re-enables them internally)
        excel.EnableEvents = False

        # Unprotect all sheets after ResetWorkbook
        for sn in ["Start", "Jobs-Ops", "Jobs-GAAP", "Settings"]:
            try:
                wb.Sheets(sn).Unprotect("password")
            except Exception:
                pass

        # Set Role to WIPAccounting — required for approval buttons.
        # In production this comes from pnp.WIPSECGetRole on the PNP server,
        # which is not available in the test environment.
        wb.Sheets("Settings").Range("Role").Value = "WIPAccounting"
        print("  Set Role = WIPAccounting on Settings sheet")

        # Set Start sheet values (events OFF to prevent Dept picker popup)
        sh_start = wb.Sheets("Start")
        sh_start.Range("StartCompany").Value = TEST_CO
        sh_start.Range("StartMonth").Value = TEST_MONTH
        sh_start.Range("StartDept").Value = int(TEST_DEPT)
        excel.EnableEvents = True
        print(f"  Set: Co={TEST_CO}, Month={TEST_MONTH}, Dept={TEST_DEPT}")

        print("  Loading Vista data (Jobs-Ops) ...")
        excel.Run("GetWipDetail2", wb.Sheets("Jobs-Ops"))
        print("  Loading Vista data (Jobs-GAAP) ...")
        excel.Run("GetWipDetail2", wb.Sheets("Jobs-GAAP"))

        # Check batch
        state = get_batch_state()
        check("Batch created", state is not None)
        check("Batch state = Open", state == "Open", f"got '{state}'")

        # Check job count on sheet
        sh11 = wb.Sheets("Jobs-Ops")
        # Use UsedRange row count minus header rows as a simpler check
        # (MyJobNos named range has COM access issues with .Cells iteration)
        used_rows = sh11.UsedRange.Rows.Count
        job_count = max(0, used_rows - 6)  # SummaryData starts at row 7
        check("Jobs-Ops has jobs loaded", job_count > 0, f"got ~{job_count} rows")

        # =================================================================
        # Ready for Ops
        # =================================================================
        print("\n  Advancing: Ready for Ops ...")
        excel.Run("RFOYes_Click")
        time.sleep(1)

        state = get_batch_state()
        check("Batch state = ReadyForOps", state == "ReadyForOps", f"got '{state}'")

        # =================================================================
        # STAGE 2: Ops marks rows Done
        # =================================================================
        print("\n========== STAGE 2: Operations Review ==========")

        # Use UpdateAllRows to mark all Jobs-Ops as Done
        print("  Marking all Ops rows Done (UpdateAllRows) ...")
        excel.Run("UpdateAllRows", sh11, "Y")
        time.sleep(3)

        # Verify a few rows saved to DB with IsOpsDone=1
        saved = db_query("""
            SELECT TOP 5 Job, IsOpsDone, OpsRevOverride, OpsRevPlugged, UserName
            FROM WipJobData
            WHERE JCCo = ? AND WipMonth = '2025-12-01' AND IsOpsDone = 1
            ORDER BY Job
        """, (TEST_CO,))
        check("Jobs saved with IsOpsDone=1", len(saved) > 0, f"got {len(saved)} rows")
        for row in saved[:3]:
            print(f"    Job {row['Job'].strip()}: IsOpsDone={row['IsOpsDone']}, "
                  f"OpsRev={row['OpsRevOverride']}, Plugged={row['OpsRevPlugged']}, "
                  f"User={row['UserName']}")

        # Ops Final Approval
        print("\n  Advancing: Ops Final Approval ...")
        excel.Run("OFAYes_Click")
        time.sleep(1)

        state = get_batch_state()
        check("Batch state = OpsApproved", state == "OpsApproved", f"got '{state}'")

        # =================================================================
        # STAGE 3: GAAP marks rows Done
        # =================================================================
        print("\n========== STAGE 3: Accounting Final Approval ==========")

        sh12 = wb.Sheets("Jobs-GAAP")
        print("  Marking all GAAP rows Done (UpdateAllRows) ...")
        excel.Run("UpdateAllRows", sh12, "Y")
        time.sleep(3)

        # Verify GAAP done flags
        gaap_saved = db_query("""
            SELECT COUNT(*) AS cnt FROM WipJobData
            WHERE JCCo = ? AND WipMonth = '2025-12-01' AND IsGAAPDone = 1
        """, (TEST_CO,))
        gaap_count = gaap_saved[0]['cnt'] if gaap_saved else 0
        check("Jobs saved with IsGAAPDone=1", gaap_count > 0, f"got {gaap_count}")

        # Accounting Final Approval
        print("\n  Advancing: Accounting Final Approval ...")
        excel.Run("AFAYes_Click")
        time.sleep(1)

        state = get_batch_state()
        check("Batch state = AcctApproved", state == "AcctApproved", f"got '{state}'")

        # =================================================================
        # Final summary
        # =================================================================
        print("\n========== Final DB State ==========")
        total_ops = db_query("""
            SELECT COUNT(*) AS cnt FROM WipJobData
            WHERE JCCo = ? AND WipMonth = '2025-12-01' AND IsOpsDone = 1
        """, (TEST_CO,))
        total_gaap = db_query("""
            SELECT COUNT(*) AS cnt FROM WipJobData
            WHERE JCCo = ? AND WipMonth = '2025-12-01' AND IsGAAPDone = 1
        """, (TEST_CO,))
        print(f"  Jobs with IsOpsDone=1:  {total_ops[0]['cnt']}")
        print(f"  Jobs with IsGAAPDone=1: {total_gaap[0]['cnt']}")
        print(f"  Batch state: {get_batch_state()}")

        print(f"\n{'='*60}")
        print(f"  RESULTS: {PASS} passed, {FAIL} failed")
        print(f"{'='*60}")

        if FAIL > 0:
            print("\n  ** There were failures — review output above **")

        # Close without saving test state
        excel.EnableEvents = False
        wb.Close(SaveChanges=False)
        excel.EnableEvents = True

    finally:
        stop_event.set()
        print("\nExcel left open. Test complete.")

    sys.exit(1 if FAIL > 0 else 0)


if __name__ == "__main__":
    main()
