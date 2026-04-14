"""
Re-import VBA modules into WIPSchedule workbook and add Distribute button.
Requires Excel to be running. Uses win32com (COM automation).
"""

import sys, os, time

try:
    import win32com.client
    from pywintypes import com_error
except ImportError:
    print("FATAL: pywin32 not installed. Run: pip install pywin32")
    sys.exit(1)

# ---------- paths ----------
WB_PATH   = r"E:\Auto-Wip\vm\WIPSchedule -Rev 5.68p.xltm"
BAS_FILES = {
    "LylesWIPData": r"E:\Auto-Wip\vba_source\LylesWIPData.bas",
    "FormButtons":  r"E:\Auto-Wip\vba_source\FormButtons.bas",
}
SHEET_PASSWORD = os.environ.get('WIP_SHEET_PASSWORD', '')

results = {}

# ---------- helpers ----------
def report(step, ok, msg=""):
    tag = "OK" if ok else "FAIL"
    results[step] = (ok, msg)
    print(f"[{tag}] Step {step}: {msg}")

# ---------- connect to Excel ----------
try:
    xl = win32com.client.GetActiveObject("Excel.Application")
    print("Connected to running Excel instance.")
except com_error:
    print("FATAL: Could not connect to a running Excel instance.")
    sys.exit(1)

xl.DisplayAlerts = False
xl.EnableEvents = False

# ===== STEP 1 — Close all open workbooks =====
try:
    count = xl.Workbooks.Count
    if count == 0:
        report(1, True, "No workbooks were open.")
    else:
        names = [xl.Workbooks(i+1).Name for i in range(count)]
        for name in names:
            try:
                xl.Workbooks(name).Close(SaveChanges=False)
            except com_error as e:
                pass  # already closed
        report(1, True, f"Closed {count} workbook(s): {', '.join(names)}")
except Exception as e:
    report(1, False, str(e))

# ===== STEP 2 — Open workbook =====
wb = None
try:
    wb = xl.Workbooks.Open(WB_PATH)
    report(2, True, f"Opened {wb.Name}")
except Exception as e:
    report(2, False, str(e))
    xl.EnableEvents = True
    sys.exit(1)

# ===== STEP 3 — Re-import standard modules =====
vbp = wb.VBProject
for mod_name, bas_path in BAS_FILES.items():
    try:
        # Remove existing module if present
        removed = False
        for comp in vbp.VBComponents:
            if comp.Name == mod_name:
                vbp.VBComponents.Remove(comp)
                removed = True
                break
        # Import fresh
        vbp.VBComponents.Import(bas_path)
        status = "replaced" if removed else "added (was not present)"
        report(f"3-{mod_name}", True, f"{mod_name}: {status}")
    except Exception as e:
        report(f"3-{mod_name}", False, f"{mod_name}: {e}")

# ===== STEP 4 — Add button on Start sheet (Sheet17) =====
btn_info = "not placed"
try:
    sh17 = wb.Sheets("Start")
    sh17.Unprotect(SHEET_PASSWORD)

    # Add a Forms button (Buttons collection = Forms toolbar buttons)
    # Position: Left=350, Top=280 puts it roughly around row 15-16, col E-F area
    btn = sh17.Buttons().Add(Left=350, Top=280, Width=130, Height=30)
    btn.Caption = "Save & Distribute"
    btn.OnAction = "SaveDistributeClick"
    btn.Font.Size = 10
    btn.Font.Bold = True
    btn_info = f"Left={btn.Left}, Top={btn.Top}, Width={btn.Width}, Height={btn.Height}"

    sh17.Protect(SHEET_PASSWORD)
    report(4, True, f"Button added on Start sheet at {btn_info}")
except com_error as e:
    # Fallback: try OLEObjects (ActiveX button)
    try:
        print("  Buttons.Add failed, trying OLEObjects.Add (ActiveX CommandButton)...")
        ole = sh17.OLEObjects().Add(ClassType="Forms.CommandButton.1",
                                     Left=350, Top=280, Width=130, Height=30)
        ole.Object.Caption = "Save & Distribute"
        # For ActiveX, OnAction doesn't work the same way; we need Click event code
        # Wire it via the sheet module
        sheet_module = None
        for comp in vbp.VBComponents:
            if comp.Name == sh17.CodeName:
                sheet_module = comp
                break
        if sheet_module:
            code = (
                "\nPrivate Sub " + ole.Name + "_Click()\n"
                "    SaveDistributeClick\n"
                "End Sub\n"
            )
            sheet_module.CodeModule.AddFromString(code)
        btn_info = f"ActiveX button, Left=350, Top=280"
        sh17.Protect(SHEET_PASSWORD)
        report(4, True, f"ActiveX button added on Start sheet at {btn_info}")
    except Exception as e2:
        # Second fallback: Shape with macro
        try:
            print("  OLEObjects also failed, trying Shapes.AddShape...")
            shp = sh17.Shapes.AddShape(1, 350, 280, 130, 30)  # msoShapeRectangle=1
            shp.TextFrame.Characters().Text = "Save & Distribute"
            shp.TextFrame.Characters().Font.Size = 10
            shp.TextFrame.Characters().Font.Bold = True
            shp.OnAction = "SaveDistributeClick"
            btn_info = f"Shape-button, Left=350, Top=280"
            sh17.Protect(SHEET_PASSWORD)
            report(4, True, f"Shape button added on Start sheet at {btn_info}")
        except Exception as e3:
            report(4, False, f"All button approaches failed. Last error: {e3}")

# ===== STEP 5 — Save as .xltm (FileFormat=53) =====
try:
    wb.SaveAs(WB_PATH, FileFormat=53, ConflictResolution=2)
    report(5, True, f"Saved as .xltm to {WB_PATH}")
except Exception as e:
    report(5, False, str(e))

# ===== STEP 6 — Close workbook =====
try:
    xl.EnableEvents = False
    wb.Close(SaveChanges=False)
    report(6, True, "Workbook closed (EnableEvents=False)")
except Exception as e:
    report(6, False, str(e))
finally:
    xl.EnableEvents = True
    xl.DisplayAlerts = True

# ---------- Summary ----------
print("\n" + "="*60)
print("SUMMARY")
print("="*60)
all_ok = True
for step, (ok, msg) in sorted(results.items(), key=lambda x: str(x[0])):
    tag = "OK  " if ok else "FAIL"
    print(f"  [{tag}] {step}: {msg}")
    if not ok:
        all_ok = False

if all_ok:
    print("\nAll steps completed successfully.")
else:
    print("\nSome steps failed — review above.")
