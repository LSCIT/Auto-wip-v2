"""
patch_570p.py
Builds Rev 5.70p from a clean 5.68p base by importing updated VBA modules.

Changes in 5.70p (2026-04-06 demo fixes):
  - VistaData.bas: Mth-based date filtering, JB Progress Bills for billing,
                   cost-exists job inclusion, PriorMonthCosts CTE
  - LylesWIPData.bas: recognized profit (not projected) in MergePriorMonthProfits,
                      BuildPriorMonthCostLookup, WriteBackToVista function
  - FormButtons.bas: PushToVistaClick button handler
  - GetWIPDetailData_Modified.bas: unchanged (reads from same recordset)

Run from E:\\Auto-Wip\\:
    python patch_570p.py
"""

import os
import sys
import time
import threading
import win32com.client
import win32con
import win32gui

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
WORKBOOK_PATH  = r"E:\Auto-Wip\vm\WIPSchedule -Rev 5.70p.xltm"
VBA_SOURCE_DIR = r"E:\Auto-Wip\vba_source_569p"

# Modules to re-import (module_name_in_workbook, source_bas_file)
MODULES_TO_REIMPORT = [
    ("VistaData",     "VistaData.bas"),
    ("LylesWIPData",  "LylesWIPData.bas"),
    ("FormButtons",   "FormButtons.bas"),
]

# ---------------------------------------------------------------------------
# Auto-dismiss Excel dialogs
# ---------------------------------------------------------------------------
def dismiss_excel_dialogs(stop_event):
    while not stop_event.is_set():
        time.sleep(0.75)
        try:
            def check_window(hwnd, _ctx):
                if not win32gui.IsWindowVisible(hwnd):
                    return
                if win32gui.GetClassName(hwnd) != "#32770":
                    return
                title = win32gui.GetWindowText(hwnd)
                known = ("Microsoft Excel", "WIP Validation", "LylesWIP Connection Error",
                         "LylesWIP Error", "Vista Connection Error",
                         "LylesWIP Connection Test", "Push to Vista",
                         "Push to Vista Error", "Push to Vista Complete",
                         "GL Period Closed", "Write-Back Error", "Not Yet Approved",
                         "Not a GAAP Quarter")
                if title not in known:
                    return
                def find_button(child, buttons):
                    if win32gui.GetClassName(child) == "Button":
                        txt = win32gui.GetWindowText(child)
                        if txt.replace("&", "") in ("OK", "Yes", "End", "No"):
                            buttons.append(child)
                buttons = []
                win32gui.EnumChildWindows(hwnd, find_button, buttons)
                if buttons:
                    win32gui.PostMessage(buttons[0], win32con.BM_CLICK, 0, 0)
                    print(f"  [auto-dismiss] Clicked OK on '{title}' dialog")
            win32gui.EnumWindows(check_window, None)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Step 1 — Re-import VBA modules
# ---------------------------------------------------------------------------
def reimport_modules(wb):
    print("\n=== Step 1: Re-import VBA modules ===")
    vbp = wb.VBProject

    for module_name, bas_file in MODULES_TO_REIMPORT:
        bas_path = os.path.join(VBA_SOURCE_DIR, bas_file)
        if not os.path.exists(bas_path):
            print(f"  ERROR: {bas_path} not found — skipping")
            continue

        try:
            existing = vbp.VBComponents(module_name)
            vbp.VBComponents.Remove(existing)
            print(f"  Removed existing {module_name}")
        except Exception:
            print(f"  {module_name} not in workbook — importing fresh")

        vbp.VBComponents.Import(bas_path)
        print(f"  Imported {bas_file}")

    print("Module import complete.")


# ---------------------------------------------------------------------------
# Step 2 — Create "Push to Vista" button on Start sheet
# ---------------------------------------------------------------------------
def create_push_to_vista_button(wb):
    print("\n=== Step 2: Create 'Push to Vista' button ===")
    ws = wb.Sheets("Start")
    ws.Activate()
    try:
        ws.Unprotect("password")
    except Exception:
        pass  # May not be protected

    # Check if button already exists
    for shp in ws.Shapes:
        if shp.Name == "btnPushToVista":
            print("  Button 'btnPushToVista' already exists — removing and recreating.")
            shp.Delete()
            break

    # Find the Save & Distribute button for reference positioning
    ref_shape = None
    for shp in ws.Shapes:
        if "SaveDistribute" in (shp.OnAction or ""):
            ref_shape = shp
            break
        if "Distribute" in shp.Name:
            ref_shape = shp
            break

    if ref_shape:
        btn_left = float(ref_shape.Left)
        btn_top = float(ref_shape.Top) + float(ref_shape.Height) + 12.0
        btn_width = max(float(ref_shape.Width), 140.0)
        btn_height = max(float(ref_shape.Height), 28.0)
        # Clamp to reasonable range
        btn_left = max(10.0, min(btn_left, 700.0))
        btn_top = max(10.0, min(btn_top, 800.0))
        btn_width = min(btn_width, 300.0)
        btn_height = min(btn_height, 50.0)
        print(f"  Positioning relative to '{ref_shape.Name}' ({btn_left}, {btn_top}, {btn_width}x{btn_height})")
    else:
        btn_left = 300.0
        btn_top = 400.0
        btn_width = 180.0
        btn_height = 32.0
        print("  No reference button found — using default position")

    # msoShapeRoundedRectangle = 5
    shp = ws.Shapes.AddShape(5, btn_left, btn_top, btn_width, btn_height)
    shp.Name = "btnPushToVista"

    # Deep navy blue background
    shp.Fill.ForeColor.RGB = (100 << 16) | (56 << 8) | 31

    # White text, centered
    tf = shp.TextFrame2
    tf.TextRange.Text = "Push to Vista"
    tf.TextRange.Font.Size = 11
    tf.TextRange.Font.Bold = True
    tf.TextRange.Font.Fill.ForeColor.RGB = (255 << 16) | (255 << 8) | 255
    tf.TextRange.ParagraphFormat.Alignment = 2  # msoAlignCenter
    tf.VerticalAnchor = 3   # msoAnchorMiddle
    tf.MarginLeft = 4
    tf.MarginRight = 4

    # Rounded corners — use SetProperty for COM indexed property
    try:
        shp.Adjustments.SetItem(1, 0.25)
    except Exception:
        pass  # Some shapes don't support adjustments

    # Subtle shadow
    shp.Shadow.Type = 1
    shp.Shadow.ForeColor.RGB = (180 << 16) | (180 << 8) | 180
    shp.Shadow.Transparency = 0.6
    shp.Shadow.OffsetX = 2
    shp.Shadow.OffsetY = 2

    # Border
    shp.Line.ForeColor.RGB = (80 << 16) | (40 << 8) | 20
    shp.Line.Weight = 1

    # Assign macro
    shp.OnAction = "FormButtons.PushToVistaClick"

    print(f"  Created button at ({btn_left}, {btn_top}), {btn_width}x{btn_height}")
    print("  OnAction = FormButtons.PushToVistaClick")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    if not os.path.exists(WORKBOOK_PATH):
        print(f"ERROR: Workbook not found: {WORKBOOK_PATH}")
        sys.exit(1)

    print(f"Opening {os.path.basename(WORKBOOK_PATH)} ...")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False

    stop_event = threading.Event()
    dismiss_thread = threading.Thread(
        target=dismiss_excel_dialogs, args=(stop_event,), daemon=True
    )
    dismiss_thread.start()

    try:
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        time.sleep(2)

        reimport_modules(wb)
        create_push_to_vista_button(wb)

        print("\nSaving workbook ...")
        # xlOpenXMLTemplateMacroEnabled = 53 (.xltm)
        wb.SaveAs(WORKBOOK_PATH, FileFormat=53)
        print("Saved.")

        print("\n=== Rev 5.70p Build Complete ===")
        print(f"Workbook: {os.path.basename(WORKBOOK_PATH)}")
        print("Changes:")
        print("  - VistaData.bas: Mth-based filters, JB Progress Bills, cost-exists inclusion")
        print("  - LylesWIPData.bas: recognized profit, WriteBackToVista")
        print("  - FormButtons.bas: PushToVistaClick handler")
        print("  - 'Push to Vista' button on Start sheet")
        print("\nLeaving workbook open for inspection.")

    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
    finally:
        stop_event.set()
        excel.DisplayAlerts = True


if __name__ == "__main__":
    main()
