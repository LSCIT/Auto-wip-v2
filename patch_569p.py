"""
patch_569p.py
Imports updated VBA modules into WIPSchedule -Rev 5.69p.xltm and creates
a "Push to Vista" button on the Start sheet (Sheet17).

Run from E:\\Auto-Wip\\:
    python patch_569p.py
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
WORKBOOK_PATH  = r"E:\Auto-Wip\vm\WIPSchedule -Rev 5.69p.xltm"
VBA_SOURCE_DIR = r"E:\Auto-Wip\vba_source_569p"

# Modules to re-import (module_name_in_workbook, source_bas_file)
MODULES_TO_REIMPORT = [
    ("LylesWIPData", "LylesWIPData.bas"),
    ("FormButtons",  "FormButtons.bas"),
]

# ---------------------------------------------------------------------------
# Auto-dismiss Excel dialogs (reused from patch_and_validate.py)
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
                if title not in ("Microsoft Excel", "WIP Validation", "LylesWIP Connection Error",
                                 "LylesWIP Error", "Vista Connection Error",
                                 "LylesWIP Connection Test", "Push to Vista",
                                 "Push to Vista Error", "Push to Vista Complete",
                                 "GL Period Closed", "Write-Back Error", "Not Yet Approved",
                                 "Not a GAAP Quarter"):
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

        # Remove existing module
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
# Step 2 — Create "Push to Vista" button on Start sheet (Sheet17)
# ---------------------------------------------------------------------------
def create_push_to_vista_button(wb):
    print("\n=== Step 2: Create 'Push to Vista' button ===")
    ws = wb.Sheets("Start")

    # Check if button already exists
    for shp in ws.Shapes:
        if shp.Name == "btnPushToVista":
            print("  Button 'btnPushToVista' already exists — removing and recreating.")
            shp.Delete()
            break

    # Find the Save & Distribute button for reference positioning
    ref_shape = None
    for shp in ws.Shapes:
        if "SaveDistribute" in shp.OnAction or "Distribute" in shp.Name:
            ref_shape = shp
            break

    if ref_shape:
        # Place below the Save & Distribute button with some spacing
        btn_left = ref_shape.Left
        btn_top = ref_shape.Top + ref_shape.Height + 12
        btn_width = ref_shape.Width
        btn_height = ref_shape.Height
        print(f"  Positioning relative to '{ref_shape.Name}' (top={ref_shape.Top}, h={ref_shape.Height})")
    else:
        # Fallback: place in a reasonable default spot on the Start sheet
        # AFA buttons are typically around row 25-30, so put this below
        btn_left = 400
        btn_top = 520
        btn_width = 200
        btn_height = 36
        print("  No reference button found — using default position")

    # Create the button using a rounded rectangle shape
    # msoShapeRoundedRectangle = 5
    shp = ws.Shapes.AddShape(5, btn_left, btn_top, btn_width, btn_height)
    shp.Name = "btnPushToVista"

    # Style the button
    shp.Fill.ForeColor.RGB = 0x00703820    # Dark teal/green (BGR: 0x20, 0x38, 0x70 -> R=32, G=56, B=112 -> deep blue)
    # Actually, let's use a professional dark blue that stands out but isn't alarming
    # RGB in VBA/COM is BGR, so we need to compute:
    #   Deep blue:  R=31, G=56, B=100 -> BGR = (100 << 16) | (56 << 8) | 31
    shp.Fill.ForeColor.RGB = (100 << 16) | (56 << 8) | 31   # Deep navy blue

    # Text
    tf = shp.TextFrame2
    tf.TextRange.Text = "Push to Vista"
    tf.TextRange.Font.Size = 11
    tf.TextRange.Font.Bold = True
    tf.TextRange.Font.Fill.ForeColor.RGB = (255 << 16) | (255 << 8) | 255   # White text
    tf.TextRange.ParagraphFormat.Alignment = 2  # msoAlignCenter
    tf.VerticalAnchor = 3   # msoAnchorMiddle
    tf.MarginLeft = 4
    tf.MarginRight = 4

    # Rounded corners
    shp.Adjustments.Item(1) = 0.25   # Corner roundness (0-1)

    # Subtle shadow
    shp.Shadow.Type = 1              # msoShadow1
    shp.Shadow.ForeColor.RGB = (180 << 16) | (180 << 8) | 180  # Light gray shadow
    shp.Shadow.Transparency = 0.6
    shp.Shadow.OffsetX = 2
    shp.Shadow.OffsetY = 2

    # Border
    shp.Line.ForeColor.RGB = (80 << 16) | (40 << 8) | 20   # Slightly darker border
    shp.Line.Weight = 1

    # Assign macro
    shp.OnAction = "FormButtons.PushToVistaClick"

    print(f"  Created button 'btnPushToVista' at ({btn_left}, {btn_top}), {btn_width}x{btn_height}")
    print(f"  OnAction = FormButtons.PushToVistaClick")
    print("Button creation complete.")


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

    # Start dialog dismisser
    stop_event = threading.Event()
    dismiss_thread = threading.Thread(
        target=dismiss_excel_dialogs, args=(stop_event,), daemon=True
    )
    dismiss_thread.start()

    try:
        wb = excel.Workbooks.Open(WORKBOOK_PATH)
        time.sleep(2)  # Let workbook initialize

        # Step 1: Import modules
        reimport_modules(wb)

        # Step 2: Create button
        create_push_to_vista_button(wb)

        # Save
        print("\nSaving workbook ...")
        wb.Save()
        print("Saved.")

        print("\n=== Done ===")
        print(f"Workbook: {os.path.basename(WORKBOOK_PATH)}")
        print("Modules imported: LylesWIPData.bas, FormButtons.bas")
        print("Button created: 'Push to Vista' on Start sheet")
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
