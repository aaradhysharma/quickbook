"""
Open real Microsoft Excel, activate each sheet, and capture the Excel window.
"""
from __future__ import annotations

import ctypes
import sys
import time
from pathlib import Path

import win32com.client
import win32con
import win32gui
import win32ui
from PIL import Image

ROOT = Path(__file__).parent
XLSX = ROOT / "Brew_Haven_Cafe_Cost_Tracker.xlsx"
OUT = ROOT / "screenshots" / "real_excel"
OUT.mkdir(parents=True, exist_ok=True)

SHEETS = [
    ("Startup Costs", "real_01_startup_costs.png"),
    ("Ongoing Expenses", "real_02_ongoing_expenses.png"),
    ("Monthly Ledger", "real_03_monthly_ledger.png"),
    ("Profitability", "real_04_profitability.png"),
    ("Regression", "real_05_regression.png"),
]

user32 = ctypes.windll.user32
PW_RENDERFULLCONTENT = 2


def find_excel_hwnd() -> int:
    targets = []

    def enum_handler(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        cls = win32gui.GetClassName(hwnd)
        if cls == "XLMAIN":
            targets.append(hwnd)
        elif "Excel" in title and "Brew_Haven" in title:
            targets.append(hwnd)

    win32gui.EnumWindows(enum_handler, None)
    if not targets:
        def enum2(hwnd, _):
            if win32gui.IsWindowVisible(hwnd) and "Excel" in win32gui.GetWindowText(hwnd):
                targets.append(hwnd)

        win32gui.EnumWindows(enum2, None)
    if not targets:
        raise RuntimeError("Could not find Excel window handle")
    return targets[0]


def capture_window(hwnd: int, path: Path) -> None:
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    time.sleep(0.3)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    time.sleep(1.0)

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top
    if width < 100 or height < 100:
        raise RuntimeError(f"Excel window too small: {width}x{height}")

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
    save_dc.SelectObject(bitmap)

    ok = user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), PW_RENDERFULLCONTENT)
    if not ok:
        screen_dc = win32gui.GetDC(0)
        src = win32ui.CreateDCFromHandle(screen_dc)
        save_dc.BitBlt((0, 0), (width, height), src, (left, top), win32con.SRCCOPY)
        src.DeleteDC()
        win32gui.ReleaseDC(0, screen_dc)

    bmpinfo = bitmap.GetInfo()
    bmpstr = bitmap.GetBitmapBits(True)
    img = Image.frombuffer(
        "RGB",
        (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
        bmpstr,
        "raw",
        "BGRX",
        0,
        1,
    )

    extrema = img.convert("L").getextrema()
    if extrema == (0, 0):
        screen_dc = win32gui.GetDC(0)
        src = win32ui.CreateDCFromHandle(screen_dc)
        save_dc.BitBlt((0, 0), (width, height), src, (left, top), win32con.SRCCOPY)
        src.DeleteDC()
        win32gui.ReleaseDC(0, screen_dc)
        bmpstr = bitmap.GetBitmapBits(True)
        img = Image.frombuffer(
            "RGB",
            (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
            bmpstr,
            "raw",
            "BGRX",
            0,
            1,
        )

    img.save(path, "PNG")
    print(f"  saved {path.name} ({width}x{height}) PrintWindow={ok}")

    win32gui.DeleteObject(bitmap.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)


def main() -> int:
    if not XLSX.exists():
        print(f"Missing workbook: {XLSX}", file=sys.stderr)
        return 1

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.WindowState = -4137  # xlMaximized

    wb = None
    try:
        wb = excel.Workbooks.Open(str(XLSX.resolve()))
        time.sleep(2.0)
        hwnd = find_excel_hwnd()
        print(f"Excel HWND={hwnd} title={win32gui.GetWindowText(hwnd)!r}")

        for sheet_name, filename in SHEETS:
            ws = wb.Worksheets(sheet_name)
            ws.Activate()
            ws.Range("A1").Select()
            try:
                excel.ActiveWindow.Zoom = 90
            except Exception:
                pass
            time.sleep(1.2)
            out_path = OUT / filename
            capture_window(hwnd, out_path)
            print(f"Captured {sheet_name}")

        print("DONE")
        return 0
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
