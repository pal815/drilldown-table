"""각 시트의 UsedRange를 PNG로 내보낸다 (Excel CopyPicture → 차트 Export 트릭)."""
import os, sys
import time

import win32com.client as win32

xlsx = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples", "drilldown_v14_color.xlsx")
base = xlsx.rsplit(".", 1)[0]

xlScreen, xlBitmap = 1, 2
excel = win32.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
out = []
try:
    wb = excel.Workbooks.Open(xlsx)
    for idx, ws in enumerate(wb.Worksheets, 1):
        rng = ws.UsedRange
        rng.CopyPicture(xlScreen, xlBitmap)
        time.sleep(0.4)
        co = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
        ch = co.Chart
        ch.ChartArea.Format.Fill.Visible = False
        ch.ChartArea.Format.Line.Visible = False
        co.Activate()
        time.sleep(0.2)
        ch.Paste()
        time.sleep(0.2)
        png = f"{base}_{idx}.png"
        ch.Export(png, "PNG")
        co.Delete()
        out.append(png)
    wb.Close(False)
finally:
    excel.Quit()
print("\n".join(out))
