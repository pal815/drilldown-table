# -*- coding: utf-8 -*-
"""방향(3)×테마(3) 미리보기 썸네일 9개를 assets/previews/ 에 생성.
   엔진으로 샘플 xlsx 생성 → Excel COM(CopyPicture→차트 Export)로 PNG → PIL 리사이즈."""
import os
import sys
import time
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), os.pardir))
import drilldown_table as G
from PIL import Image
import win32com.client as win32

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "previews")
os.makedirs(OUT, exist_ok=True)
RESAMPLE = getattr(Image, "Resampling", Image).LANCZOS
xlScreen, xlBitmap = 1, 2

excel = win32.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    for orient in ("row", "column", "both"):
        model = G.SAMPLE_2D if orient == "both" else G.SAMPLE
        for theme in ("color", "grey", "mono"):
            tmp = os.path.join(tempfile.gettempdir(), f"_prev_{orient}_{theme}.xlsx")
            G.generate(model, tmp, orient, theme)
            wb = excel.Workbooks.Open(tmp)
            ws = wb.Worksheets(1)
            ws.Activate()
            rng = ws.UsedRange
            for attempt in range(4):
                try:
                    rng.CopyPicture(xlScreen, xlBitmap)
                    break
                except Exception:
                    time.sleep(0.6)
            else:
                raise RuntimeError("CopyPicture 반복 실패")
            time.sleep(0.4)
            co = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
            ch = co.Chart
            ch.ChartArea.Format.Fill.Visible = False
            ch.ChartArea.Format.Line.Visible = False
            co.Activate()
            time.sleep(0.2)
            ch.Paste()
            time.sleep(0.2)
            png = os.path.join(OUT, f"{orient}_{theme}.png")
            ch.Export(png, "PNG")
            co.Delete()
            wb.Close(False)
            os.remove(tmp)
            im = Image.open(png)
            maxw = 900
            if im.width > maxw:
                im = im.resize((maxw, int(im.height * maxw / im.width)), RESAMPLE)
            im.save(png)
            print("preview:", os.path.basename(png), im.size)
finally:
    excel.Quit()
print("done ->", OUT)
