# -*- coding: utf-8 -*-
"""README 쇼케이스용 결과 스크린샷 생성기 (Windows + Excel 필요).
   엔진으로 .xlsx를 만들고 Excel COM(CopyPicture→차트 Export)으로 PNG 렌더 → docs/images/.
   재현: cd docs && py make_screenshots.py
"""
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # 리포 루트
sys.path.insert(0, ROOT)
import drilldown_table as G
from PIL import Image
import win32com.client as win32

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "images")
os.makedirs(OUT, exist_ok=True)
TMP = os.environ.get("TEMP", os.getcwd())
RES = getattr(Image, "Resampling", Image).LANCZOS

# 다단계(3단계) 모델 — 가지별 깊이 다름(A=3단계, B=2단계), 합계 명시
MULTI = {
    "title": "2024년 사업부별 실적 (다단계)",
    "unit": "단위: 억원 / %",
    "level_labels": ["사업부", "지역", "분기"],
    "attributes": ["매출", "영업이익"],
    "nodes": [
        {"label": "A사업부", "summary": [1200, 180], "children": [
            {"label": "국내", "summary": [700, 120], "children": [
                {"label": "Q1", "values": [300, 55]},
                {"label": "Q2", "values": [400, 65]},
            ]},
            {"label": "해외", "values": [500, 60]},
        ]},
        {"label": "B사업부", "summary": [1000, 230], "children": [
            {"label": "Q1", "values": [480, 110]},
            {"label": "Q2", "values": [520, 120]},
        ]},
    ],
}

# (모델, orient, theme, 출력파일명)
JOBS = [
    (G.SAMPLE,    "row",    "color", "row_color"),
    (G.SAMPLE,    "column", "color", "column_color"),
    (G.SAMPLE_2D, "both",   "color", "both_color"),
    (MULTI,       "row",    "color", "multilevel_color"),
    (G.SAMPLE,    "row",    "grey",  "row_grey"),
    (G.SAMPLE,    "row",    "mono",  "row_mono"),
]

# 1) xlsx 생성
made = []
for model, orient, theme, name in JOBS:
    xlsx = os.path.join(TMP, f"_shot_{name}.xlsx")
    G.generate(model, xlsx, orient, theme)
    made.append((xlsx, os.path.join(OUT, name + ".png")))

# 2) Excel COM 렌더
xl = win32.Dispatch("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    for xlsx, png in made:
        wb = xl.Workbooks.Open(xlsx)
        ws = wb.Worksheets(1)
        ws.Activate()
        try:
            xl.ActiveWindow.Zoom = 150   # 선명도↑
        except Exception:
            pass
        rng = ws.UsedRange
        for _ in range(4):
            try:
                rng.CopyPicture(1, 2)
                break
            except Exception:
                time.sleep(0.6)
        time.sleep(0.4)
        co = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
        ch = co.Chart
        ch.ChartArea.Format.Fill.Visible = False
        ch.ChartArea.Format.Line.Visible = False
        co.Activate()
        time.sleep(0.2)
        ch.Paste()
        time.sleep(0.2)
        ch.Export(png, "PNG")
        co.Delete()
        wb.Close(False)
        im = Image.open(png)
        if im.width > 1000:
            im = im.resize((1000, int(im.height * 1000 / im.width)), RES)
            im.save(png)
        print("rendered:", os.path.basename(png), im.size)
finally:
    xl.Quit()

for xlsx, _ in made:
    try:
        os.remove(xlsx)
    except OSError:
        pass
print("done ->", OUT)
