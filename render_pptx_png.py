"""pptx 각 슬라이드를 PNG로 (PowerPoint COM)."""
import os, sys

import win32com.client as win32

pptx = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples", "drilldown_v14_color.pptx")
base = pptx.rsplit(".", 1)[0]

ppt = win32.Dispatch("PowerPoint.Application")
out = []
try:
    pres = ppt.Presentations.Open(pptx, WithWindow=False)
    for i, slide in enumerate(pres.Slides, 1):
        png = f"{base}_ppt{i}.png"
        slide.Export(png, "PNG", 1800, 1013)
        out.append(png)
    pres.Close()
finally:
    ppt.Quit()
print("\n".join(out))
