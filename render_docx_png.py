"""docx → PDF(Word COM) → PNG(pymupdf)."""
import os, sys

import fitz
import win32com.client as win32

docx = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples", "drilldown_v14_color.docx")
base = docx.rsplit(".", 1)[0]
pdf = base + ".pdf"

word = win32.Dispatch("Word.Application")
word.Visible = False
try:
    d = word.Documents.Open(docx)
    d.ExportAsFixedFormat(pdf, 17)  # wdExportFormatPDF
    d.Close(False)
finally:
    word.Quit()

out = []
doc = fitz.open(pdf)
for i, page in enumerate(doc, 1):
    png = f"{base}_p{i}.png"
    page.get_pixmap(dpi=150).save(png)
    out.append(png)
print("\n".join(out))
