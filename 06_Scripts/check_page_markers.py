# -*- coding: utf-8 -*-
"""Check page markers in docx to find where to insert images."""
import os, sys, io, re
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

PATH = '03_Ban-dich/VN022_UNICEF_VN_2022_SchoolFactors_FULL.docx'
d = Document(PATH)

target_pages = [1, 6, 8, 11, 23, 26, 28, 35, 40, 43, 55, 62, 66, 70, 75, 76]
found = {}

for i, p in enumerate(d.paragraphs):
    t = p.text.strip()
    # Match patterns like "--- Trang 6 ---" or "Trang 6" or "--- Page 6 ---"
    m = re.search(r'[Tt]rang\s+(\d+)|[Pp]age\s+(\d+)', t)
    if m:
        page = int(m.group(1) or m.group(2))
        if page in target_pages:
            if page not in found:
                found[page] = (i, t[:80])

print('Page markers found:')
for page in sorted(target_pages):
    if page in found:
        idx, text = found[page]
        print(f'  p{page:03d}: paragraph #{idx} — "{text}"')
    else:
        print(f'  p{page:03d}: NOT FOUND')
