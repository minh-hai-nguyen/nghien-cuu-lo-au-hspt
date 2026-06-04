# -*- coding: utf-8 -*-
"""Extract all text from VN022 UNICEF 2022 PDF for verification."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from pypdf import PdfReader

PDF = '02_Papers-goc/Viet-Nam/VN022_UNICEF_VN_2022_SchoolFactors.pdf'
reader = PdfReader(PDF)
print(f'Total pages: {len(reader.pages)}')

OUT = 'C:/Users/HLC/AppData/Local/Temp/vn022_pdf_full.txt'
with open(OUT, 'w', encoding='utf-8') as f:
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        f.write(f'\n===== PAGE {i+1} =====\n')
        f.write(text or '')
        f.write('\n')

print(f'Wrote {OUT} ({os.path.getsize(OUT):,} bytes)')
