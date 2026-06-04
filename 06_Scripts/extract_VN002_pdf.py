# -*- coding: utf-8 -*-
"""Extract VN002 VNAMHS 2022 PDF text per-page for translation."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from pypdf import PdfReader

PDF = '02_Papers-goc/Viet-Nam/VN002_VNAMHS_2022_National.pdf'
OUT = 'C:/Users/HLC/AppData/Local/Temp/vnamhs_pdf_full.txt'

reader = PdfReader(PDF)
total_chars = 0
with open(OUT, 'w', encoding='utf-8') as f:
    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ''
        f.write(f'\n===== PAGE {i+1} =====\n')
        f.write(text)
        f.write('\n')
        total_chars += len(text)

print(f'Pages: {len(reader.pages)}')
print(f'Chars: {total_chars:,}')
print(f'Out: {OUT} ({os.path.getsize(OUT):,} bytes)')
