# -*- coding: utf-8 -*-
"""List ALL page markers."""
import os, sys, io, re
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

PATH = '03_Ban-dich/VN022_UNICEF_VN_2022_SchoolFactors_FULL.docx'
d = Document(PATH)

for i, p in enumerate(d.paragraphs):
    t = p.text.strip()
    m = re.match(r'---\s*Trang\s+([\d–\-]+)', t)
    if m:
        print(f'  #{i}: {t[:70]}')
