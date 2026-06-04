# -*- coding: utf-8 -*-
"""Final cleanup: remove Essau 2012 + Du 2008 (not verified), keep only Goodman 1997."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
PATH = os.path.join(ROOT, '03_Ban-dich', 'VN022_UNICEF_VN_2022_SchoolFactors_FULL.docx')

FIXES = [
    ('Literature quốc tế về SDQ ở bối cảnh châu Á (Goodman 1997 — khung bên ngoài; Essau et al. 2012; Du et al. 2008) ghi nhận subscale',
     'Literature về psychometric properties của SDQ-25 (Goodman 1997, khung bên ngoài refs UNICEF) nhìn chung ghi nhận subscale'),
]

d = Document(PATH)
changes = 0
for p in d.paragraphs:
    t = p.text
    if not t.strip():
        continue
    new_t = t
    for find, rep in FIXES:
        if find in new_t:
            new_t = new_t.replace(find, rep)
    if new_t != t:
        for run in p.runs:
            run.text = ''
        if p.runs:
            p.runs[0].text = new_t
        changes += 1

d.save(PATH)
print(f'Fixed {changes} paragraphs')
