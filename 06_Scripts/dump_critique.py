# -*- coding: utf-8 -*-
"""Dump current PHẢN BIỆN section content."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
PATH = os.path.join(ROOT, '03_Ban-dich', 'VN022_UNICEF_VN_2022_SchoolFactors_FULL.docx')

d = Document(PATH)
for i in range(1030, 1071):
    t = d.paragraphs[i].text
    print(f'#{i}: {t}')
    print('-'*70)
