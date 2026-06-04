# -*- coding: utf-8 -*-
"""Verify my cross-reference claims about VN001 (Hoa 2024), VN029 (Duong 2025), VN030 (Happy House)."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

for name in ['VN001_Hoa_2024_Frontiers_HaNoi.docx',
             # try to find VN029 and VN030
             ]:
    path = os.path.join('Tom-tat-tung-bai', name)
    if not os.path.exists(path):
        print(f'NOT FOUND: {path}')
        continue
    d = Document(path)
    text = '\n'.join(p.text for p in d.paragraphs if p.text.strip())
    print(f'=== {name} ===')
    print(f'Total chars: {len(text):,}')
    # Check key claims
    for claim in ['3910', '3.910', '40,6', '23,5', 'GAD-7', 'PHQ-9', 'nữ', 'THPT', 'Thái Nguyên', 'Cần Thơ', 'Hà Nội']:
        found = claim in text
        print(f'  "{claim}": {"YES" if found else "NO"}')
    print()
