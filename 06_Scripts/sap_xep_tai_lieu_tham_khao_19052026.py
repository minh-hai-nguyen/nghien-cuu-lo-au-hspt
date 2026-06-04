# -*- coding: utf-8 -*-
"""Sap xep lai thu tu tai lieu tham khao file 'Tham khao 1.docx' theo quy dinh VJES:
- Tai lieu tac gia nuoc ngoai: xep truoc, theo Ho (surname).
- Tai lieu tac gia Viet Nam: xep sau, theo Ten (given name = tu cuoi).
Boi do (font do) cac muc bi doi vi tri. 19/05/2026."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FOLDER = r'c:\Users\OS\OneDrive\read_books\Lo-au\bai-bao-khgdvn\tai-lieu'
SRC = os.path.join(FOLDER, 'Tham khảo 1.docx')
OUT = os.path.join(FOLDER, 'Tham khảo 1_da_sap_xep_19052026.docx')

doc = Document(SRC)
content = [p for p in doc.paragraphs if p.text.strip()]
print('So muc tai lieu:', len(content))

# Thu tu moi: khoi 1 (nuoc ngoai, theo Ho) + khoi 2 (Viet Nam, theo Ten)
new_order_keys = [
    # --- Khoi 1: tac gia nuoc ngoai, theo Ho ---
    'Brunborg', 'Chen, Z.', 'Compas', 'Fassi', 'Jagiello',
    'Lee, J.', 'Moore, S. E.', 'Pascoe', 'Sowislo', 'Xu, Q.',
    # --- Khoi 2: tac gia Viet Nam, theo Ten ---
    'Phạm Thị Thu Hoa',     # ten: Hoa
    'Hoàng Trung Học',      # ten: Học
    'Trần Hồ Vĩnh Lộc',     # ten: Lộc
    'Trần Thị Mỵ Lương',    # ten: Lương
    'Phạm Thị Ngọc',        # ten: Ngọc
    'Trúc Thanh Thái',      # ten: Thái
    'Phạm Sỹ Tiến',         # ten: Tiến
    'Đinh Thị Hồng Vân',    # ten: Vân
    'Trần Thảo Vi',         # ten: Vi
    'Viện Xã hội học',      # to chuc: Viện
    'Ngô Anh Vinh',         # ten: Vinh
]
assert len(new_order_keys) == len(content) == 21, (len(new_order_keys), len(content))

# Anh xa key -> doan van (paragraph): khop tu DAU dong (ten tac gia dau tien)
new_paras = []
for k in new_order_keys:
    matches = [p for p in content if p.text.strip().startswith(k)]
    assert len(matches) == 1, ('Key khong duy nhat: %r -> %d' % (k, len(matches)))
    new_paras.append(matches[0])
assert len({id(p) for p in new_paras}) == 21

old_index = {id(p): i for i, p in enumerate(content)}

def color_red(p):
    """To mau do (font) cho moi run trong doan van, ke ca run trong hyperlink."""
    for r in p._p.iter(qn('w:r')):
        rPr = r.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr'); r.insert(0, rPr)
        for c in rPr.findall(qn('w:color')): rPr.remove(c)
        c = OxmlElement('w:color'); c.set(qn('w:val'), 'CC0000'); rPr.append(c)

moved = []
for newpos, p in enumerate(new_paras):
    oldpos = old_index[id(p)]
    if oldpos != newpos:
        color_red(p); moved.append((oldpos, newpos, p.text[:55]))

# Sap xep lai cac phan tu <w:p> trong than van ban
body = doc.element.body
for p in content:
    body.remove(p._p)
anchor = body[0]            # doan rong cuoi cung (hoac sectPr)
for p in new_paras:
    anchor.addprevious(p._p)

doc.save(OUT)

print('So muc bi doi vi tri (to do):', len(moved))
print('--- THU TU MOI ---')
for i, p in enumerate(new_paras, 1):
    flag = 'ĐỎ ' if old_index[id(p)] != i-1 else '   '
    print(f'{i:2d}. {flag}{p.text[:70]}')
print('Da luu:', OUT)
