# -*- coding: utf-8 -*-
"""Sap xep tai lieu tham khao theo dung 2 bai bao mau VJES:
MOT danh sach tron chung, sap theo CHU DAU (ho/surname) cua moi muc.
Boi do (font do) cac muc bi doi vi tri. 19/05/2026."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FOLDER = r'c:\Users\OS\OneDrive\read_books\Lo-au\bai-bao-khgdvn\tai-lieu'

# --- Bang xep van tieng Viet (a < a < a < b < c < d < d ...), gop dau thanh ---
BASE = ['a', 'ă', 'â', 'b', 'c', 'd', 'đ', 'e', 'ê', 'f', 'g', 'h', 'i', 'j',
        'k', 'l', 'm', 'n', 'o', 'ô', 'ơ', 'p', 'q', 'r', 's', 't', 'u', 'ư',
        'v', 'w', 'x', 'y', 'z']
RANK = {ch: i for i, ch in enumerate(BASE)}
VARIANTS = {
    'a': 'àáảãạ', 'ă': 'ằắẳẵặ', 'â': 'ầấẩẫậ',
    'e': 'èéẻẽẹ', 'ê': 'ềếểễệ', 'i': 'ìíỉĩị',
    'o': 'òóỏõọ', 'ô': 'ồốổỗộ', 'ơ': 'ờớởỡợ',
    'u': 'ùúủũụ', 'ư': 'ừứửữự', 'y': 'ỳýỷỹỵ',
}
for base, vs in VARIANTS.items():
    for v in vs:
        RANK[v] = RANK[base]

def colkey(text):
    """Khoa sap xep theo van tieng Viet, tren toan bo noi dung muc.
    Ky tu khong phai chu cai (dau phay, khoang trang, so...) xep TRUOC moi
    chu cai (-1) => 'khong co gi' di truoc 'co gi' (vd. 'Li' truoc 'Liu')."""
    return tuple(RANK.get(c, -1) for c in text.strip().lower())

def color_red(p):
    for r in p._p.iter(qn('w:r')):
        rPr = r.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr'); r.insert(0, rPr)
        for c in rPr.findall(qn('w:color')): rPr.remove(c)
        c = OxmlElement('w:color'); c.set(qn('w:val'), 'CC0000'); rPr.append(c)

def process(src, out):
    doc = Document(src)
    content = [p for p in doc.paragraphs if p.text.strip()]
    new = sorted(content, key=lambda p: colkey(p.text))
    oldidx = {id(p): i for i, p in enumerate(content)}
    moved = 0
    for newpos, p in enumerate(new):
        if oldidx[id(p)] != newpos:
            color_red(p); moved += 1
    body = doc.element.body
    for p in content:
        body.remove(p._p)
    anchor = body[0]                       # doan rong cuoi / sectPr
    for p in new:
        anchor.addprevious(p._p)
    doc.save(out)
    return content, new, oldidx, moved

for n in (1, 2, 3):
    src = os.path.join(FOLDER, 'Tham khảo %d.docx' % n)
    out = os.path.join(FOLDER, 'Tham khảo %d_da_sap_xep_19052026.docx' % n)
    content, new, oldidx, moved = process(src, out)
    print('=== Tham khao %d: %d muc | %d muc doi vi tri ===' % (n, len(content), moved))
    for i, p in enumerate(new, 1):
        flag = 'ĐỎ' if oldidx[id(p)] != i-1 else '  '
        print('%2d. %s %s' % (i, flag, p.text[:64]))
    print()
