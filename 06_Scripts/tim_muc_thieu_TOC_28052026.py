# -*- coding: utf-8 -*-
"""Tim cac heading co trong doc nhung CHUA co trong bang TOC.
Phan loai theo muc do quan trong.
28/05/2026."""
import os, sys, io, re, shutil
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import win32com.client, pythoncom
from docx import Document

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
FILE = os.path.join(ROOT, 'Luận án TS', '01_LuanAn_v3_1_FixCoping_28052026.docx')

# Extract heading + page via Word COM
TEMP = os.path.join(ROOT, 'Luận án TS', '_temp_find_missing.docx')
shutil.copy(FILE, TEMP)
pythoncom.CoInitialize()
word = None
headings = []
try:
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False; word.DisplayAlerts = False
    doc = word.Documents.Open(os.path.abspath(TEMP), ReadOnly=False)
    doc.Fields.Update(); doc.Repaginate()
    for level in [1, 2, 3, 4, 5]:
        for style_name in [f'Heading {level}', f'Tiêu đề {level}']:
            try:
                rng = doc.Content
                find = rng.Find
                find.ClearFormatting(); find.Style = doc.Styles(style_name)
                find.Text = ''; find.Forward = True; find.Wrap = 0; find.Format = True
                while find.Execute():
                    para = rng.Paragraphs.First
                    text = para.Range.Text.strip().replace('\r','').replace('\x07','')
                    if para.Range.Tables.Count > 0: rng.Collapse(0); continue
                    if not text or len(text) > 250: rng.Collapse(0); continue
                    page = para.Range.Information(3)
                    headings.append({'level': level, 'text': text, 'page': page})
                    rng.Collapse(0)
            except: pass
    doc.Close(SaveChanges=False)
finally:
    if word:
        try: word.Quit()
        except: pass
    pythoncom.CoUninitialize()
    if os.path.exists(TEMP):
        try: os.remove(TEMP)
        except: pass

# Dedupe
seen = set(); unique = []
for h in headings:
    k = (h['text'], h['page'])
    if k not in seen:
        seen.add(k); unique.append(h)
headings = sorted(unique, key=lambda h: h['page'])

# Get TOC table content
d = Document(FILE)
toc_table = None
for tb in d.tables:
    if len(tb.rows) > 0 and tb.rows[0].cells[0].text.strip().upper() == 'MỞ ĐẦU':
        toc_table = tb; break
toc_names = [row.cells[0].text.strip() for row in toc_table.rows]

def normalize(s):
    s = s.strip().lower()
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'^[\d\.]+\s*', '', s)
    return s.strip(' .')

toc_norms = set()
for n in toc_names:
    toc_norms.add(normalize(n))
    toc_norms.add(re.sub(r'\s+', ' ', n.strip().lower()))

# Find missing
print("="*80)
print("HEADINGS CO TRONG DOC NHUNG CHUA CO TRONG BANG MUC LUC")
print("="*80)

missing_by_level = {1: [], 2: [], 3: [], 4: [], 5: []}
for h in headings:
    norm = normalize(h['text'])
    full = re.sub(r'\s+', ' ', h['text'].strip().lower())
    if norm in toc_norms or full in toc_norms: continue
    # Fuzzy
    if any(norm and len(norm) > 8 and (norm in t or (len(t) > 8 and t in norm)) for t in toc_norms):
        continue
    # Skip cover-type headings
    upper = h['text'].upper().strip()
    if upper in ('MỤC LỤC', 'DANH MỤC CÁC TỪ VIẾT TẮT', 'DANH MỤC BẢNG',
                 'DANH MỤC HÌNH', 'DANH MỤC SƠ ĐỒ', 'MỞ ĐẦU',
                 'LỜI CAM ĐOAN', 'LỜI CẢM ƠN'):
        continue
    missing_by_level[h['level']].append(h)

print()
total = 0
for lv in [1, 2, 3, 4, 5]:
    items = missing_by_level[lv]
    if not items: continue
    print(f"\n--- Level {lv} ({len(items)} items) ---")
    for h in items:
        print(f"  p{h['page']:>3}: {h['text']}")
    total += len(items)
print(f"\nTong cong THIEU: {total} heading")
