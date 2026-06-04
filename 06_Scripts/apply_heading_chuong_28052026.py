# -*- coding: utf-8 -*-
"""Apply Heading style cho cac doan auto-numbered trong CHUONG sections
ma chua co Heading style.
28/05/2026."""
import os, sys, io, re
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import win32com.client, pythoncom
import shutil

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
FILE = os.path.join(ROOT, 'Luận án TS', '01_LuanAn_v3_1_FixCoping_28052026.docx')
TEMP = os.path.join(ROOT, 'Luận án TS', '_temp_scan.docx')
shutil.copy(FILE, TEMP)

# Use Word COM to scan auto-numbered paragraphs WITH their list level
print("Scanning all auto-numbered paragraphs in doc...")
pythoncom.CoInitialize()
word = None
candidates = []  # (list_str, text, style, doc_index, page)
try:
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False; word.DisplayAlerts = False
    doc = word.Documents.Open(os.path.abspath(TEMP), ReadOnly=False)
    doc.Fields.Update(); doc.Repaginate()

    # Iterate paragraphs (might be slow but only check ones in body, not table)
    print("  Iterating paragraphs...")
    paragraphs = doc.Paragraphs
    # Use OK iteration with index limit
    for i in range(1, paragraphs.Count + 1):
        try:
            p = paragraphs.Item(i)
            if p.Range.Tables.Count > 0: continue
            # Check if has list numbering
            list_str = ''
            try:
                list_str = p.Range.ListFormat.ListString
                list_str = ''.join(ch for ch in list_str if ord(ch) >= 32).strip()
            except: continue
            if not list_str: continue
            # Check current style
            style_name = p.Style.NameLocal
            if 'Heading' in style_name or 'Tiêu đề' in style_name: continue
            # Get text
            text = p.Range.Text.strip()
            text = ''.join(ch for ch in text if ord(ch) >= 32 or ch in '\t').strip()
            if not text or len(text) > 250: continue
            # Get page
            page = p.Range.Information(3)
            # Compute heading level from list_str
            # E.g., "1." = L2 in mo_dau context, "1.1" = L2 in chuong, "1.1.5" = L3, "1.1.5.1" = L4
            n_dots = list_str.count('.')
            level = None
            # Heuristic: count number parts
            parts = re.findall(r'\d+', list_str)
            n_parts = len(parts)
            if n_parts == 1:
                level = 2  # "1." → H2
            elif n_parts == 2:
                level = 2  # "1.1" → H2 (in chương) or H3 (in mo_dau)
            elif n_parts == 3:
                level = 3  # "1.1.1" → H3
            elif n_parts == 4:
                level = 4  # "1.1.1.1" → H4
            if level is None: continue
            candidates.append({
                'list_str': list_str, 'text': text, 'style': style_name,
                'index': i, 'page': page, 'level': level, 'n_parts': n_parts,
            })
        except Exception as e:
            continue
    doc.Close(SaveChanges=False)
finally:
    if word:
        try: word.Quit()
        except: pass
    pythoncom.CoUninitialize()
    if os.path.exists(TEMP):
        try: os.remove(TEMP)
        except: pass

print(f"\n  Found {len(candidates)} auto-numbered Normal-styled paragraphs")
print()
for c in candidates:
    print(f"  L{c['level']} p{c['page']} [{c['list_str']}] (style={c['style']!r}): {c['text'][:80]}")
