# -*- coding: utf-8 -*-
"""Update TOC table dung Word COM - lay page number chinh xac.
CHI dung Heading style cua Word (KHONG dung regex text pattern de tranh nham doan dai).
28/05/2026."""
import os, sys, io, shutil, time
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import win32com.client
from docx import Document

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
FILE = os.path.join(ROOT, 'Luận án TS', '01_LuanAn_v3_1_FixCoping_28052026.docx')
BACKUP = FILE.replace('.docx', '_BEFORE_TOC_UPDATE.docx')
shutil.copy(FILE, BACKUP)
print(f"Backup: {BACKUP}")

# ============================================================
# 1. Mo file qua Word COM, repaginate, lay heading + page
# ============================================================
print("\n[1] Mo file qua Word COM...")
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

try:
    doc = word.Documents.Open(FILE, ReadOnly=False)
    # Update fields + repaginate
    print("  Updating fields + repaginating...")
    doc.Fields.Update()
    doc.Repaginate()

    # Constants
    wdActiveEndPageNumber = 3

    # Iterate paragraphs, find headings (style starts with 'Heading')
    print("  Scanning paragraphs for Heading styles...")
    headings = []  # list of (level, text, page)
    paragraphs = doc.Paragraphs
    total = paragraphs.Count
    print(f"  Total paragraphs: {total}")

    for i in range(1, total + 1):
        p = paragraphs.Item(i)
        try:
            style_name = p.Style.NameLocal
            # Vietnamese Word may have 'Tiêu đề 1', 'Tiêu đề 2', or English 'Heading 1', 'Heading 2'
            level = None
            for kw, lv in [('Heading 1', 1), ('Heading 2', 2), ('Heading 3', 3),
                          ('Heading 4', 4), ('Heading 5', 5),
                          ('Tiêu đề 1', 1), ('Tiêu đề 2', 2), ('Tiêu đề 3', 3),
                          ('Tiêu đề 4', 4), ('Tiêu đề 5', 5)]:
                if kw in style_name:
                    level = lv
                    break
            if level is None:
                continue
            text = p.Range.Text.strip().replace('\r', '').replace('\x07', '')
            if not text:
                continue
            # Skip overly long (paragraph text that happens to be styled heading by mistake)
            if len(text) > 250:
                print(f"    SKIP (qua dai, level={level}): {text[:80]!r}...")
                continue
            page = p.Range.Information(wdActiveEndPageNumber)
            headings.append((level, text, page, i))
        except Exception as e:
            pass

    print(f"\n  Tim thay {len(headings)} headings:")
    for lv, txt, pg, idx in headings[:20]:
        print(f"    L{lv} p{pg}: {txt[:60]!r}")
    if len(headings) > 20:
        print(f"    ... va {len(headings)-20} headings khac")

    doc.Close(SaveChanges=False)
finally:
    word.Quit()

# Save headings list for next step
import json
HEADINGS_FILE = os.path.join(ROOT, '06_Scripts', '_headings_28052026.json')
with open(HEADINGS_FILE, 'w', encoding='utf-8') as f:
    json.dump([{'level': lv, 'text': txt, 'page': pg, 'idx': idx} for lv, txt, pg, idx in headings], f, ensure_ascii=False, indent=2)
print(f"\n[2] Headings saved to: {HEADINGS_FILE}")
print(f"    Total headings: {len(headings)}")
