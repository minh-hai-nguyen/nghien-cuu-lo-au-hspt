# -*- coding: utf-8 -*-
"""Lay heading + page tu Word COM, BO QUA paragraph trong bang.
Dung Range.Tables.Count == 0 de loc.
28/05/2026."""
import os, sys, io, shutil, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import win32com.client
import pythoncom

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
FILE = os.path.join(ROOT, 'Luận án TS', '01_LuanAn_v3_1_FixCoping_28052026.docx')
TEMP = os.path.join(ROOT, 'Luận án TS', '_temp_TOC_extract2.docx')

shutil.copy(FILE, TEMP)
print(f"Working on temp: {TEMP}")

pythoncom.CoInitialize()
word = None
headings = []
try:
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False

    print("Opening...")
    doc = word.Documents.Open(os.path.abspath(TEMP), ReadOnly=False)
    print("Update fields + repaginate...")
    doc.Fields.Update()
    doc.Repaginate()

    wdActiveEndPageNumber = 3

    # Use Document.Range with Find for each heading style — faster than iterating all paragraphs
    print("Searching for Heading-styled paragraphs via Find...")

    for level in [1, 2, 3, 4]:
        for style_name in [f'Heading {level}', f'Tiêu đề {level}']:
            try:
                rng = doc.Content
                find = rng.Find
                find.ClearFormatting()
                find.Style = doc.Styles(style_name)
                find.Text = ''
                find.Forward = True
                find.Wrap = 0  # wdFindStop
                find.Format = True

                # Iterate all instances
                while find.Execute():
                    # rng now points to the found paragraph
                    para = rng.Paragraphs.First
                    text = para.Range.Text.strip().replace('\r', '').replace('\x07', '')
                    # Skip if inside table
                    in_table = para.Range.Tables.Count > 0
                    if in_table:
                        # Move past this paragraph and continue
                        rng.Collapse(0)  # wdCollapseEnd
                        continue
                    if not text:
                        rng.Collapse(0)
                        continue
                    if len(text) > 250:
                        rng.Collapse(0)
                        continue
                    page = para.Range.Information(wdActiveEndPageNumber)
                    headings.append({'level': level, 'text': text, 'page': page})
                    rng.Collapse(0)
            except Exception as e:
                pass  # Style not found for this language

    # Also catch CHƯƠNG paragraphs even if no Heading style applied (only if NOT in table)
    print(f"Found {len(headings)} headings via Heading styles")

    doc.Close(SaveChanges=False)
finally:
    if word is not None:
        try:
            word.Quit()
        except: pass
    pythoncom.CoUninitialize()

# Dedupe
seen = set()
unique_headings = []
for h in headings:
    key = (h['text'], h['page'])
    if key in seen: continue
    seen.add(key)
    unique_headings.append(h)
print(f"After dedup: {len(unique_headings)} headings")

# Sort by page then by level
unique_headings.sort(key=lambda h: (h['page'], h['level']))

# Save
OUT = os.path.join(ROOT, '06_Scripts', '_headings_real_28052026.json')
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(unique_headings, f, ensure_ascii=False, indent=2)
print(f"Saved: {OUT}")

# Cleanup
if os.path.exists(TEMP):
    try:
        os.remove(TEMP)
    except: pass

# Display
print()
print(f"=== {len(unique_headings)} real headings (sorted by page) ===")
for i, h in enumerate(unique_headings[:40]):
    print(f"  [{i:>3}] L{h['level']} p{h['page']:>3}: {h['text'][:70]}")
if len(unique_headings) > 40:
    print(f"  ... and {len(unique_headings)-40} more")
