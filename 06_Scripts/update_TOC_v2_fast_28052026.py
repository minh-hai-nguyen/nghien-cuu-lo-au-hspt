# -*- coding: utf-8 -*-
"""APPROACH FAST: Open Word COM, update fields, read populated TOC.
28/05/2026."""
import os, sys, io, shutil, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import win32com.client
import pythoncom

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
FILE = os.path.join(ROOT, 'Luận án TS', '01_LuanAn_v3_1_FixCoping_28052026.docx')
TEMP = os.path.join(ROOT, 'Luận án TS', '_temp_TOC_extract.docx')

# Make a temp copy to work with (don't modify original)
shutil.copy(FILE, TEMP)
print(f"Working on temp copy: {TEMP}")

pythoncom.CoInitialize()
word = None
try:
    word = win32com.client.DispatchEx('Word.Application')  # DispatchEx for new instance
    word.Visible = False
    word.DisplayAlerts = False

    print("Opening doc...")
    abs_temp = os.path.abspath(TEMP)
    doc = word.Documents.Open(abs_temp, ReadOnly=False)

    print("Update fields...")
    doc.Fields.Update()

    print("Reading TOC field content...")
    toc_entries = []
    # TablesOfContents collection
    tocs = doc.TablesOfContents
    print(f"  TablesOfContents count: {tocs.Count}")

    if tocs.Count > 0:
        toc = tocs.Item(1)
        toc.Update()  # ensure populated
        toc_range = toc.Range
        toc_text = toc_range.Text
        print(f"  TOC text length: {len(toc_text)} chars")
        # Print first 2000 chars
        print(f"  First 500 chars: {toc_text[:500]!r}")

        # Parse TOC text - typical format: "Heading text\tpage_number\r\n" or "Heading text......page\r\n"
        for line in toc_text.split('\r'):
            line = line.strip('\r\n\t ')
            if not line:
                continue
            # Find last number (page)
            import re
            m = re.search(r'(.+?)[\s\.\t]+(\d+)\s*$', line)
            if m:
                name = m.group(1).strip(' .\t')
                page = int(m.group(2))
                toc_entries.append({'name': name, 'page': page})

    print(f"  Extracted {len(toc_entries)} TOC entries")

    doc.Close(SaveChanges=False)
finally:
    if word is not None:
        try:
            word.Quit()
        except: pass
    pythoncom.CoUninitialize()

# Save entries
OUT = os.path.join(ROOT, '06_Scripts', '_toc_entries_28052026.json')
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(toc_entries, f, ensure_ascii=False, indent=2)
print(f"\nSaved entries: {OUT}")

# Cleanup temp
if os.path.exists(TEMP):
    try:
        os.remove(TEMP)
        print(f"Cleaned temp: {TEMP}")
    except: pass

# Print first 30 entries
print()
print("=== First 30 TOC entries ===")
for i, e in enumerate(toc_entries[:30]):
    print(f"  [{i}] p{e['page']:>3}: {e['name'][:70]}")
