# -*- coding: utf-8 -*-
"""Better matching: search PDF files by author anywhere in filename."""
import os, sys, io, re, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

with open(os.path.join(ROOT, '02_Papers-goc', 'canonical_index.json'), encoding='utf-8') as f:
    canon = json.load(f)

unmatched = [
    ('Brown', '2024'), ('Cao', '2025'), ('Fitzgerald', '2024'),
    ('Hartanto', '2024'), ('Jiang', '2024'), ('Kang', '2026'),
    ('Kuyken', '2022'), ('Li', '2025'), ('Ngọc', '2018'),
    ('Walkup', '2008'), ('Wang', '2025'), ('Zugman', '2024'),
]

print('Searching for each unmatched citation in canonical_index + all PDFs:')
print('='*80)
import glob

all_pdfs = []
for sub in ['Viet-Nam', 'The-gioi_Au-My-Uc', 'The-gioi_Khac', 'Dong-Nam-A',
            '11-bai-ban-dau-va-mo-rong', 'Chua-phan-loai']:
    folder = os.path.join(ROOT, '02_Papers-goc', sub)
    if os.path.isdir(folder):
        for f in os.listdir(folder):
            if f.endswith('.pdf'):
                all_pdfs.append((sub, f))

for au, yr in unmatched:
    au_low = au.lower()
    found_any = False
    # Search canonical_index descriptors
    for cid, meta in canon.items():
        descr = meta.get('descriptor', '').lower()
        if au_low in descr and yr in descr:
            pdf = meta.get('pdf', '?')
            folder = meta.get('pdf_folder', '?')
            print(f'  {au} {yr}  →  {cid} [{meta.get("descriptor")}]  ({folder}/{pdf})')
            found_any = True
    # Also search raw PDF filenames
    for folder, fn in all_pdfs:
        if au_low in fn.lower() and yr in fn:
            m = re.match(r'(VN\d{3}|QT\d{3})', fn)
            cid = m.group(1) if m else 'no-canon-id'
            if cid != 'no-canon-id' and cid in canon:
                # Already reported above?
                continue
            print(f'  {au} {yr}  →  [RAW FILE]  {folder}/{fn}')
            found_any = True
    if not found_any:
        print(f'  {au} {yr}  ✗ NOT FOUND in library')
