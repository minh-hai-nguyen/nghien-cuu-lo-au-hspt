# -*- coding: utf-8 -*-
"""Audit toan dien: 02_Papers-goc/ vs canonical_index vs CSDL vs KG vs RAG.
Bao cao chi tiet trong file .md gui thay."""
import os, sys, io, json
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
OUT = os.path.join(ROOT, '06_Scripts', 'AUDIT_PapersGoc_07062026.md')

# 1. Load canonical
with open(os.path.join(ROOT, '02_Papers-goc', 'canonical_index.json'),
          'r', encoding='utf-8') as f:
    canonical = json.load(f)

# 2. Load KG v2
with open(os.path.join(ROOT, '06_Scripts', 'author_kg_data',
                       'author_kg_v2.json'), 'r', encoding='utf-8') as f:
    kg = json.load(f)

# 3. Load CSDL
with open(os.path.join(ROOT, '04_Co-so-du-lieu',
                       'DATABASE_BAI_BAO_LO_AU.md'), 'r',
          encoding='utf-8') as f:
    csdl = f.read()

# === Inventory ===
all_pdfs = {}
by_folder = {}
for root, dirs, files in os.walk(os.path.join(ROOT, '02_Papers-goc')):
    for f in files:
        if f.lower().endswith('.pdf'):
            full = os.path.join(root, f)
            rel = os.path.relpath(full, os.path.join(ROOT, '02_Papers-goc'))
            all_pdfs[f.lower()] = rel
            folder = os.path.dirname(rel) or '(root)'
            by_folder[folder] = by_folder.get(folder, 0) + 1

# Canonical PDFs
canonical_pdfs = set()
for k, v in canonical.items():
    pdf = v.get('pdf')
    if pdf: canonical_pdfs.add(pdf.lower())

# Not in canonical
not_in_canonical = sorted(p for p in all_pdfs if p not in canonical_pdfs)
# In canonical but PDF missing
all_lower = {p.lower() for p in all_pdfs}
missing_pdfs = [(k, v.get('pdf')) for k, v in canonical.items()
                if v.get('pdf') and v.get('pdf').lower() not in all_lower]

# KG paper IDs
kg_papers = {n.get('id'): n for n in kg['nodes'] if n.get('type') == 'Paper'}

# === Write audit report ===
lines = []
lines.append('# AUDIT 02_Papers-goc/ vs CSDL + KG + RAG (07/06/2026)\n')
lines.append('## Tổng quan\n')
lines.append(f'- **Tổng số PDF trong `02_Papers-goc/`:** {len(all_pdfs)}\n')
lines.append(f'- **canonical_index.json:** {len(canonical)} entries '
             f'(trong đó {len(canonical_pdfs)} có file PDF)\n')
lines.append(f'- **KG v2 papers:** {len(kg_papers)} nodes\n')
lines.append(f'- **PDF không có trong canonical:** {len(not_in_canonical)}\n')
lines.append(f'- **Canonical trỏ tới PDF KHÔNG tồn tại:** {len(missing_pdfs)}\n')
lines.append('\n')

lines.append('## Phân bố PDF theo folder\n')
lines.append('| Folder | Số PDF |\n|---|---|\n')
for folder, n in sorted(by_folder.items(), key=lambda x: -x[1]):
    lines.append(f'| `{folder}` | {n} |\n')
lines.append('\n')

lines.append(f'## PDF chưa có trong canonical_index ({len(not_in_canonical)})\n')
lines.append('Cần phân loại + thêm vào canonical_index.json:\n\n')
for p in not_in_canonical:
    lines.append(f'- `{all_pdfs[p]}`\n')
lines.append('\n')

if missing_pdfs:
    lines.append(f'## Canonical entries trỏ tới PDF KHÔNG tồn tại '
                 f'({len(missing_pdfs)})\n')
    for k, p in missing_pdfs:
        lines.append(f'- `{k}` → `{p}` (không tìm thấy)\n')
    lines.append('\n')

# Cross-check 4 papers Q1 cần
lines.append('## Cross-check 4 paper bài Q1 vừa thêm (07/06/2026)\n\n')
q1_papers_check = [
    ('Karasu 1986 AJ Psychotherapy', 'karasu1986.pdf',
     'PA_KARASU_1986_AJPSY'),
    ('Karasu Einstein bio sketch', 'biosketch_karasu_062112.pdf',
     'AU_KARASU_TB (trong author node)'),
    ('Rose 2002 — Co-rumination', '(chưa download)',
     'PA_ROSE_2002'),
    ('Stankov 2010 — Confucian', '(chưa download)',
     'PA_STANKOV_2010'),
    ('Small & Blanc 2021 — Tam giao', '(chưa download)',
     'PA_SMALL_BLANC_2021'),
]
lines.append('| Paper | PDF | KG node |\n|---|---|---|\n')
for name, pdf, kg_id in q1_papers_check:
    pdf_status = '✓ có' if pdf.lower() in all_pdfs else '⏳ chưa có'
    if pdf == '(chưa download)':
        pdf_status = '⏳ cần download'
    lines.append(f'| {name} | {pdf_status} `{pdf}` | `{kg_id}` ✓ |\n')
lines.append('\n')

lines.append('## Hành động ưu tiên\n\n')
lines.append('1. **Thêm `karasu1986.pdf` và `biosketch_karasu_062112.pdf` '
             'vào `canonical_index.json`** — em sẽ làm trong vòng tiếp '
             'theo.\n')
lines.append('2. **Download 3 PDF còn thiếu** (Rose 2002, Stankov 2010, '
             'Small & Blanc 2021) — Rose + Stankov paywall, Small & Blanc '
             'mở qua PMC PMC7820702.\n')
lines.append(f'3. **{len(not_in_canonical)} PDF chưa categorize** trong '
             '02_Papers-goc/ — chủ yếu là `Chua-phan-loai/` và '
             '`11-bai-ban-dau-va-mo-rong/`. Cần script categorize + '
             'metadata extraction.\n')
lines.append('4. **RAG rebuild** sau khi categorize: chạy '
             '`tro-ly-nghien-cuu-tam-ly/build_rag.py`.\n')

with open(OUT, 'w', encoding='utf-8') as f:
    f.writelines(lines)
print(f'Saved: {OUT}')
print(f'  {len(all_pdfs)} PDFs total')
print(f'  {len(canonical)} canonical entries')
print(f'  {len(not_in_canonical)} not in canonical')
print(f'  {len(missing_pdfs)} canonical missing PDF')
