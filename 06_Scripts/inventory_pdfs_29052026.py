# -*- coding: utf-8 -*-
"""Inventory toan bo PDF trong kho: trich metadata + first-page text.
Output: catalog JSON cho buoc tiep theo (cross-check vs cited references).
29/05/2026."""
import os, sys, io, json, re
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import fitz  # PyMuPDF

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
SCAN_FOLDERS = [
    '02_Papers-goc',
    'bai-bao-khgdvn',
    'Bai_goc_BaoCao_CanThiep_10042026',
    'paper-may',
    'madam Thao',
]

results = []
errors = []


def extract_year_from_text(text, max_year=2027):
    """Tim year (1990-2027) trong text."""
    matches = re.findall(r'\b(19[89]\d|20[0-2]\d)\b', text)
    if not matches:
        return None
    years = [int(m) for m in matches]
    years_valid = [y for y in years if y <= max_year]
    if not years_valid:
        return None
    # Most common year in first 500 chars (likely publication year)
    return max(set(years_valid), key=years_valid.count)


def extract_doi(text):
    m = re.search(r'10\.\d{4,9}/[-._;()/:A-Z0-9]+', text, re.IGNORECASE)
    return m.group(0) if m else None


def process_pdf(path):
    try:
        doc = fitz.open(path)
    except Exception as e:
        return None, str(e)

    metadata = doc.metadata or {}
    first_page = ''
    try:
        if doc.page_count > 0:
            first_page = doc[0].get_text()
        if doc.page_count > 1 and len(first_page) < 200:
            first_page += '\n' + doc[1].get_text()
    except:
        pass
    doc.close()

    # Limit text
    first_chunk = first_page[:2000]

    year = extract_year_from_text(first_chunk)
    doi = extract_doi(first_chunk)

    # Try to extract title (first non-empty line that's long enough)
    lines = [l.strip() for l in first_chunk.split('\n') if l.strip()]
    title_candidates = []
    for line in lines[:15]:
        # Title heuristic: length 30-200, not all caps, has at least 1 lowercase
        if 20 <= len(line) <= 250 and any(c.islower() for c in line):
            title_candidates.append(line)

    title = title_candidates[0] if title_candidates else (metadata.get('title') or '')

    return {
        'path': path.replace(ROOT, '').replace('\\', '/').lstrip('/'),
        'filename': os.path.basename(path),
        'size_kb': os.path.getsize(path) // 1024,
        'meta_title': metadata.get('title', ''),
        'meta_author': metadata.get('author', ''),
        'extracted_title': title,
        'extracted_year': year,
        'doi': doi,
        'first_500': first_chunk[:500].replace('\n', ' | ')[:500],
    }, None


total = 0
for folder in SCAN_FOLDERS:
    folder_path = os.path.join(ROOT, folder)
    if not os.path.exists(folder_path):
        continue
    for root, dirs, files in os.walk(folder_path):
        for f in files:
            if f.lower().endswith('.pdf'):
                total += 1
                full = os.path.join(root, f)
                result, err = process_pdf(full)
                if err:
                    errors.append({'path': full, 'error': err})
                else:
                    results.append(result)
                if total % 20 == 0:
                    print(f'  Processed {total} PDFs...')


print(f'\nDONE: {len(results)} PDFs, {len(errors)} errors')

# Save catalog
out_json = os.path.join(ROOT, '06_Scripts', 'pdf_catalog_29052026.json')
with open(out_json, 'w', encoding='utf-8') as f:
    json.dump({'pdfs': results, 'errors': errors}, f, ensure_ascii=False, indent=2)
print(f'Saved: {out_json}')

# Quick summary
print(f'\n=== SUMMARY ===')
years = [r['extracted_year'] for r in results if r['extracted_year']]
print(f'PDFs with detected year: {len(years)}/{len(results)}')
if years:
    print(f'Year range: {min(years)} - {max(years)}')
dois = [r for r in results if r['doi']]
print(f'PDFs with DOI: {len(dois)}/{len(results)}')
