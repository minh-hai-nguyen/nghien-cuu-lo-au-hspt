# -*- coding: utf-8 -*-
"""
Phase A2 — PDF fallback cho 23 papers thiếu author + cleanup noisy entries.

Use pypdf to extract first page → search for author block.
"""
import os, sys, io, re, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    from pypdf import PdfReader
except ImportError:
    from PyPDF2 import PdfReader

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
OUT_DIR = os.path.join(os.path.dirname(__file__), 'author_kg_data')

with open(os.path.join(OUT_DIR, 'authors_raw.json'), encoding='utf-8') as f:
    data = json.load(f)

with open(os.path.join(ROOT, '02_Papers-goc', 'canonical_index.json'), encoding='utf-8') as f:
    canon = json.load(f)

def pdf_first_page_text(pdf_path, pages_to_read=2):
    if not os.path.exists(pdf_path): return ''
    try:
        reader = PdfReader(pdf_path)
        txt = ''
        for i in range(min(pages_to_read, len(reader.pages))):
            t = reader.pages[i].extract_text() or ''
            txt += t + '\n'
        return txt
    except Exception as e:
        return ''

def extract_authors_from_pdf_text(pdf_text, title_hint=''):
    """Extract author block from PDF first page text.
    Strategy: look between title and 'Abstract'/'Background'/'Introduction'."""
    if not pdf_text: return ''

    # Find abstract/intro marker
    m = re.search(r'\b(Abstract|ABSTRACT|Background|Summary|Correspondence)\b', pdf_text)
    if not m:
        return ''

    pre_abstract = pdf_text[:m.start()]

    # Remove title hint if present
    if title_hint:
        # Title is usually longest sentence in first 500 chars
        pass

    # Find author-like patterns: lines with multiple "Firstname LastName" patterns
    lines = pre_abstract.split('\n')
    author_lines = []
    for line in lines:
        line = line.strip()
        if len(line) < 10 or len(line) > 500:
            continue
        # Skip lines that look like affiliation (contain "university", "school", etc.)
        if re.search(r'\b(university|institute|department|college|school|hospital|laboratory|center|faculty)\b', line, re.IGNORECASE):
            continue
        # Count capitalized words (potential author names)
        cap_words = re.findall(r'\b[A-Z][a-zA-Z\-]+(?:[.,]\s*)?', line)
        if len(cap_words) >= 3 and re.search(r',|\sand\s|;|\s&\s', line):
            author_lines.append(line)

    if author_lines:
        # Take longest author line (most likely full author list)
        return max(author_lines, key=len)
    return ''

def parse_author_list(raw):
    if not raw: return [], False
    raw = raw.strip()
    raw = re.sub(r',?\s*(?:et al\.?|và cộng sự|và cs\.?)\b.*', ' [et al.]', raw, flags=re.IGNORECASE)
    has_etal = '[et al.]' in raw
    raw = raw.replace(' [et al.]', '').strip()
    raw = raw.rstrip('.,;:')

    # Remove affiliation superscripts (numbers or letters)
    raw = re.sub(r'[\d\u00b9\u00b2\u00b3\u2070-\u2079]+\s*,?', '', raw)

    parts = re.split(r'\s*(?:,|&|;|\s+và\s+|\s+and\s+)\s*', raw)
    parts = [p.strip() for p in parts if p.strip() and len(p.strip()) > 2]

    authors = []
    for i, p in enumerate(parts):
        if re.match(r'^\d+', p): continue
        if len(p) < 3: continue
        if not re.search(r'[a-zA-ZÀ-ỹ]', p): continue
        # Filter out strings that look like affiliation fragments
        if re.search(r'(?i)\b(of|the|and|in|at)\b$', p): continue
        authors.append({
            'name': p.strip('.'),
            'position': i + 1,
            'is_corresponding': (i == 0),
        })
    return authors, has_etal

# Identify empty papers + noisy papers (authors_list < 2 but descriptor suggests multi-author)
empty_cids = []
for cid, r in data.items():
    n = r['n_authors_extracted']
    raw = r['authors_raw']
    # Check if noisy: contains "giải thích" or "thiếu" or "nhưng" (sentences, not names)
    is_noisy = raw and any(marker in raw.lower() for marker in [
        'giải thích', 'thiếu', 'nhưng', 'bao phủ', 'nhóm', 'được', 'tại',
        'không giải thích', 'chưa', 'the authors', 'theo', 'cho', 'với', 'bộ'
    ])
    if n == 0 or (n == 1 and not r['has_etal']) or is_noisy:
        empty_cids.append(cid)

print(f'Papers needing PDF fallback or cleanup: {len(empty_cids)}')
print(f'List: {empty_cids}')

updated = 0
for cid in empty_cids:
    meta = canon.get(cid, {})
    pdf_fn = meta.get('pdf')
    folder = meta.get('pdf_folder')
    if not pdf_fn or not folder or folder == 'None':
        continue
    pdf_path = os.path.join(ROOT, '02_Papers-goc', folder, pdf_fn)
    if not os.path.exists(pdf_path):
        continue
    pdf_text = pdf_first_page_text(pdf_path, pages_to_read=1)
    raw = extract_authors_from_pdf_text(pdf_text)
    if raw:
        authors_list, has_etal = parse_author_list(raw)
        if authors_list:
            data[cid]['authors_raw_pdf'] = raw[:300]
            data[cid]['authors_list_pdf'] = authors_list
            data[cid]['has_etal_pdf'] = has_etal
            data[cid]['n_authors_pdf'] = len(authors_list)
            # Merge: prefer PDF extraction if original was noisy/empty
            if not data[cid]['authors_list'] or len(data[cid]['authors_list']) < 2:
                data[cid]['authors_list'] = authors_list
                data[cid]['n_authors_extracted'] = len(authors_list)
                data[cid]['has_etal'] = has_etal
                data[cid]['authors_raw'] = raw[:300]
                data[cid]['source'] = 'pdf_fallback'
                updated += 1
                print(f'  PDF FALLBACK {cid}: {len(authors_list)} authors — {raw[:80]}')

print(f'\nUpdated via PDF fallback: {updated}')

# Save
with open(os.path.join(OUT_DIR, 'authors_raw.json'), 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

# Stats
total = len(data)
with_auth = sum(1 for r in data.values() if r['n_authors_extracted'] >= 2)
empty = sum(1 for r in data.values() if r['n_authors_extracted'] == 0)
print(f'\nFinal: {with_auth}/{total} papers with ≥2 authors extracted, {empty} still empty')
