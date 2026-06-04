# -*- coding: utf-8 -*-
"""
Phase B — Normalize + dedupe author names.

Input: authors_raw.json (Phase A + A2 output)
Output: authors_normalized.json với unique author IDs
"""
import os, sys, io, re, json, unicodedata
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
OUT_DIR = os.path.join(os.path.dirname(__file__), 'author_kg_data')

with open(os.path.join(OUT_DIR, 'authors_raw.json'), encoding='utf-8') as f:
    data = json.load(f)

# Known alias mapping (historical errors + conventional equivalents)
KNOWN_ALIASES = {
    # QT045: Sasaki → Matsumoto (memory feedback_QT045_author_correction.md)
    'sasaki': None,  # exclude — this is a historical error, never use
    # Vietnamese name ordering (surname-first VN convention)
    'dang hoang minh': 'AU_DANG_HM',
    'hoang minh dang': 'AU_DANG_HM',
    'hoang m.d.': 'AU_DANG_HM',
    'dang hoang m.': 'AU_DANG_HM',
    'h.m. dang': 'AU_DANG_HM',
    'đặng hoàng minh': 'AU_DANG_HM',
    # Other common
    'hoa pt': 'AU_HOA_PT',
    'pham thi thu hoa': 'AU_HOA_PT',
    'phạm thị thu hoa': 'AU_HOA_PT',
}

def nfc(s):
    """Unicode normalize NFC."""
    return unicodedata.normalize('NFC', s)

def strip_diacritics(s):
    """Remove Vietnamese diacritics for matching."""
    s = nfc(s)
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = s.replace('đ', 'd').replace('Đ', 'D')
    return s

def is_vietnamese_name(name):
    """Heuristic: Vietnamese names contain Vietnamese diacritics or common VN surnames."""
    vn_surnames = {'nguyễn', 'trần', 'lê', 'phạm', 'hoàng', 'phan', 'vũ', 'võ', 'đặng',
                   'bùi', 'đỗ', 'hồ', 'ngô', 'dương', 'lý', 'đinh', 'cao', 'lưu',
                   'trương', 'đoàn', 'đào', 'thái', 'khuất', 'tạ', 'mai', 'tô',
                   # Non-diacritic variants
                   'nguyen', 'tran', 'le', 'pham', 'hoang', 'phan', 'vu', 'vo', 'dang',
                   'bui', 'do', 'ho', 'ngo', 'duong', 'ly', 'dinh', 'cao', 'luu',
                   'truong', 'doan', 'dao', 'thai', 'khuat', 'ta', 'mai', 'to'}
    words_lower = [w.lower().strip('.,') for w in name.split()]
    if words_lower and words_lower[0] in vn_surnames:
        return True
    if re.search(r'[àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđĐ]', name, re.IGNORECASE):
        return True
    return False

def parse_name(raw_name):
    """Parse a single author name into components.
    Returns dict {full_name, surname, given, initials, is_vietnamese}"""
    name = nfc(raw_name.strip())
    name = re.sub(r'\s+', ' ', name)
    name = name.strip('.,;:')

    # Remove superscript digits (affiliation markers)
    name = re.sub(r'[\d\u00b9\u00b2\u00b3\u2070-\u2079*†‡§¶]+', '', name).strip()

    # Check for "LastName, FirstName" format (EN academic)
    if ',' in name and name.count(',') == 1:
        parts = [p.strip() for p in name.split(',')]
        if len(parts[0].split()) <= 2 and len(parts[1].split()) <= 4:
            # "Lastname, Firstname Middle"
            surname = parts[0]
            given = parts[1]
            return {
                'full_name': f'{given} {surname}',
                'surname': surname,
                'given': given,
                'initials': _initials(given),
                'is_vietnamese': False,
            }

    is_vn = is_vietnamese_name(name)

    words = name.split()
    if len(words) < 2:
        return {
            'full_name': name,
            'surname': name,
            'given': '',
            'initials': '',
            'is_vietnamese': is_vn,
        }

    if is_vn:
        # VN convention: surname first
        surname = words[0]
        given = ' '.join(words[1:])
    else:
        # EN convention: last word = surname (unless "Jr." or "III")
        if words[-1].rstrip('.') in ('Jr', 'Sr', 'II', 'III', 'IV'):
            surname = words[-2]
            given = ' '.join(words[:-2]) if len(words) > 2 else ''
        else:
            surname = words[-1]
            given = ' '.join(words[:-1])

    return {
        'full_name': name,
        'surname': surname,
        'given': given,
        'initials': _initials(given),
        'is_vietnamese': is_vn,
    }

def _initials(given):
    """Extract initials from given name."""
    if not given: return ''
    # Remove existing dots
    parts = re.split(r'[\s.]+', given)
    parts = [p for p in parts if p]
    return ''.join(p[0].upper() for p in parts if p[0].isalpha())

def make_author_id(parsed):
    """Generate canonical author ID: AU_SURNAME_INITIALS."""
    surname_ascii = strip_diacritics(parsed['surname']).upper()
    surname_ascii = re.sub(r'[^A-Z]', '', surname_ascii)
    initials = parsed['initials'] or 'X'
    return f'AU_{surname_ascii}_{initials}'

# Main: iterate through all papers, normalize names, dedupe
authors = {}  # author_id → author record
paper_author_edges = []  # (paper_id, author_id, position)
aliases_log = {}  # surname+initials → list of raw variants

for paper_id, info in data.items():
    author_list = info.get('authors_list', [])
    for entry in author_list:
        raw = entry['name']
        position = entry['position']

        # Apply known aliases first
        alias_key = raw.strip().lower()
        if alias_key in KNOWN_ALIASES:
            target_id = KNOWN_ALIASES[alias_key]
            if target_id is None:
                # Excluded alias (e.g., Sasaki for QT045)
                print(f'  [SKIP SASAKI] Paper {paper_id}: raw name "{raw}" excluded (historical error)')
                continue

        parsed = parse_name(raw)

        # Skip garbage
        if not parsed['surname'] or len(parsed['surname']) < 2:
            continue
        # Skip common noise words that slipped through
        if parsed['surname'].lower() in ('and', 'et', 'al', 'the', 'of', 'in', 'co', 'for'):
            continue

        # Check aliases
        if alias_key in KNOWN_ALIASES and KNOWN_ALIASES[alias_key]:
            author_id = KNOWN_ALIASES[alias_key]
        else:
            author_id = make_author_id(parsed)

        # Create or update author record
        if author_id not in authors:
            authors[author_id] = {
                'canonical_id': author_id,
                'surname': parsed['surname'],
                'given': parsed['given'],
                'full_name': parsed['full_name'],
                'is_vietnamese': parsed['is_vietnamese'],
                'initials': parsed['initials'],
                'aliases': set(),
                'papers': set(),
                'role_in_papers': {},
                'country_hint': info.get('country_hint', ''),
            }
        authors[author_id]['aliases'].add(raw)
        authors[author_id]['papers'].add(paper_id)
        authors[author_id]['role_in_papers'][paper_id] = position

        paper_author_edges.append({
            'paper_id': paper_id,
            'author_id': author_id,
            'position': position,
            'is_first_author': (position == 1),
            'raw_name': raw,
        })

# Convert sets to lists for JSON
for aid, rec in authors.items():
    rec['aliases'] = sorted(rec['aliases'])
    rec['papers'] = sorted(rec['papers'])
    rec['n_papers'] = len(rec['papers'])

# Stats
n_authors = len(authors)
multi_paper_authors = [a for a in authors.values() if a['n_papers'] >= 2]
top_authors = sorted(authors.values(), key=lambda a: -a['n_papers'])[:20]

print('='*70)
print('PHASE B — NORMALIZE + DEDUPE')
print('='*70)
print(f'Unique authors: {n_authors}')
print(f'Multi-paper authors (≥ 2 papers): {len(multi_paper_authors)}')
print(f'Paper-author edges: {len(paper_author_edges)}')
print()
print('Top 20 authors by paper count:')
for a in top_authors:
    print(f'  {a["canonical_id"]:25s} ({a["full_name"][:40]:40s}) — {a["n_papers"]} papers: {a["papers"]}')

# Check for remaining Sasaki
sasaki_check = [aid for aid in authors if 'SASAKI' in aid]
if sasaki_check:
    print(f'\n⚠️ WARNING: Sasaki still in authors! {sasaki_check}')
else:
    print('\n✓ Sasaki NOT in authors (rule QT045 satisfied)')

# Save
out_authors = os.path.join(OUT_DIR, 'authors_normalized.json')
with open(out_authors, 'w', encoding='utf-8') as f:
    json.dump(authors, f, ensure_ascii=False, indent=2)

out_edges = os.path.join(OUT_DIR, 'paper_author_edges.json')
with open(out_edges, 'w', encoding='utf-8') as f:
    json.dump(paper_author_edges, f, ensure_ascii=False, indent=2)

print(f'\nSaved: {out_authors}')
print(f'Saved: {out_edges}')
