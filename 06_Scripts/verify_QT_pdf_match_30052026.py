# -*- coding: utf-8 -*-
"""Buoc 1: Verify ALL QT*.pdf filename khop content (doc title trang 1).
Phat hien hom qua: QT002_Saikia_2023.pdf thuc te la Bhardwaj 2020 content.
Can check toan bo de tim cac case mislabel khac.
30/05/2026."""
import os, sys, io, re, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import fitz

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

# Find all QT*.pdf files
qt_pdfs = []
for root, dirs, files in os.walk(os.path.join(ROOT, '02_Papers-goc')):
    for f in files:
        if f.startswith('QT') and f.lower().endswith('.pdf'):
            qt_pdfs.append(os.path.join(root, f))

print(f'Tổng QT PDFs: {len(qt_pdfs)}')

mismatches = []
ok = []
errors = []

for pdf_path in sorted(qt_pdfs):
    fn = os.path.basename(pdf_path)
    # Extract expected author + year from filename
    # Pattern: QT###_AuthorName_YYYY_...
    m = re.match(r'QT(\d{3})_([A-Z][a-zA-Z]+(?:_et_al)?)_(\d{4})_', fn)
    if not m:
        # Try alternate patterns
        m2 = re.match(r'QT(\d{3})_([A-Z][a-zA-Z]+).*?_(\d{4})', fn)
        if m2:
            qt_num, expected_author, expected_year = m2.group(1), m2.group(2), m2.group(3)
        else:
            errors.append(f'{fn}: filename pattern unparseable')
            continue
    else:
        qt_num, expected_author, expected_year = m.group(1), m.group(2), m.group(3)

    expected_author = expected_author.replace('_et_al', '').replace('_', ' ')

    try:
        doc = fitz.open(pdf_path)
        text_p1 = doc[0].get_text()[:2000]
        doc.close()
    except Exception as e:
        errors.append(f'{fn}: cannot read PDF - {e}')
        continue

    # Check if expected author appears in first 2000 chars
    author_found = expected_author.lower() in text_p1.lower()
    # Check if expected year appears
    year_found = expected_year in text_p1

    if author_found and year_found:
        ok.append(qt_num)
    elif not author_found:
        # Extract what authors ARE in the PDF
        # Try to find author pattern in title area
        # Look for first 5 capitalized words
        first_caps = re.findall(r'\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\b', text_p1[:500])
        likely_authors = [w for w in first_caps if len(w) > 3 and w not in
                          ['The', 'This', 'These', 'Those', 'Article', 'Research', 'Open', 'Access',
                           'Original', 'Abstract', 'Journal', 'Author', 'Background', 'Methods',
                           'Results', 'Conclusion', 'Department', 'Faculty', 'School', 'University']][:5]
        mismatches.append({
            'qt': qt_num,
            'filename': fn,
            'expected_author': expected_author,
            'expected_year': expected_year,
            'year_in_pdf': year_found,
            'likely_pdf_authors': likely_authors,
            'first_300': text_p1[:300].replace('\n', ' ').strip(),
        })

print()
print(f'✓ OK (filename matches content): {len(ok)}')
print(f'✗ MISMATCH (filename != content): {len(mismatches)}')
print(f'? ERRORS (cannot verify): {len(errors)}')
print()

if mismatches:
    print('=== MISMATCHES ===')
    for m in mismatches:
        print(f'\n[QT{m["qt"]}] {m["filename"]}')
        print(f'  Expected author: {m["expected_author"]}, year: {m["expected_year"]} (year_in_pdf: {m["year_in_pdf"]})')
        print(f'  Likely PDF authors: {m["likely_pdf_authors"]}')
        print(f'  First 300 chars: {m["first_300"]}')

if errors:
    print()
    print('=== ERRORS ===')
    for e in errors[:10]:
        print(f'  {e}')

# Save report
out = {
    'total': len(qt_pdfs),
    'ok_count': len(ok),
    'mismatches': mismatches,
    'errors': errors,
}
with open(os.path.join(ROOT, '06_Scripts', 'QT_pdf_verify_30052026.json'),
          'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
print()
print(f'Saved report: 06_Scripts/QT_pdf_verify_30052026.json')
