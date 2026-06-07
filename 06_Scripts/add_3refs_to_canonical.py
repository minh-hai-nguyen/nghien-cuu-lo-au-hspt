# -*- coding: utf-8 -*-
"""Add Small&Blanc 2021 (downloaded) + Stankov 2010 (placeholder, awaiting PDF)
+ Rose 2002 (placeholder, awaiting PDF) to canonical_index."""
import sys, io, json, os, shutil
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
PATH = os.path.join(ROOT, '02_Papers-goc', 'canonical_index.json')

with open(PATH, 'r', encoding='utf-8') as f:
    idx = json.load(f)

# Find next QT id (excluding BB)
qt_ids = [k for k in idx if k.startswith('QT')]
next_qt = max(int(q[2:]) for q in qt_ids) + 1
print(f'Next QT id: QT{next_qt:03d}')

new_entries = {
    f'QT{next_qt:03d}': {
        'id': f'QT{next_qt:03d}',
        'descriptor': 'Small_Blanc_2021_USA_TamGiao_VN',
        'summary': None,
        'translation': None,
        'pdf': 'Small_Blanc_2021_TamGiao_Vietnam.pdf',
        'pdf_path': '02_Papers-goc/Chua-phan-loai/'
                    'Small_Blanc_2021_TamGiao_Vietnam.pdf',
        'authors': ['Stuart Small', 'Judite Blanc'],
        'title': 'Mental Health During COVID-19: Tam Giao and Vietnam\'s Response',
        'journal': 'Frontiers in Psychiatry',
        'volume': '11', 'pages': '589618', 'year': 2021,
        'doi': '10.3389/fpsyt.2020.589618',
        'pmc': 'PMC7820702',
        'topic': 'Tam Giao Vietnam cultural framework',
        'used_for': 'Q1 Section 4.4 — peer support / cultural restraint',
        'verification_status': 'PMC_FULLTEXT_VERIFIED + PDF downloaded',
        'added': '2026-06-07',
    },
    f'QT{next_qt+1:03d}': {
        'id': f'QT{next_qt+1:03d}',
        'descriptor': 'Stankov_2010_Australia_ConfucianAcademic',
        'summary': None,
        'translation': None,
        'pdf': None,
        'pdf_path': '02_Papers-goc/Chua-phan-loai/'
                    'Stankov_2010_Confucian_Academic.pdf (CHỜ THẦY TẢI)',
        'authors': ['Lazar Stankov'],
        'title': 'Unforgiving Confucian culture: A breeding ground for high '
                 'academic achievement, test anxiety and self-doubt?',
        'journal': 'Learning and Individual Differences',
        'volume': '20', 'issue': '6', 'pages': '555-563', 'year': 2010,
        'doi': '10.1016/j.lindif.2010.05.003',
        'eric_id': 'EJ905763',
        'topic': 'Confucian academic culture; test anxiety',
        'used_for': 'Q1 Section 1.1 — Vietnamese context',
        'verification_status': 'CITATION_VERIFIED via ERIC; PDF chờ thầy tải '
                               '(paywall ScienceDirect)',
        'download_hints': [
            'https://www.researchgate.net/publication/222426691',
            'https://www.academia.edu/27969005',
            'https://doi.org/10.1016/j.lindif.2010.05.003',
        ],
        'added': '2026-06-07',
    },
    f'QT{next_qt+2:03d}': {
        'id': f'QT{next_qt+2:03d}',
        'descriptor': 'Rose_2002_USA_CoRumination',
        'summary': None,
        'translation': None,
        'pdf': None,
        'pdf_path': '02_Papers-goc/Chua-phan-loai/Rose_2002_Co-rumination.pdf '
                    '(CHỜ THẦY TẢI)',
        'authors': ['Amanda J. Rose'],
        'title': 'Co-rumination in the friendships of girls and boys',
        'journal': 'Child Development',
        'volume': '73', 'issue': '6', 'pages': '1830-1843', 'year': 2002,
        'pmid': '12487497',
        'doi': '10.1111/1467-8624.00509',
        'topic': 'Co-rumination hypothesis; peer disclosure; '
                 'depression/anxiety symptoms',
        'used_for': 'Q1 Section 4.4 — co-rumination explanation peer support null',
        'verification_status': 'PUBMED_VERIFIED; PDF chờ thầy tải (paywall Wiley)',
        'download_hints': [
            'https://www.researchgate.net/publication/11150944',
            'https://doi.org/10.1111/1467-8624.00509',
            'https://psychology.missouri.edu/people/rose',
        ],
        'added': '2026-06-07',
    },
}

# Backup
shutil.copy(PATH, PATH + '.bak_3refs_07062026')

idx.update(new_entries)
with open(PATH, 'w', encoding='utf-8') as f:
    json.dump(idx, f, ensure_ascii=False, indent=2)

print(f'Added 3 entries: QT{next_qt:03d}-QT{next_qt+2:03d}')
print(f'Total canonical entries: {len(idx)}')
print(f'  QT{next_qt:03d}: Small & Blanc 2021 (PDF có)')
print(f'  QT{next_qt+1:03d}: Stankov 2010 (PDF chờ)')
print(f'  QT{next_qt+2:03d}: Rose 2002 (PDF chờ)')
