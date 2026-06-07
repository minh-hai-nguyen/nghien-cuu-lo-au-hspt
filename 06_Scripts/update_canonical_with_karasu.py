# -*- coding: utf-8 -*-
"""Them Karasu 1986 + Karasu bio sketch vao canonical_index.json"""
import sys, io, json, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
PATH = os.path.join(ROOT, '02_Papers-goc', 'canonical_index.json')

with open(PATH, 'r', encoding='utf-8') as f:
    idx = json.load(f)

# Find next QT id
qt_ids = [k for k in idx if k.startswith('QT')]
next_qt = max(int(q[2:]) for q in qt_ids) + 1
print(f'Next QT id: QT{next_qt:03d}')

# Add Karasu entries
new_entries = {
    f'QT{next_qt:03d}': {
        'id': f'QT{next_qt:03d}',
        'descriptor': 'Karasu_1986_USA_Psychotherapies',
        'summary': None,
        'translation': None,
        'pdf': 'karasu1986.pdf',
        'pdf_path': '02_Papers-goc/karasu1986.pdf',
        'authors': ['T. Byram Karasu'],
        'title': 'The Psychotherapies: Benefits and Limitations',
        'journal': 'American Journal of Psychotherapy',
        'volume': '40', 'issue': '3', 'pages': '324-342', 'year': 1986,
        'pmid': '3094389',
        'doi': '10.1176/appi.psychotherapy.1986.40.3.324',
        'topic': 'Psychotherapy schools / Number of approaches',
        'key_finding': 'over 450 types polled nationwide (page 325)',
        'used_for': 'Verify_BaoVeLA Q&A 2 (số trường phái tâm lý trị liệu)',
        'added': '2026-06-07',
    },
    f'QT{next_qt+1:03d}': {
        'id': f'QT{next_qt+1:03d}',
        'descriptor': 'Karasu_BioSketch_Einstein_2014',
        'summary': None,
        'translation': None,
        'pdf': 'biosketch_karasu_062112.pdf',
        'pdf_path': '02_Papers-goc/biosketch_karasu_062112.pdf',
        'authors': ['T. Byram Karasu'],
        'title': 'Biographical sketch (T. Byram Karasu, M.D.)',
        'source': 'Albert Einstein College of Medicine official bio',
        'date': '2014-02-14',
        'topic': 'Author credentials verification',
        'key_facts': '21 books, 100+ papers, Distinguished Life Fellow APA, '
                     'Editor-in-Chief emeritus AJ Psychotherapy, '
                     '400+ scholars task force chair (1981)',
        'used_for': 'Verify_BaoVeLA Karasu authority confirmation',
        'added': '2026-06-07',
    },
}

idx.update(new_entries)

# Backup before write
import shutil
shutil.copy(PATH, PATH + '.bak_07062026')

with open(PATH, 'w', encoding='utf-8') as f:
    json.dump(idx, f, ensure_ascii=False, indent=2)

print(f'Updated: {PATH}')
print(f'  Added 2 entries: QT{next_qt:03d} (Karasu 1986), '
      f'QT{next_qt+1:03d} (Karasu bio sketch)')
print(f'  Total entries: {len(idx)}')
print(f'  Backup: {PATH}.bak_07062026')
