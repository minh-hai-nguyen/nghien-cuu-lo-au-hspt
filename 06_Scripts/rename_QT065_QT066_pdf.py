"""Rename + move PDF QT065 + QT066 vào folder canonical + update canonical_index."""
import sys, io, json
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = Path('c:/Users/HLC/OneDrive/read_books/Lo-au')
IDX_PATH = BASE / '02_Papers-goc' / 'canonical_index.json'
F_AU = BASE / '02_Papers-goc' / 'The-gioi_Au-My-Uc'

RENAMES = [
    {
        'cid': 'QT065',
        'src': BASE / '02_Papers-goc' / '1-s2.0-S002244052500010X-main.pdf',
        'dst': F_AU / 'QT065_Bradshaw_Lochman_EACP_RCT_JSchPsy_2025.pdf',
    },
    {
        'cid': 'QT066',
        'src': BASE / '02_Papers-goc' / 'Journal Community Psychology - 2023 - Murphy - A systematic scoping review of peer support interventions in integrated.pdf',
        'dst': F_AU / 'QT066_Murphy_PeerSupport_ScopingReview_JCommPsych_2024.pdf',
    },
]

with open(IDX_PATH, encoding='utf-8') as f:
    idx = json.load(f)

log = []
for r in RENAMES:
    if not r['src'].exists():
        log.append(f"SKIP {r['cid']}: source missing — {r['src'].name}")
        continue
    if r['dst'].exists():
        log.append(f"SKIP {r['cid']}: dst exists")
        continue
    r['src'].rename(r['dst'])
    log.append(f"RENAMED {r['cid']}: → {r['dst'].relative_to(BASE)}")
    # Update index
    idx[r['cid']]['pdf'] = r['dst'].name
    idx[r['cid']]['pdf_folder'] = 'The-gioi_Au-My-Uc'
    idx[r['cid']]['has_pdf'] = True
    idx[r['cid']].pop('pdf_pending_filename', None)
    log.append(f"INDEX {r['cid']}: pdf field updated")

with open(IDX_PATH, 'w', encoding='utf-8') as f:
    json.dump(idx, f, ensure_ascii=False, indent=2)

n_with_pdf = sum(1 for v in idx.values() if v.get('pdf'))
log.append(f'\nTotal canonical: {len(idx)} | có PDF: {n_with_pdf}')
for l in log: print(l)
