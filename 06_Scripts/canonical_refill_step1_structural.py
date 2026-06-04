"""
STEP 1: Structural changes - xóa duplicate + rename 4 PDF cho canonical đã có thiếu PDF.
Safe/reversible: backup canonical_index.json trước khi edit.
"""
import sys, io, os, json, shutil
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = Path('c:/Users/HLC/OneDrive/read_books/Lo-au')
IDX_PATH = BASE / '02_Papers-goc' / 'canonical_index.json'

# === NHÓM A: Duplicate — xóa ===
to_delete = BASE / '02_Papers-goc' / 'The-gioi_Au-My-Uc' / 'bress_2024_oi_240871_1723229502.27231 (1).pdf'

# === NHÓM B: Rename 4 PDF cho QT006/007/009/011 ===
# (từ, đến, canonical_id, folder_name)
renames = [
    (BASE / '02_Papers-goc/The-gioi_Khac/Nakie_2022_Ethiopia_HighSchool_BMCPsych.pdf',
     BASE / '02_Papers-goc/The-gioi_Khac/QT006_Nakie_et_al_2022_Ethiopia.pdf',
     'QT006',
     'The-gioi_Khac'),
    (BASE / '02_Papers-goc/The-gioi_Khac/Chen_2023_Chinese_SecondarySchool_BMCPsych.pdf',
     BASE / '02_Papers-goc/The-gioi_Khac/QT007_Chen_et_al_2023_China_BMCPsych.pdf',
     'QT007',
     'The-gioi_Khac'),
    (BASE / '02_Papers-goc/The-gioi_Khac/Qiu_2022_ParentingStyle_Resilience_FrontPsych.pdf',
     BASE / '02_Papers-goc/The-gioi_Khac/QT009_Qiu_et_al_2022_China_Parenting.pdf',
     'QT009',
     'The-gioi_Khac'),
    (BASE / '02_Papers-goc/The-gioi_Khac/Bhardwaj_2020_Chandigarh_DASS21.pdf',
     BASE / '02_Papers-goc/The-gioi_Khac/QT011_Bhardwaj_et_al_2020_India_Chandigarh.pdf',
     'QT011',
     'The-gioi_Khac'),
]

# === Execute ===
log = []

# Backup index
bkp = IDX_PATH.with_suffix('.json.backup')
shutil.copy(IDX_PATH, bkp)
log.append(f'Backup: {bkp.name}')

with open(IDX_PATH, encoding='utf-8') as f:
    idx = json.load(f)

# Delete duplicate
if to_delete.exists():
    to_delete.unlink()
    log.append(f'DELETED duplicate: {to_delete.name}')
else:
    log.append(f'SKIP delete (not found): {to_delete.name}')

# Rename + update index
for src, dst, cid, fold in renames:
    if not src.exists():
        log.append(f'SKIP {cid}: source not found {src.name}')
        continue
    if dst.exists():
        log.append(f'SKIP {cid}: destination already exists {dst.name}')
        continue
    src.rename(dst)
    # Update canonical_index
    idx[cid]['pdf'] = dst.name
    idx[cid]['pdf_folder'] = fold
    log.append(f'RENAMED {cid}: {src.name} → {dst.name}')

with open(IDX_PATH, 'w', encoding='utf-8') as f:
    json.dump(idx, f, ensure_ascii=False, indent=2)
n_with_pdf = sum(1 for v in idx.values() if v.get('pdf'))
log.append(f'UPDATED canonical_index.json (now {n_with_pdf} canonical entries have PDF)')

for l in log:
    print(l)
