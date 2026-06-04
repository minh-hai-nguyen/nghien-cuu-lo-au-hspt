# -*- coding: utf-8 -*-
"""Copy all available PDF originals for papers cited in báo cáo can thiệp 10/04/2026 to a single folder."""
import os, sys, io, shutil, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
OUT_DIR = os.path.join(ROOT, 'Bai_goc_BaoCao_CanThiep_10042026')
os.makedirs(OUT_DIR, exist_ok=True)

with open(os.path.join(ROOT, '02_Papers-goc', 'canonical_index.json'), encoding='utf-8') as f:
    canon = json.load(f)

# Map: Bài số → canonical ID → folder
# (Order matches report's "TÀI LIỆU THAM KHẢO" section)
PAPERS_TO_COPY = [
    # (bài_số, canonical_id, note)
    ('Bai_28', 'QT028', 'Zugman et al. 2024 — AJP Pediatric Anxiety Treatment Review'),
    ('Bai_29', 'QT029', 'Li et al. 2025 — BMC Bayesian NMA CBT anxiety children'),
    ('Bai_41', 'QT037', 'Praptomojati & Hartanto 2024 — SEA CA-CBT systematic review'),
    ('Bai_42', 'QT038', 'De Silva et al. 2024 — Sri Lanka cluster RCT'),
    ('Bai_43', 'QT039', 'Xian, Zhang & Jiang 2024 — NMA SAD children/adolescents'),
    ('Bai_44', 'QT040', 'Walder et al. 2025 — JMIR DMHI for SAD'),
    ('Bai_48', 'QT042', 'Brown & Carter 2025 — BESST UK editorial'),
    ('Bai_49', 'QT043', 'Bress et al. 2024 — Maya app JAMA Network Open'),
    ('Bai_50', 'QT044', 'Cao / Cai et al. 2025 — Frontiers Resilience SR+MA'),
    ('Bai_51', 'QT045', 'Sasaki et al. 2024 — Japan iCBT JMIR'),
    ('Bai_56', 'QT049', 'Zhang, Liang & Kang 2026 — Bayesian MA universal CBT (ABSTRACT only)'),
    ('Bai_57', 'QT050', 'Qiaochu & Wang 2025 — Mobile CBT (ABSTRACT only)'),
    ('Bai_58', 'QT051', 'Menon et al. 2025 — LMIC Scoping (ABSTRACT only)'),
    ('Bai_54', 'QT047', 'Dong, Wang & Lin 2025 — YaAn PLOS ONE (comparison)'),
    ('Bai_55', 'QT048', 'Chen et al. 2025 — COVID meta (ABSTRACT only, comparison)'),
]

# Special cases
SPECIAL = [
    ('VN031_TranNguyenNgoc_2018_LuanAn_TSYH_BachMai',
     '02_Papers-goc/Viet-Nam/TranNguyenNgoc_2018_LuanAn_TSYH_BachMai.pdf',
     'Trần Nguyễn Ngọc 2018 — Luận án TSYH Bạch Mai (bài VN duy nhất, chưa canonical hoá — đề xuất VN031)'),
]

copied = []
skipped = []
missing = []

for order, cid, note in PAPERS_TO_COPY:
    meta = canon.get(cid, {})
    pdf = meta.get('pdf')
    folder = meta.get('pdf_folder')
    if pdf and folder and folder != 'None':
        src = os.path.join(ROOT, '02_Papers-goc', folder, pdf)
        if os.path.exists(src):
            # Prepend order prefix to filename for easy sorting
            dst_name = f'{order}__{pdf}'
            dst = os.path.join(OUT_DIR, dst_name)
            shutil.copy2(src, dst)
            copied.append((order, cid, dst_name, os.path.getsize(dst), note))
            print(f'  ✓ {order}: {cid} → {dst_name}')
        else:
            missing.append((order, cid, f'Source not found: {src}', note))
            print(f'  ✗ {order}: {cid} source missing at {src}')
    else:
        skipped.append((order, cid, 'No PDF (abstract-only or missing)', note))
        print(f'  ⚠ {order}: {cid} — {note}')

# Copy special cases
for order, src, note in SPECIAL:
    src_full = os.path.join(ROOT, src)
    if os.path.exists(src_full):
        dst_name = f'{order}.pdf'
        dst = os.path.join(OUT_DIR, dst_name)
        shutil.copy2(src_full, dst)
        copied.append((order, '-', dst_name, os.path.getsize(dst), note))
        print(f'  ✓ {order}: → {dst_name}')

print()
print('='*70)
print(f'COPIED: {len(copied)} PDFs')
print(f'ABSTRACT-ONLY / MISSING: {len(skipped) + len(missing)}')
print(f'Output folder: {OUT_DIR}')

# Calculate total size
total_bytes = sum(s for _, _, _, s, _ in copied)
print(f'Total size: {total_bytes/1024/1024:.1f} MB')
