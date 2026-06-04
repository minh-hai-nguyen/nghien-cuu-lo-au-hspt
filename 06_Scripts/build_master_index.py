# -*- coding: utf-8 -*-
"""Tao MASTER INDEX voi canonical names + update DANH_SACH."""
import os, sys, io, json, re
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
os.chdir(ROOT)

with open('02_Papers-goc/canonical_index.json', encoding='utf-8') as f:
    idx = json.load(f)

# Build MASTER INDEX markdown
lines = []
lines.append(f'# DANH SÁCH CHÍNH — {len(idx)} bài nghiên cứu dự án Lo âu\n')
lines.append('**Cập nhật:** 12/04/2026 | **Chuẩn hoá tên:** VN001-VN030 + QT001-QT051 (3-digit padding)\n')
n_vn_total = sum(1 for k in idx if k.startswith('VN'))
n_qt_total = sum(1 for k in idx if k.startswith('QT'))
n_sum = sum(1 for v in idx.values() if v.get('summary'))
n_trans = sum(1 for v in idx.values() if v.get('translation'))
n_pdf = sum(1 for v in idx.values() if v.get('pdf'))
lines.append(f'**Tổng:** {len(idx)} bài ({n_vn_total} VN + {n_qt_total} QT) — '
             f'{n_sum} tóm tắt + {n_trans} bản dịch + {n_pdf} PDF gốc\n\n')
lines.append('> Tên file canonical: `{VN|QT}{3-digit}_{Descriptor}.docx/pdf`\n')
lines.append('> Ví dụ: `VN001_Hoa_2024_Frontiers_HaNoi.docx`, `QT043_Bress_JAMA_Maya_App_2024.pdf`\n\n')
lines.append('---\n\n')

# Group: VN, QT
vn_items = sorted([v for k, v in idx.items() if k.startswith('VN')], key=lambda x: x['id'])
qt_items = sorted([v for k, v in idx.items() if k.startswith('QT')], key=lambda x: x['id'])

lines.append(f'## PHẦN I — NGHIÊN CỨU VIỆT NAM ({n_vn_total} bài)\n\n')
lines.append('| ID | Tác giả — chủ đề — năm | Tóm tắt | Bản dịch | PDF gốc |\n')
lines.append('|----|------------------------|---------|----------|--------|\n')
for v in vn_items:
    s = '✓' if v['summary'] else '—'
    t = '✓' if v['translation'] else '—'
    p = '✓' if v['pdf'] else '—'
    folder = f' ({v["pdf_folder"]})' if v['pdf_folder'] else ''
    lines.append(f'| {v["id"]} | {v["descriptor"].replace("_", " ")} | {s} | {t} | {p}{folder} |\n')

lines.append(f'\n## PHẦN II — NGHIÊN CỨU QUỐC TẾ ({n_qt_total} bài)\n\n')
lines.append('| ID | Tác giả — chủ đề — năm | Tóm tắt | Bản dịch | PDF gốc |\n')
lines.append('|----|------------------------|---------|----------|--------|\n')
for v in qt_items:
    s = '✓' if v['summary'] else '—'
    t = '✓' if v['translation'] else '—'
    p = '✓' if v['pdf'] else '—'
    folder = f' ({v["pdf_folder"]})' if v['pdf_folder'] else ''
    lines.append(f'| {v["id"]} | {v["descriptor"].replace("_", " ")} | {s} | {t} | {p}{folder} |\n')

lines.append('\n---\n\n')
lines.append('## THỐNG KÊ\n\n')
n_vn = len(vn_items)
n_qt = len(qt_items)
n_sum = sum(1 for v in idx.values() if v['summary'])
n_trans = sum(1 for v in idx.values() if v['translation'])
n_pdf = sum(1 for v in idx.values() if v['pdf'])
lines.append(f'- **Tổng số bài:** {len(idx)} (VN: {n_vn}, QT: {n_qt})\n')
lines.append(f'- **Có tóm tắt:** {n_sum}/{len(idx)}\n')
lines.append(f'- **Có bản dịch:** {n_trans}/{len(idx)}\n')
lines.append(f'- **Có PDF gốc:** {n_pdf}/{len(idx)}\n\n')

lines.append('## BÀI CÒN THIẾU THÀNH PHẦN\n\n')
lines.append('### Thiếu tóm tắt (6 bài)\n')
lines.append('- QT012–QT017 là 6 bài chuyển tiếp (GBD ASEAN, Zhameden, Anderson, Zhu, Mudunna, Puyat) — có bản dịch nhưng chưa có tóm tắt riêng trong `Tom-tat-tung-bai/`.\n\n')

lines.append('### Thiếu PDF gốc (5 bài)\n')
lines.append('- QT037 Praptomojati (CA-CBT SEA): bản dịch dựa trên abstract SR\n')
lines.append('- QT048–QT051 (Chen COVID Meta, Zhang Bayesian, Qiaochu Mobile CBT, Menon LMIC): paywall — bản dịch abstract-only\n\n')

lines.append('---\n\n')
lines.append('## QUY TẮC ĐẶT TÊN FILE (canonical v1 — 11/04/2026)\n\n')
lines.append('1. **Prefix:** `VN{001-999}` cho bài Việt Nam, `QT{001-999}` cho bài quốc tế\n')
lines.append('2. **Padding:** 3 chữ số (để mở rộng lên 999+ bài trong tương lai)\n')
lines.append('3. **Descriptor:** `TacGia_ChuDe/Diadiem_NamXB` (dùng `_`, không dấu cách)\n')
lines.append('4. **Tóm tắt:** `Tom-tat-tung-bai/{ID}_{Descriptor}.docx`\n')
lines.append('5. **Bản dịch:** `03_Ban-dich/{ID}_{Descriptor}.docx`\n')
lines.append('6. **PDF gốc:** `02_Papers-goc/{Folder}/{ID}_{Descriptor}.pdf`\n\n')
lines.append('- Folders: `Viet-Nam/`, `The-gioi_Au-My-Uc/`, `The-gioi_Khac/`, `Dong-Nam-A/`, `11-bai-ban-dau-va-mo-rong/`\n')

out = '02_Papers-goc/MASTER_INDEX.md'
with open(out, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print(f'Saved: {out}')
print(f'Total: {len(idx)} papers')
