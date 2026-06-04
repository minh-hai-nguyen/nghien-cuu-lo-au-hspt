# -*- coding: utf-8 -*-
"""Generate PROJECT_INVENTORY.md for Lo-au project."""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from datetime import datetime

ROOT = r'c:\Users\OS\OneDrive\read_books\Lo-au'
OUT = os.path.join(ROOT, 'PROJECT_INVENTORY_13052026.md')

exclude_dirs = {'.git', 'node_modules', '__pycache__', '.venv', 'venv', '.claude', '_workspace', '_workspace_inventory'}
exclude_exts = {'.pyc', '.pyo'}


def fmt_size(sz):
    sz = float(sz)
    for unit in ['B', 'KB', 'MB', 'GB']:
        if sz < 1024:
            return f'{sz:.1f}{unit}'
        sz /= 1024
    return f'{sz:.1f}TB'


# Collect all files
all_files = []
for root, dirs, files in os.walk(ROOT):
    dirs[:] = [d for d in dirs if d not in exclude_dirs]
    for f in files:
        if any(f.endswith(e) for e in exclude_exts):
            continue
        full = os.path.join(root, f)
        rel = os.path.relpath(full, ROOT).replace(os.sep, '/')
        try:
            st = os.stat(full)
            all_files.append((rel, st.st_size, st.st_mtime))
        except Exception:
            pass

all_files.sort()
total_size = sum(f[1] for f in all_files)
print(f'Total files: {len(all_files)}')
print(f'Total size: {fmt_size(total_size)}')

# Group by top-level dir
by_top = {}
for rel, sz, mt in all_files:
    parts = rel.split('/', 1)
    top = parts[0] if len(parts) > 1 else '(root)'
    if top not in by_top:
        by_top[top] = []
    by_top[top].append((rel, sz, mt))

top_order = sorted(by_top.keys(), key=lambda k: -sum(x[1] for x in by_top[k]))

descriptions = {
    '(root)': 'File cấp gốc — config, README, etc.',
    '00_Meta': 'Metadata + ghi chú dự án',
    '01_Bao-cao': 'Báo cáo + trả lời thầy',
    '02_Papers-goc': 'PDF gốc tài liệu (Lancet, JAMA, Springer, WHO, GBD…)',
    '03_Ban-dich': 'Bản dịch + phản biện docx (Brown, Herres, Steinhoff, GBD, WHO)',
    '04_Co-so-du-lieu': 'Cơ sở dữ liệu Excel/CSV (V-NAMHS, GBD)',
    '05_Cong-cu': 'Công cụ thang đo (DASS-21, SCAS, PHQ-9, GAD-7, DSM-5 Emerging Measures)',
    '06_Scripts': 'Script Python sinh doc + RAG + KG + workspace',
    'Bai_goc_BaoCao_CanThiep_10042026': 'Bài gốc báo cáo can thiệp',
    'Cropped_Figures': 'Hình ảnh đã cắt từ PDF',
    'DocFiles': 'File docx tổng hợp',
    'Materials': 'Vật liệu thô',
    'Tom-tat-tung-bai': 'Tóm tắt từng bài đã dịch',
    'Tập-viết-theo-phong-cách': 'Bài tập viết theo phong cách',
    '_Archive': 'File lưu trữ cũ',
    '_Archive_Scripts': 'Script cũ',
    'chưa-sắp-xếp': 'Tài liệu chưa phân loại',
    'lib': 'Thư viện code',
    'madam Thao': 'Tài liệu của madam Thao',
    'paper-may': 'Bài cần sửa cấp tốc (Cong Thi Hang Turnitin 33%)',
    'rag_bao_cao_can_thiep': 'ChromaDB RAG cho báo cáo can thiệp',
    'rag_db': 'ChromaDB RAG chung',
    'rag_dich_phan_bien': 'ChromaDB RAG cho bản dịch + phản biện',
    'tro-ly-nghien-cuu-tam-ly': 'Trợ lý nghiên cứu (full)',
    'tro-ly-nghien-cuu-tam-ly-light': 'Trợ lý nghiên cứu (light)',
    'web': 'Web frontend',
}

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('# PROJECT INVENTORY — Thư mục Lo-au\n\n')
    f.write(f'**Ngày xuất**: {datetime.now().strftime("%Y-%m-%d %H:%M")}\n\n')
    f.write(f'**Máy**: Windows (c:/Users/OS/OneDrive/read_books/Lo-au)\n\n')
    f.write(f'**Tổng số file**: {len(all_files):,}\n\n')
    f.write(f'**Tổng dung lượng**: {fmt_size(total_size)}\n\n')
    f.write(f'**Loại trừ**: .git, node_modules, __pycache__, .venv, venv, .claude, _workspace, *.pyc\n\n')
    f.write('## Mục đích\n\nFile này được xuất để mang sang **máy thứ hai** đối chiếu — máy đó sẽ '
            'so với inventory của chính nó và đề xuất các file/thư mục nên sync từ máy này sang.\n\n')
    f.write('---\n\n')

    f.write('## Tổng quan các thư mục cấp cao\n\n')
    f.write('| Thư mục | Số file | Dung lượng | Mô tả |\n')
    f.write('|---|---|---|---|\n')
    for top in top_order:
        files = by_top[top]
        sz = sum(x[1] for x in files)
        f.write(f'| `{top}` | {len(files)} | {fmt_size(sz)} | {descriptions.get(top, "")} |\n')
    f.write('\n---\n\n')

    f.write('## Chi tiết từng thư mục\n\n')
    for top in top_order:
        files = by_top[top]
        sz = sum(x[1] for x in files)
        f.write(f'### `{top}/` — {len(files)} file, {fmt_size(sz)}\n\n')
        if descriptions.get(top):
            f.write(f'*{descriptions[top]}*\n\n')
        # If many files, only list large/important ones + summary
        if len(files) > 50:
            f.write(f'<details><summary>Click để xem đầy đủ {len(files)} file</summary>\n\n')
        for rel, fsz, mt in files:
            dt = datetime.fromtimestamp(mt).strftime('%Y-%m-%d')
            f.write(f'- `{rel}` — {fmt_size(fsz)}, mod {dt}\n')
        if len(files) > 50:
            f.write('\n</details>\n')
        f.write('\n')

print(f'Wrote: {OUT}')
print(f'Size: {os.path.getsize(OUT) // 1024} KB')
