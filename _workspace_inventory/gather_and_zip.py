# -*- coding: utf-8 -*-
"""Đọc MISSING_FILES_FOR_OTHER_MACHINE.txt → copy 185 file vào folder riêng → tạo zip."""
import sys, io, os, re, shutil, zipfile
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = r'c:\Users\OS\OneDrive\read_books\Lo-au'
LIST_FILE = os.path.join(ROOT, 'MISSING_FILES_FOR_OTHER_MACHINE.txt')
STAGE_DIR = os.path.join(ROOT, '_to_send_to_other_machine')
ZIP_PATH = os.path.join(ROOT, f'Lo-au_sync_to_other_machine_14052026.zip')


def fmt_size(sz):
    sz = float(sz)
    for unit in ['B', 'KB', 'MB', 'GB']:
        if sz < 1024:
            return f'{sz:.1f}{unit}'
        sz /= 1024
    return f'{sz:.1f}TB'


# === Parse MISSING_FILES file ===
# Mỗi dòng cần lấy có dạng:
#     path/to/file.pdf    (522.8KB, mod 2026-03-28)
# Loại bỏ section PHỤ LỤC (file thừa trên máy này — không cần gom)

paths = []
with open(LIST_FILE, 'r', encoding='utf-8') as f:
    in_appendix = False
    for line in f:
        line = line.rstrip('\n')
        if 'PHỤ LỤC' in line:
            in_appendix = True
            continue
        if in_appendix:
            continue
        # Match line dạng: spaces + path + spaces + (size + mod date)
        # Size pattern: (522.8KB, mod 2026-03-28) — hoặc (0.0B, mod ...) hoặc (400.0B, ...)
        m = re.match(r'^\s+(.+?)\s+\(\d+\.?\d*\s*(?:[KMGT]i?)?B,\s+mod\s+\d{4}-\d{2}-\d{2}\)\s*$', line)
        if not m:
            continue
        path = m.group(1).strip()
        # Skip nếu không phải đường dẫn file (vd "Tổng file thiếu trên máy này:  185")
        if not ('/' in path or '.' in path):
            continue
        # Skip lines like "[02_Papers-goc/]" (heading dạng [path])
        if path.startswith('[') and path.endswith(']'):
            continue
        if path.startswith('('):
            continue
        paths.append(path)

print(f'Parsed {len(paths)} file paths from MISSING_FILES list')

# === Verify + copy ===
os.makedirs(STAGE_DIR, exist_ok=True)
copied = []
missing = []
total_size = 0

for p in paths:
    src = os.path.join(ROOT, p.replace('/', os.sep))
    if not os.path.exists(src):
        missing.append(p)
        continue
    dst = os.path.join(STAGE_DIR, p.replace('/', os.sep))
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    try:
        shutil.copy2(src, dst)
        sz = os.path.getsize(src)
        copied.append((p, sz))
        total_size += sz
    except Exception as e:
        missing.append(f'{p} (copy error: {e})')

print(f'\nCopied: {len(copied)} files')
print(f'Missing on this machine: {len(missing)}')
if missing:
    print('Missing files:')
    for m in missing[:20]:
        print(f'  - {m}')
    if len(missing) > 20:
        print(f'  ... và {len(missing)-20} file khác')
print(f'Total size: {fmt_size(total_size)}')

# === Create zip ===
print(f'\nCreating zip: {ZIP_PATH}')
with zipfile.ZipFile(ZIP_PATH, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
    for p, sz in copied:
        src = os.path.join(STAGE_DIR, p.replace('/', os.sep))
        # Trong zip, dùng đường dẫn tương đối với ROOT
        zf.write(src, arcname=p)

zip_sz = os.path.getsize(ZIP_PATH)
print(f'Zip size: {fmt_size(zip_sz)} (compression ratio: {zip_sz/max(total_size,1)*100:.1f}%)')

# === Summary report ===
SUMMARY = os.path.join(ROOT, 'SYNC_SUMMARY_14052026.txt')
with open(SUMMARY, 'w', encoding='utf-8') as f:
    f.write('SYNC SUMMARY — Files to send to OTHER MACHINE\n')
    f.write('='*70 + '\n\n')
    f.write(f'Date: 2026-05-14\n')
    f.write(f'Source machine: c:\\Users\\OS\\OneDrive\\read_books\\Lo-au\n')
    f.write(f'Target machine: c:\\Users\\HLC\\OneDrive\\read_books\\Lo-au\n\n')
    f.write(f'Files copied to staging: {len(copied)}\n')
    f.write(f'Files missing on this machine: {len(missing)}\n')
    f.write(f'Total uncompressed size: {fmt_size(total_size)}\n')
    f.write(f'Zip size: {fmt_size(zip_sz)}\n\n')
    f.write(f'Staging folder: {STAGE_DIR}\n')
    f.write(f'Zip file: {ZIP_PATH}\n\n')

    # Group by top-level dir
    by_top = {}
    for p, sz in copied:
        top = p.split('/')[0]
        if top not in by_top:
            by_top[top] = [0, 0]
        by_top[top][0] += 1
        by_top[top][1] += sz
    f.write('BREAKDOWN by top-level directory:\n')
    f.write('-'*70 + '\n')
    for top, (cnt, sz) in sorted(by_top.items(), key=lambda x: -x[1][1]):
        f.write(f'  {top:50s}  {cnt:5d} file  {fmt_size(sz):>10s}\n')

    if missing:
        f.write('\n\nFILES MISSING ON THIS MACHINE (cannot copy):\n')
        f.write('-'*70 + '\n')
        for m in missing:
            f.write(f'  - {m}\n')

print(f'\nWrote summary: {SUMMARY}')
