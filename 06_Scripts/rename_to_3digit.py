# -*- coding: utf-8 -*-
"""
Rename: VN01 -> VN001, QT01 -> QT001 (3-digit padding for future scalability)
"""
import os, sys, io, shutil, re, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
os.chdir(ROOT)

# Find all files with VNNN_ or QTNN_ prefix and rename to 3-digit
PATTERN = re.compile(r'^(VN|QT)(\d{2})_')

dirs_to_scan = [
    'Tom-tat-tung-bai',
    '03_Ban-dich',
    '02_Papers-goc/Viet-Nam',
    '02_Papers-goc/The-gioi_Au-My-Uc',
    '02_Papers-goc/The-gioi_Khac',
    '02_Papers-goc/Dong-Nam-A',
    '02_Papers-goc/11-bai-ban-dau-va-mo-rong',
]

renamed = 0
skipped = 0
already_3 = 0

for d in dirs_to_scan:
    if not os.path.exists(d):
        continue
    print(f'\n=== {d} ===')
    for f in os.listdir(d):
        m = PATTERN.match(f)
        if not m:
            continue
        prefix = m.group(1)  # VN or QT
        num_str = m.group(2)  # 2-digit
        num = int(num_str)
        new_prefix = f'{prefix}{num:03d}_'
        old_prefix = f'{prefix}{num_str}_'
        new_name = f.replace(old_prefix, new_prefix, 1)
        if new_name == f:
            skipped += 1
            continue
        src = os.path.join(d, f)
        dst = os.path.join(d, new_name)
        if os.path.exists(dst):
            print(f'  ALREADY: {new_name}')
            continue
        shutil.move(src, dst)
        renamed += 1
        # Print compact
        print(f'  {f[:40]:<40} -> {new_name[:50]}')

print(f'\n=== TOTAL: {renamed} renamed ===')
