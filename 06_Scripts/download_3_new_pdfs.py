"""
Tải 3 PDF cho VN007 + QT059 + QT060 từ open access.
Update canonical_index.json với pdf path sau khi tải.
"""
import sys, io, json
from pathlib import Path
import urllib.request
import ssl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = Path('c:/Users/HLC/OneDrive/read_books/Lo-au')
IDX_PATH = BASE / '02_Papers-goc' / 'canonical_index.json'

DOWNLOADS = [
    {
        'id': 'VN007',
        'url': 'https://www.italjmed.org/ijm/article/view/1888/1839',
        'dst': BASE / '02_Papers-goc' / 'Viet-Nam' / 'VN007_NguyenTL_Gender_PostPandemic_ItalJMed_2025.pdf',
        'folder': 'Viet-Nam',
    },
    {
        'id': 'QT059',
        'url': 'https://pmc.ncbi.nlm.nih.gov/articles/PMC12127306/pdf/fpsyt-16-1594658.pdf',
        'dst': BASE / '02_Papers-goc' / 'The-gioi_Khac' / 'QT059_Cai_SchoolResilience_SR_MA_FrontPsych_2025.pdf',
        'folder': 'The-gioi_Khac',
    },
    {
        'id': 'QT060',
        'url': 'https://www.frontiersin.org/journals/psychiatry/articles/10.3389/fpsyt.2024.1489427/pdf',
        'dst': BASE / '02_Papers-goc' / 'The-gioi_Au-My-Uc' / 'QT060_Bie_GlobalAnxiety_GBD_10-24y_1990_2021_FrontPsych_2024.pdf',
        'folder': 'The-gioi_Au-My-Uc',
    },
]

# Bypass SSL verify for some servers
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

with open(IDX_PATH, encoding='utf-8') as f:
    idx = json.load(f)

log = []
for d in DOWNLOADS:
    if d['dst'].exists():
        log.append(f"{d['id']}: ALREADY EXISTS at {d['dst'].name} ({d['dst'].stat().st_size//1024} KB)")
        continue
    try:
        req = urllib.request.Request(d['url'], headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, context=ctx, timeout=60) as resp:
            data = resp.read()
        # Verify it's a PDF
        if not data[:4] == b'%PDF':
            log.append(f"{d['id']}: NOT A PDF (got {data[:50]})")
            continue
        d['dst'].write_bytes(data)
        size_kb = len(data) // 1024
        log.append(f"{d['id']}: DOWNLOADED → {d['dst'].name} ({size_kb} KB)")
        # Update index
        idx[d['id']]['pdf'] = d['dst'].name
        idx[d['id']]['pdf_folder'] = d['folder']
        idx[d['id']]['scope'] = 'main'
    except Exception as e:
        log.append(f"{d['id']}: FAILED — {e}")

with open(IDX_PATH, 'w', encoding='utf-8') as f:
    json.dump(idx, f, ensure_ascii=False, indent=2)

for l in log:
    print(l)
