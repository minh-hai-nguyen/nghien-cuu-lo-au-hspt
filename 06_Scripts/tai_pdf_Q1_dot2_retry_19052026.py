# -*- coding: utf-8 -*-
"""Tai lai PDF cac bai con thieu qua Europe PMC (than thien voi script hon NCBI)."""
import os, sys, io, time
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import urllib.request

OUT = r'c:\Users\OS\OneDrive\read_books\Lo-au\bai-bao-khgdvn\bài báo tham khảo\Q1_dot2_19052026'
os.makedirs(OUT, exist_ok=True)
HEADERS = {'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                          '(KHTML, like Gecko) Chrome/124.0 Safari/537.36'),
           'Accept': 'application/pdf,*/*'}

# Bai con thieu -> ma PMC (de tai qua Europe PMC render)
REMAIN = [
    ('V1_NguyenTT_2024_SocialSupport_HPR', 'PMC11093747'),
    ('V2_TranLCT_2025_CanTho_GMH',        'PMC12720387'),
    ('V3_Ho_2022_AcademicStress_CurrPsychol', 'PMC9574843'),
    ('Q2_Wang_2025_AcademicBurden_JAdolesc',  'PMC12128909'),
    ('Q3_Sutic_2025_ParentingAnxiety_JCPP',   'PMC12350824'),
]

def fetch(url):
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=60) as r:
        data = r.read()
    return data if data[:4] == b'%PDF' else None

results = []
for name, pmc in REMAIN:
    if os.path.exists(os.path.join(OUT, name + '.pdf')):
        print('  [DA CO]', name); results.append((name, 'OK')); continue
    saved = None; err = ''
    for url in ['https://europepmc.org/articles/%s?pdf=render' % pmc,
                'https://www.ebi.ac.uk/europepmc/webservices/rest/%s/fullTextPDF' % pmc]:
        try:
            data = fetch(url)
            if data:
                with open(os.path.join(OUT, name + '.pdf'), 'wb') as f:
                    f.write(data)
                saved = (len(data), url); break
            err = 'khong phai PDF'
        except Exception as e:
            err = '%s: %s' % (type(e).__name__, str(e)[:70])
        time.sleep(3)
    if saved:
        print('  [OK]   %-42s %6.0f KB' % (name, saved[0]/1024))
        results.append((name, 'OK'))
    else:
        print('  [FAIL] %-42s %s' % (name, err))
        results.append((name, 'FAIL'))
    time.sleep(2)

print()
print('=== Tong ket thu muc ===')
for f in sorted(os.listdir(OUT)):
    p = os.path.join(OUT, f)
    print('  %-46s %7.0f KB' % (f, os.path.getsize(p)/1024))
