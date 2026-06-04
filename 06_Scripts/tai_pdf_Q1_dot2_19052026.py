# -*- coding: utf-8 -*-
"""Tai PDF toan van 8 bai ung vien tham khao bai Q1 (dot 2). 19/05/2026.
Thu nhieu URL/nguon cho moi bai; bao cao bai nao tai duoc, bai nao can tai tay."""
import os, sys, io, time
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import urllib.request

OUT = r'c:\Users\OS\OneDrive\read_books\Lo-au\bai-bao-khgdvn\bài báo tham khảo\Q1_dot2_19052026'
os.makedirs(OUT, exist_ok=True)

HEADERS = {
    'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                   '(KHTML, like Gecko) Chrome/124.0 Safari/537.36'),
    'Accept': 'application/pdf,text/html,*/*',
}

# Moi bai: ten file + danh sach URL PDF ung vien (thu lan luot)
PAPERS = [
    ('V1_NguyenTT_2024_SocialSupport_HPR', [
        'https://pmc.ncbi.nlm.nih.gov/articles/PMC11093747/pdf/',
        'https://www.ncbi.nlm.nih.gov/pmc/articles/PMC11093747/pdf/',
    ]),
    ('V2_TranLCT_2025_CanTho_GMH', [
        'https://pmc.ncbi.nlm.nih.gov/articles/PMC12720387/pdf/',
        'https://www.cambridge.org/core/services/aop-cambridge-core/content/view/'
        '10.1017/gmh.2025.10096',
    ]),
    ('V3_Ho_2022_AcademicStress_CurrPsychol', [
        'https://pmc.ncbi.nlm.nih.gov/articles/PMC9574843/pdf/',
        'https://link.springer.com/content/pdf/10.1007/s12144-022-03661-3.pdf',
    ]),
    ('Q1_Chen_2025_SchoolClimate_BMCPsychol', [
        'https://bmcpsychology.biomedcentral.com/counter/pdf/10.1186/s40359-025-02364-1.pdf',
        'https://pmc.ncbi.nlm.nih.gov/articles/PMC11748506/pdf/',
    ]),
    ('Q2_Wang_2025_AcademicBurden_JAdolesc', [
        'https://pmc.ncbi.nlm.nih.gov/articles/PMC12128909/pdf/',
        'https://onlinelibrary.wiley.com/doi/pdfdirect/10.1002/jad.12471',
    ]),
    ('Q3_Sutic_2025_ParentingAnxiety_JCPP', [
        'https://pmc.ncbi.nlm.nih.gov/articles/PMC12350824/pdf/',
        'https://onlinelibrary.wiley.com/doi/pdfdirect/10.1111/jcpp.14161',
        'https://acamh.onlinelibrary.wiley.com/doi/pdfdirect/10.1111/jcpp.14161',
    ]),
    ('Q4_Yang_2024_HarshParenting_CAPMH', [
        'https://capmh.biomedcentral.com/counter/pdf/10.1186/s13034-024-00826-9.pdf',
        'https://pmc.ncbi.nlm.nih.gov/articles/PMC11515719/pdf/',
    ]),
    ('Q5_Tan_2025_AnxietyTransmission_PsychiatrQ', [
        'https://link.springer.com/content/pdf/10.1007/s11126-025-10166-2.pdf',
    ]),
]

def try_download(url):
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=45) as r:
        data = r.read()
    if data[:4] == b'%PDF':
        return data
    # doi khi PDF nam sau redirect/HTML -> khong nhan
    return None

results = []
for name, urls in PAPERS:
    saved = None; err = ''
    for u in urls:
        try:
            data = try_download(u)
            if data:
                path = os.path.join(OUT, name + '.pdf')
                with open(path, 'wb') as f:
                    f.write(data)
                saved = (path, len(data), u)
                break
            else:
                err = 'khong phai PDF (co the bi chan/redirect HTML)'
        except Exception as e:
            err = '%s: %s' % (type(e).__name__, str(e)[:80])
        time.sleep(1)
    if saved:
        print('  [OK]   %-44s %6.0f KB' % (name, saved[1]/1024))
        results.append((name, 'OK', saved[2]))
    else:
        print('  [FAIL] %-44s %s' % (name, err))
        results.append((name, 'FAIL', err))

ok = sum(1 for _, s, _ in results if s == 'OK')
print()
print('Tai duoc %d/%d PDF. Thu muc: %s' % (ok, len(PAPERS), OUT))
