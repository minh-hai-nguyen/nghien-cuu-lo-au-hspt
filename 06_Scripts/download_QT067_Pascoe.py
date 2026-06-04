"""Download Pascoe 2020 PDF (open access Victoria U)."""
import sys, io, ssl, urllib.request
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

URLS = [
    'https://vuir.vu.edu.au/39399/1/02673843.2019..pdf',
    'https://www.tandfonline.com/doi/pdf/10.1080/02673843.2019.1596823',
]
DST = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/02_Papers-goc/The-gioi_Au-My-Uc/QT067_Pascoe_AcademicStress_IJAY_2020.pdf')

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

if DST.exists():
    print(f'EXISTS: {DST.name} ({DST.stat().st_size//1024} KB)')
else:
    for url in URLS:
        print(f'Trying: {url}')
        try:
            req = urllib.request.Request(url, headers={'User-Agent':'Mozilla/5.0'})
            with urllib.request.urlopen(req, context=ctx, timeout=60) as resp:
                data = resp.read()
            if data[:4] == b'%PDF':
                DST.write_bytes(data)
                print(f'  DOWNLOADED → {DST.name} ({len(data)//1024} KB)')
                break
            else:
                print(f'  NOT A PDF (got {data[:50]!r})')
        except Exception as e:
            print(f'  FAILED: {e}')
