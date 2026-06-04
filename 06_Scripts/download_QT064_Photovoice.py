"""Download Photovoice scoping review 2023 — open access UCL Discovery."""
import sys, io, ssl, urllib.request
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

URL = 'https://discovery.ucl.ac.uk/id/eprint/10179741/1/A%20systematic%20scoping%20review%20of%20Photovoice%20within%20mental%20health%20research%20involving%20adolescents.pdf'
DST = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/02_Papers-goc/The-gioi_Au-My-Uc/QT064_Stephens_Photovoice_ScopingReview_IntlJAdolYouth_2023.pdf')

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

if DST.exists():
    print(f'EXISTS: {DST.name} ({DST.stat().st_size//1024} KB)')
else:
    req = urllib.request.Request(URL, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        with urllib.request.urlopen(req, context=ctx, timeout=90) as resp:
            data = resp.read()
        if data[:4] == b'%PDF':
            DST.write_bytes(data)
            print(f'DOWNLOADED: {DST.name} ({len(data)//1024} KB)')
        else:
            print(f'NOT A PDF: {data[:50]}')
    except Exception as e:
        print(f'FAILED: {e}')
