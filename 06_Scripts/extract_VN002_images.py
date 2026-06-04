# -*- coding: utf-8 -*-
"""Extract all images from VN002 VNAMHS PDF."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from pypdf import PdfReader
from PIL import Image

PDF = '02_Papers-goc/Viet-Nam/VN002_VNAMHS_2022_National.pdf'
OUTDIR = 'C:/Users/HLC/AppData/Local/Temp/vnamhs_imgs/'
os.makedirs(OUTDIR, exist_ok=True)

reader = PdfReader(PDF)
count = 0
for pi, page in enumerate(reader.pages):
    try:
        images = page.images
    except Exception:
        continue
    for ii, img_obj in enumerate(images):
        try:
            data = img_obj.data
            name = img_obj.name
            out = os.path.join(OUTDIR, f'p{pi+1:03d}_img{ii+1}_{name}')
            with open(out, 'wb') as f:
                f.write(data)
            # Convert JP2 to PNG if needed
            if out.lower().endswith(('.jp2', '.jpx', '.j2k')):
                try:
                    im = Image.open(out)
                    if im.mode not in ('RGB', 'L'):
                        im = im.convert('RGB')
                    png_out = os.path.splitext(out)[0] + '.png'
                    im.save(png_out, 'PNG')
                    print(f'  converted {os.path.basename(out)} -> {os.path.basename(png_out)}')
                except Exception as e:
                    print(f'  JP2 conversion failed: {e}')
            count += 1
            print(f'  p{pi+1:03d} img{ii+1}: {os.path.basename(out)} ({len(data):,} bytes)')
        except Exception as e:
            print(f'  ERROR p{pi+1} img{ii}: {e}')

print(f'\nTotal extracted: {count}')
print(f'Output dir: {OUTDIR}')
