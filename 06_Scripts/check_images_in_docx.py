# -*- coding: utf-8 -*-
"""Check how many images are embedded in the docx."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
import zipfile

PATH = '03_Ban-dich/VN022_UNICEF_VN_2022_SchoolFactors_FULL.docx'

# Check via zipfile (docx is a zip)
with zipfile.ZipFile(PATH) as z:
    media = [n for n in z.namelist() if n.startswith('word/media/')]
    print(f'Images in docx: {len(media)}')
    for m in sorted(media):
        info = z.getinfo(m)
        print(f'  {m} ({info.file_size:,} bytes)')
