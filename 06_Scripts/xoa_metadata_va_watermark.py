# -*- coding: utf-8 -*-
"""Xoa SACH metadata + watermark khoi cac file v2, v3, v4.
- Core properties (author, title, subject, comments, keywords...)
- App properties (Application, Company, Manager...)
- Custom properties
- Watermark (header/footer text)
- Revisions, tracked changes
- Comments
01/06/2026."""
import os, sys, io, zipfile, shutil, re
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

FILES = [
    'bai-bao-Q1/ThamKhao_Titles_Q1Q3_AsiaChauPhi_TiengViet_v2_01062026.docx',
    'bai-bao-Q1/ThamKhao_Titles_Q1Q3_AsiaChauPhi_TiengViet_v3_01062026.docx',
    'bai-bao-Q1/ThamKhao_Titles_Q1Q3_AsiaChauPhi_TiengViet_v4_01062026.docx',
    'bai-bao-Q1/ThamKhao_MoRong_QuocTe_TiengViet_01062026.docx',
    'bai-bao-Q1/ThamKhao_MoRong_QuocTe_TiengViet_v2_01062026.docx',
    'bai-bao-Q1/ThamKhao_MoRong_QuocTe_TiengViet_v3_01062026.docx',
    'bai-bao-Q1/Draft_Q1_SongNgu_v1_01062026.docx',
    'bai-bao-Q1/Draft_Q1_SongNgu_v2_01062026.docx',
    'bai-bao-Q1/Draft_Q1_SongNgu_v3_01062026.docx',
    'bai-bao-Q1/4VanDe_BLOCKING_Q1Q3_01062026.docx',
    'bai-bao-Q1/Verify_BaoVeLA_DSM_ICD_Karasu_01062026.docx',
    'bai-bao-Q1/Verify_BaoVeLA_DSM_ICD_Karasu_v2_01062026.docx',
    'bai-bao-Q1/Verify_BaoVeLA_DSM_ICD_Karasu_v3_01062026.docx',
    'bai-bao-Q1/AuditBaoCao_ThamKhao_Titles_01062026.docx',
]


def strip_one(path):
    """Strip metadata/watermark by manipulating the docx (zip) directly."""
    if not os.path.exists(path):
        print(f'  SKIP (not exist): {path}')
        return False

    tmp = path + '.tmp'

    with zipfile.ZipFile(path, 'r') as zin:
        names = zin.namelist()
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                data = zin.read(name)

                if name == 'docProps/core.xml':
                    # Replace with minimal core
                    data = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'\
                           b'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '\
                           b'xmlns:dc="http://purl.org/dc/elements/1.1/" '\
                           b'xmlns:dcterms="http://purl.org/dc/terms/" '\
                           b'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '\
                           b'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'\
                           b'<dc:creator></dc:creator>'\
                           b'<dc:title></dc:title>'\
                           b'<dc:subject></dc:subject>'\
                           b'<dc:description></dc:description>'\
                           b'<cp:keywords></cp:keywords>'\
                           b'<cp:lastModifiedBy></cp:lastModifiedBy>'\
                           b'<cp:category></cp:category>'\
                           b'<cp:contentStatus></cp:contentStatus>'\
                           b'<cp:revision>1</cp:revision>'\
                           b'</cp:coreProperties>'

                elif name == 'docProps/app.xml':
                    # Strip application/company/manager
                    data = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'\
                           b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '\
                           b'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'\
                           b'<Application></Application>'\
                           b'<Company></Company>'\
                           b'<Manager></Manager>'\
                           b'<DocSecurity>0</DocSecurity>'\
                           b'<HyperlinksChanged>false</HyperlinksChanged>'\
                           b'<AppVersion>1.0</AppVersion>'\
                           b'</Properties>'

                elif name == 'docProps/custom.xml':
                    # Empty custom properties
                    data = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'\
                           b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '\
                           b'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'\
                           b'</Properties>'

                elif name.startswith('word/comments') or name.startswith('word/people'):
                    # Skip comments entirely
                    continue

                elif name.startswith('word/header') or name.startswith('word/footer'):
                    # Strip any watermark in headers/footers
                    text = data.decode('utf-8', errors='replace')
                    # Remove watermark VML (v:shape with type "_x0000_t136" used for WordArt watermark)
                    text = re.sub(r'<v:shape[^>]*_x0000_t136[^>]*>.*?</v:shape>',
                                  '', text, flags=re.DOTALL)
                    # Remove any <w:pict> containing watermark-related elements
                    text = re.sub(r'<w:pict>.*?(watermark|Watermark|WATERMARK).*?</w:pict>',
                                  '', text, flags=re.DOTALL)
                    data = text.encode('utf-8')

                elif name == 'word/document.xml':
                    # Also strip any watermark in body just in case + strip rsid
                    text = data.decode('utf-8', errors='replace')
                    text = re.sub(r'<v:shape[^>]*_x0000_t136[^>]*>.*?</v:shape>',
                                  '', text, flags=re.DOTALL)
                    # Strip rsid attributes (revision save IDs)
                    text = re.sub(r'\sw:rsid[A-Za-z]*="[^"]*"', '', text)
                    data = text.encode('utf-8')

                elif name == 'word/settings.xml':
                    # Strip rsid list + revisions
                    text = data.decode('utf-8', errors='replace')
                    text = re.sub(r'<w:rsids>.*?</w:rsids>', '', text, flags=re.DOTALL)
                    text = re.sub(r'<w:trackChanges/?>', '', text)
                    data = text.encode('utf-8')

                zout.writestr(name, data)

    shutil.move(tmp, path)
    return True


for relpath in FILES:
    fullpath = os.path.join(ROOT, relpath)
    print(f'Processing: {os.path.basename(fullpath)}')
    if strip_one(fullpath):
        size = os.path.getsize(fullpath)
        print(f'  OK ({size/1024:.1f} KB)')

print('\nDONE. Verify by opening in Word > File > Info > Inspect Document.')
