# -*- coding: utf-8 -*-
"""Strip metadata + watermark TAT CA file docx trong bai-bao-Q1.
Aggressive: core/app/custom, header/footer, comments, watermark, rsids."""
import os, sys, io, zipfile, shutil, re, glob
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
TARGETS = sorted(glob.glob(os.path.join(ROOT, 'bai-bao-Q1', '*.docx')))


def strip_one(path):
    if '~$' in path: return False
    tmp = path + '.tmp'
    try:
        with zipfile.ZipFile(path, 'r') as zin:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
                for name in zin.namelist():
                    data = zin.read(name)
                    if name == 'docProps/core.xml':
                        data = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                                b'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                                b'xmlns:dc="http://purl.org/dc/elements/1.1/" '
                                b'xmlns:dcterms="http://purl.org/dc/terms/" '
                                b'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
                                b'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                                b'<dc:creator></dc:creator><dc:title></dc:title>'
                                b'<dc:subject></dc:subject><dc:description></dc:description>'
                                b'<cp:keywords></cp:keywords><cp:lastModifiedBy></cp:lastModifiedBy>'
                                b'<cp:category></cp:category><cp:contentStatus></cp:contentStatus>'
                                b'<cp:revision>1</cp:revision></cp:coreProperties>')
                    elif name == 'docProps/app.xml':
                        data = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                                b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
                                b'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                                b'<Application></Application><Company></Company><Manager></Manager>'
                                b'<DocSecurity>0</DocSecurity><HyperlinksChanged>false</HyperlinksChanged>'
                                b'<AppVersion>1.0</AppVersion></Properties>')
                    elif name == 'docProps/custom.xml':
                        data = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                                b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
                                b'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>')
                    elif name.startswith('word/comments') or name.startswith('word/people'):
                        continue
                    elif name == 'word/document.xml' or name.startswith('word/header') or name.startswith('word/footer'):
                        text = data.decode('utf-8', errors='replace')
                        text = re.sub(r'<v:shape[^>]*_x0000_t136[^>]*>.*?</v:shape>', '', text, flags=re.DOTALL)
                        text = re.sub(r'<w:pict>.*?(watermark|Watermark|WATERMARK).*?</w:pict>', '', text, flags=re.DOTALL)
                        text = re.sub(r'\sw:rsid[A-Za-z]*="[^"]*"', '', text)
                        data = text.encode('utf-8')
                    elif name == 'word/settings.xml':
                        text = data.decode('utf-8', errors='replace')
                        text = re.sub(r'<w:rsids>.*?</w:rsids>', '', text, flags=re.DOTALL)
                        text = re.sub(r'<w:trackChanges/?>', '', text)
                        data = text.encode('utf-8')
                    zout.writestr(name, data)
        shutil.move(tmp, path)
        return True
    except PermissionError:
        if os.path.exists(tmp): os.remove(tmp)
        return None
    except Exception as e:
        if os.path.exists(tmp): os.remove(tmp)
        print(f'  ERROR {os.path.basename(path)}: {e}')
        return False


print(f'Stripping {len(TARGETS)} files in bai-bao-Q1/...')
ok = locked = failed = 0
for path in TARGETS:
    name = os.path.basename(path)
    result = strip_one(path)
    if result is True: ok += 1
    elif result is None:
        locked += 1
        print(f'  LOCKED (in use): {name}')
    elif result is False: failed += 1
print(f'\nDone. OK: {ok}, Locked: {locked}, Failed: {failed}')
