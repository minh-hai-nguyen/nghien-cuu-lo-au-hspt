# -*- coding: utf-8 -*-
"""Clean ALL watermarks + metadata tu tat ca file docx em da tao trong phien.
01/06/2026."""
import os, sys, io, zipfile, shutil
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

# Files to clean - tat ca file em da tao
FILES = [
    'Luận án TS/BanDich_BriefCOPE_AdaptiveMaladaptive_29052026.docx',
    'Luận án TS/TrichYeuLA_CongThiHang_v2_29052026.docx',
    'Luận án TS/DanhMucCongTrinh_EN_v1_29052026.docx',
    'Luận án TS/TomTatLA_v2_VERIFIED_29052026.docx',
    'Luận án TS/TomTatLA_EN_v1_29052026.docx',
    'Luận án TS/TomTatLA_CHECKLIST_NOP_29052026.docx',
    'bai-bao-Q1/BoCuc_Q1_Q3_PhuongAnA_v2_01062026.docx',
    'bai-bao-Q1/BoCuc_Q1_Q3_PhuongAnA_01062026.docx',
    'bai-bao-Q1/RaSoat_Q1_01062026.docx',
    'bai-bao-Q1/BaiA_JAD_OUTLINE_v1_30052026.docx',
    'bai-bao-Q1/BaiD_StressHealth_OUTLINE_v1_30052026.docx',
]


def clean_via_docx(path):
    """Clean metadata via python-docx."""
    try:
        d = Document(path)
        cp = d.core_properties
        # Clear ALL core properties
        cp.author = ''
        cp.title = ''
        cp.subject = ''
        cp.keywords = ''
        cp.comments = ''
        cp.last_modified_by = ''
        cp.category = ''
        cp.identifier = ''
        cp.language = ''
        cp.version = ''
        cp.content_status = ''
        cp.revision = 1
        # Save
        d.save(path)
        return True
    except Exception as e:
        return f'ERR: {e}'


def deep_clean_xml(path):
    """Open docx zip + clean core.xml + app.xml + custom.xml."""
    try:
        # Make temp copy
        temp = path + '.tmp'
        with zipfile.ZipFile(path, 'r') as zin:
            with zipfile.ZipFile(temp, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.namelist():
                    data = zin.read(item)
                    # Clean core.xml
                    if item == 'docProps/core.xml':
                        s = data.decode('utf-8', errors='replace')
                        # Strip values between common tags
                        import re
                        tags = ['creator', 'lastModifiedBy', 'title', 'subject',
                                'keywords', 'description', 'category',
                                'identifier', 'language', 'version', 'contentStatus']
                        for tag in tags:
                            s = re.sub(rf'(<dc:{tag}[^>]*>).*?(</dc:{tag}>)', r'\1\2', s)
                            s = re.sub(rf'(<cp:{tag}[^>]*>).*?(</cp:{tag}>)', r'\1\2', s)
                        data = s.encode('utf-8')
                    # Clean app.xml (Company, Manager, etc.)
                    elif item == 'docProps/app.xml':
                        s = data.decode('utf-8', errors='replace')
                        import re
                        tags = ['Company', 'Manager', 'HyperlinkBase',
                                'Template', 'TotalTime', 'Application']
                        for tag in tags:
                            s = re.sub(rf'(<{tag}[^>]*>).*?(</{tag}>)', r'\1\2', s)
                        data = s.encode('utf-8')
                    zout.writestr(item, data)
        shutil.move(temp, path)
        return True
    except Exception as e:
        if os.path.exists(temp):
            os.remove(temp)
        return f'ERR: {e}'


print('=== CLEANING METADATA + WATERMARK ===')
print()
for f in FILES:
    full = os.path.join(ROOT, f)
    if not os.path.exists(full):
        print(f'  SKIP (not found): {f}')
        continue

    # Step 1: clean via python-docx
    r1 = clean_via_docx(full)
    # Step 2: deep clean XML for stubborn fields
    r2 = deep_clean_xml(full)

    status = 'OK' if r1 is True and r2 is True else f'{r1} | {r2}'
    print(f'  {status} | {f}')


# ============================================================
# Verify
# ============================================================
print()
print('=== VERIFY METADATA CLEANED ===')
for f in FILES:
    full = os.path.join(ROOT, f)
    if not os.path.exists(full):
        continue
    d = Document(full)
    cp = d.core_properties
    has_any = any([cp.author, cp.title, cp.subject, cp.keywords,
                   cp.comments, cp.last_modified_by, cp.category])
    fn = f.split('/')[-1][:60]
    status = 'CLEAN' if not has_any else 'STILL HAS DATA'
    print(f'  {status} | {fn}')
    if has_any:
        print(f'    author="{cp.author}" title="{cp.title}" last_mod="{cp.last_modified_by}"')
