# -*- coding: utf-8 -*-
"""Dedup 17 canonical IDs in 03_Ban-dich and Tom-tat-tung-bai.
Keep _FIXED_27052026 (or newest), archive older versions to _Archive/.
Then strip metadata of kept files.
"""
import os, sys, io, shutil, zipfile, re
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = r'c:\Users\OS\OneDrive\read_books\Lo-au'
BAN_DICH = os.path.join(ROOT, '03_Ban-dich')
TOM_TAT  = os.path.join(ROOT, 'Tom-tat-tung-bai')
ARCH_BD  = os.path.join(ROOT, '_Archive', 'BanDich_dedup_07062026')
ARCH_TT  = os.path.join(ROOT, '_Archive', 'TomTat_dedup_07062026')
os.makedirs(ARCH_BD, exist_ok=True)
os.makedirs(ARCH_TT, exist_ok=True)

# Dedup plan: { source_dir: [ (keep_file, [archive_files...]), ... ] }
PLAN = {
    BAN_DICH: [
        ('DICH_Anderson_2025_FIXED_27052026.docx', ['DICH_Anderson_2025.docx']),
        ('DICH_Hoa_2024_Frontiers_FIXED_27052026.docx', ['DICH_Hoa_2024_Frontiers.docx']),
        ('DICH_VNAMHS_2022_FIXED_27052026.docx', ['DICH_VNAMHS_2022.docx']),
        ('QT001_Jenkins_2023_USA_SanDiego_FIXED_27052026.docx', ['QT001_Jenkins_2023_USA_SanDiego.docx']),
        ('QT010_Xu_2021_China_LargestEpi_FIXED_27052026.docx', ['QT010_Xu_2021_China_LargestEpi.docx']),
        ('QT014_Anderson_2025_Wiley_Narrative_FIXED_27052026.docx', ['QT014_Anderson_2025_Wiley_Narrative.docx']),
        ('QT021_Norway_Brunborg_2025_SocSciMed_FIXED_27052026.docx', ['QT021_Norway_Brunborg_2025_SocSciMed.docx']),
        ('QT022_ScreenTime_Li_2025_BJCP_FIXED_27052026.docx', ['QT022_ScreenTime_Li_2025_BJCP.docx']),
        ('QT026_UK_NHS_Parliament_2025_FIXED_27052026.docx', ['QT026_UK_NHS_Parliament_2025.docx']),
        ('QT028_AJP_Zugman_PediatricAnxiety_2024_FIXED_27052026.docx', ['QT028_AJP_Zugman_PediatricAnxiety_2024.docx']),
        ('QT035_Jefferies_SocialAnxiety_7Countries_2020_FIXED_27052026.docx', ['QT035_Jefferies_SocialAnxiety_7Countries_2020.docx']),
        ('QT036_Moon_Korea_GAD_ML_2025_FIXED_27052026.docx', ['QT036_Moon_Korea_GAD_ML_2025.docx']),
        ('QT043_Bress_JAMA_Maya_App_2024_FIXED_27052026.docx', ['QT043_Bress_JAMA_Maya_App_2024.docx']),
        ('VN001_Hoa_2024_Frontiers_HaNoi_FIXED_27052026.docx', ['VN001_Hoa_2024_Frontiers_HaNoi.docx']),
        # VN002: 5 versions. Keep FIXED_27052026. Archive base, v1_backup, v1_backup_FIXED.
        # FULL is flagged for manual review — NOT archived.
        ('VN002_VNAMHS_2022_National_FIXED_27052026.docx', [
            'VN002_VNAMHS_2022_National.docx',
            'VN002_VNAMHS_2022_National_v1_backup.docx',
            'VN002_VNAMHS_2022_National_v1_backup_FIXED_27052026.docx',
        ]),
        ('VN017_NguyenDanhLam_2022_YHVN_FIXED_27052026.docx', ['VN017_NguyenDanhLam_2022_YHVN.docx']),
        ('VN018_AnGiang_2025_YHVN_FIXED_27052026.docx', ['VN018_AnGiang_2025_YHVN.docx']),
        ('VN023_NguyenLX_VN_COVID_Medicine_2023_FIXED_27052026.docx', ['VN023_NguyenLX_VN_COVID_Medicine_2023.docx']),
    ],
    TOM_TAT: [
        ('00_Mẫu tóm tắt bài 1_FIXED_27052026.docx', ['00_Mẫu tóm tắt bài 1.docx']),
        ('QT003_Mandaknalli_Malusare_2021_FIXED_27052026.docx', ['QT003_Mandaknalli_Malusare_2021.docx']),
        ('QT005_Alharbi_et_al_2019_SaudiArabia_FIXED_27052026.docx', ['QT005_Alharbi_et_al_2019_SaudiArabia.docx']),
        ('QT007_Chen_et_al_2023_China_BMCPsych_FIXED_27052026.docx', ['QT007_Chen_et_al_2023_China_BMCPsych.docx']),
        ('QT010_Xu_et_al_2021_China_LargestEpi_FIXED_27052026.docx', ['QT010_Xu_et_al_2021_China_LargestEpi.docx']),
        # QT011: keep current base (newer fix44 applied), archive BEFORE_FIX44 backup
        ('QT011_Bhardwaj_et_al_2020_India_Chandigarh.docx', ['QT011_Bhardwaj_et_al_2020_India_Chandigarh_BEFORE_FIX44_01062026.docx']),
        ('QT035_Jefferies_SocialAnxiety_7Countries_2020_FIXED_27052026.docx', ['QT035_Jefferies_SocialAnxiety_7Countries_2020.docx']),
        ('VN001_Hoa_2024_Frontiers_HaNoi_FIXED_27052026.docx', ['VN001_Hoa_2024_Frontiers_HaNoi.docx']),
        # VN002 in Tom-tat: only v1_backup is duplicate; keep the newer base (Jun 2)
        ('VN002_VNAMHS_2022_National.docx', ['VN002_VNAMHS_2022_National_v1_backup.docx']),
        ('VN017_NguyenDanhLam_2022_YHVN_FIXED_27052026.docx', ['VN017_NguyenDanhLam_2022_YHVN.docx']),
        ('VN018_AnGiang_2025_YHVN_FIXED_27052026.docx', ['VN018_AnGiang_2025_YHVN.docx']),
    ],
}

ARCHIVE_MAP = {BAN_DICH: ARCH_BD, TOM_TAT: ARCH_TT}

results = {BAN_DICH: {'ids': 0, 'archived': []}, TOM_TAT: {'ids': 0, 'archived': []}}

print('=== STEP 1: Archive older versions ===')
for src_dir, entries in PLAN.items():
    arch_dir = ARCHIVE_MAP[src_dir]
    for keep, archive_list in entries:
        keep_path = os.path.join(src_dir, keep)
        if not os.path.exists(keep_path):
            print(f'  MISSING KEEP FILE: {keep}')
            continue
        results[src_dir]['ids'] += 1
        for arch in archive_list:
            src = os.path.join(src_dir, arch)
            dst = os.path.join(arch_dir, arch)
            if not os.path.exists(src):
                print(f'  SKIP (not found): {arch}')
                continue
            if os.path.exists(dst):
                print(f'  WARN (already in archive): {arch}')
                continue
            try:
                shutil.move(src, dst)
                results[src_dir]['archived'].append(arch)
                print(f'  MOVED: {os.path.basename(src_dir)}/{arch} -> {os.path.basename(arch_dir)}/')
            except Exception as e:
                print(f'  ERROR moving {arch}: {e}')


# ---- STEP 2: Strip metadata of kept files ----
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


print('\n=== STEP 2: Strip metadata of kept files ===')
strip_ok = strip_locked = strip_failed = 0
for src_dir, entries in PLAN.items():
    for keep, _ in entries:
        keep_path = os.path.join(src_dir, keep)
        if not os.path.exists(keep_path):
            continue
        res = strip_one(keep_path)
        if res is True: strip_ok += 1
        elif res is None:
            strip_locked += 1
            print(f'  LOCKED: {keep}')
        else: strip_failed += 1

print(f'\nStrip results: OK={strip_ok}, Locked={strip_locked}, Failed={strip_failed}')

# ---- STEP 3: Write report ----
report_path = os.path.join(ROOT, '06_Scripts', 'DEDUP_RESULT_07062026.md')
lines = ['# Dedup Result — 07/06/2026', '']
lines.append(f'**Target folders:** `03_Ban-dich/`, `Tom-tat-tung-bai/`')
lines.append(f'**Archive folders:** `_Archive/BanDich_dedup_07062026/`, `_Archive/TomTat_dedup_07062026/`')
lines.append('')
lines.append('## Per-ID decisions')
lines.append('')

for src_dir, entries in PLAN.items():
    folder_label = os.path.basename(src_dir)
    lines.append(f'### {folder_label}')
    lines.append('')
    lines.append('| Keep | Archived |')
    lines.append('| --- | --- |')
    for keep, archive_list in entries:
        arch_str = '<br>'.join(archive_list) if archive_list else '—'
        lines.append(f'| `{keep}` | {arch_str} |')
    lines.append('')

lines.append('## Totals')
lines.append('')
lines.append(f'- IDs deduplicated in `03_Ban-dich/`: **{results[BAN_DICH]["ids"]}**')
lines.append(f'- Files archived from `03_Ban-dich/`: **{len(results[BAN_DICH]["archived"])}**')
lines.append(f'- IDs deduplicated in `Tom-tat-tung-bai/`: **{results[TOM_TAT]["ids"]}**')
lines.append(f'- Files archived from `Tom-tat-tung-bai/`: **{len(results[TOM_TAT]["archived"])}**')
total_archived = len(results[BAN_DICH]["archived"]) + len(results[TOM_TAT]["archived"])
lines.append(f'- **Total files archived:** {total_archived}')
lines.append('')
lines.append(f'## Metadata strip')
lines.append('')
lines.append(f'- OK: {strip_ok}')
lines.append(f'- Locked (in use): {strip_locked}')
lines.append(f'- Failed: {strip_failed}')
lines.append('')
lines.append('## Manual-review cases')
lines.append('')
lines.append('### VN002 in `03_Ban-dich/` — `VN002_VNAMHS_2022_National_FULL.docx`')
lines.append('')
lines.append('- Kept as canonical: `VN002_VNAMHS_2022_National_FIXED_27052026.docx` (566 KB, has 27052026 audit fixes).')
lines.append('- **NOT archived:** `VN002_VNAMHS_2022_National_FULL.docx` (1003 KB, ~2x larger than FIXED).')
lines.append('- Reason: FULL is significantly larger, may contain extended content not in FIXED. Needs manual content comparison before deciding whether to keep FULL alongside FIXED, replace FIXED with FULL, or merge.')
lines.append('')

with open(report_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print(f'\nReport written: {report_path}')
