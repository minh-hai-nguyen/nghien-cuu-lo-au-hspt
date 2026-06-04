# -*- coding: utf-8 -*-
"""
RENAME APPLY: doi ten thuc te tat ca file theo canonical mapping.

Cac file se rename:
1. Tom-tat-tung-bai/*.docx (61 file: 11 origs + 50 VN/QT)
2. 03_Ban-dich/NN_*.docx (60 file: translation files)
3. 02_Papers-goc/*/NN_*.pdf (60 file: original PDFs already renamed by previous script)
"""
import os, sys, io, json, shutil, re
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
os.chdir(ROOT)

# Load canonical mapping
with open(os.path.join(os.path.dirname(__file__), 'canonical_mapping.json'), encoding='utf-8') as f:
    CANONICAL = [(item['old_num'], item['canonical'], item['descriptor']) for item in json.load(f)]

print(f'Loaded {len(CANONICAL)} canonical mappings')

# ============================================================
# STEP 1: Rename Tom-tat-tung-bai/ summaries
# ============================================================
print('\n=== STEP 1: Rename Tom-tat-tung-bai/ summaries ===')

# Map: current_summary_filename → new_canonical_name
# (current basename without .docx → new basename without .docx)
SUMMARY_RENAMES = {
    # Pad VN1/2/3
    'VN1_Hoa_2024': 'VN01_Hoa_2024_Frontiers_HaNoi',
    'VN2_VNAMHS_2022': 'VN02_VNAMHS_2022_National',
    'VN3_Pham_2024': 'VN03_Pham_2024_Hue_SocialSupport',
    # Add QT prefix to 11 originals
    '01_Jenkins_et_al_2023': 'QT01_Jenkins_et_al_2023_USA',
    '02_Saikia_et_al_2023': 'QT02_Saikia_et_al_2023_India_Assam',
    '03_Mandaknalli_Malusare_2021': 'QT03_Mandaknalli_Malusare_2021',  # orphan, no translation
    '04_NSCH_2020': 'QT04_NSCH_2020_USA',  # orphan
    '05_Alharbi_et_al_2019': 'QT05_Alharbi_et_al_2019_SaudiArabia',  # orphan
    '06_Nakie_et_al_2022': 'QT06_Nakie_et_al_2022_Ethiopia',  # orphan
    '07_Chen_et_al_2023': 'QT07_Chen_et_al_2023_China_BMCPsych',  # orphan
    '08_Wen_et_al_2020': 'QT08_Wen_et_al_2020_China_Rural',
    '09_Qiu_et_al_2022': 'QT09_Qiu_et_al_2022_China_Parenting',  # orphan
    '10_Xu_et_al_2021': 'QT10_Xu_et_al_2021_China_LargestEpi',
    '11_Bhardwaj_et_al_2020': 'QT11_Bhardwaj_et_al_2020_India_Chandigarh',  # orphan
    # VN14-22 already have correct prefix; just enrich descriptor
    'VN14_HoangTrungHoc_2025': 'VN14_HoangTrungHoc_2025_VN_COVID',
    'VN15_NgoAnhVinh_2024': 'VN15_NgoAnhVinh_2024_LangSon_DTTS',
    'VN16_BaoQuyen_2025': 'VN16_BaoQuyen_2025_YHCD',
    'VN17_DanhLam_2022': 'VN17_NguyenDanhLam_2022_YHVN',
    'VN18_AnGiang_2025': 'VN18_AnGiang_2025_YHVN',
    'VN19_TranThaoVi_2025': 'VN19_HoThiTrucQuynh_2025_Hue_TLH',  # FIX: actual author
    'VN20_TranHoVinhLoc_2024': 'VN20_TranHoVinhLoc_2024_TPHCM_DASS-Y',
    'VN21_TranThaoVi_2024': 'VN21_TranThaoVi_2024_JRuralMed_Hue',
    'VN22_UNICEF_2022_SchoolFactors': 'VN22_UNICEF_VN_2022_SchoolFactors',
    'VN23_NguyenLX_2023_Medicine': 'VN23_NguyenLX_VN_COVID_Medicine_2023',
    'VN24_VinhLong_2024': 'VN24_NguyenThanhTruyen_VinhLong_YHVN_2024',
    'VN25_HaiPhong_2024': 'VN25_PhamThiNgoc_HaiPhong_VinhBao_2024',
    'VN26_LongAn_2025': 'VN26_TranDucSi_LongAn_PNT_2025',
    'VN27_Dinh_SchoolFactors_2021': 'VN27_Dinh_SchoolFactors_VN_2021',
    'VN28_DaoThiNgoan_2025_SVY4_HMU': 'VN28_DaoThiNgoan_TCNCYH_SVY4_HMU_2025',
    'VN29_Duong_2025_2631HS_TPHCM': 'VN29_Duong_SocPsychiatry_2631HS_TPHCM_2025',
    # QT21-51 already have prefix, just enrich some
    'QT21_Norway_2025': 'QT21_Norway_Brunborg_2025_SocSciMed',
    'QT22_ScreenTime_2025': 'QT22_ScreenTime_Li_2025_BJCP',
    'QT23_JAACAP_US_2024': 'QT23_JAACAP_USA_Mojtabai_Trends_2024',
    'QT24_WHO_Europe_2025': 'QT24_WHO_Europe_2025_LancetRegional',
    'QT25_EpiPsychSci_2025': 'QT25_EpiPsychSci_Crisp_Australia_2025',
    'QT26_UK_NHS_2025': 'QT26_UK_NHS_Parliament_2025',
    'QT27_Nature_SocialMedia_2025': 'QT27_NatureHumanBehav_Fassi_SocialMedia_2025',
    'QT28_AJP_Treatment_2024': 'QT28_AJP_Zugman_PediatricAnxiety_2024',
    'QT29_CBT_NetworkMeta_2025': 'QT29_BMC_Li_CBT_NMA_2025',
    'QT30_GBD_Trends': 'QT30_GBD_Trends_10-24y_2025',
    'QT31_59Countries': 'QT31_Islam_59Countries_2025',
    'QT32_Ireland': 'QT32_Ireland_Fitzgerald_MyWorld_2024',
    'QT33_JAMA_Screen': 'QT33_JAMA_SchmidtPersson_ScreenMedia_2024',
    'QT34_Korea_Trends': 'QT34_Korea_Cho_MH_Trends_NatSciRep_2024',
    'QT35_SocialAnxiety_7Countries': 'QT35_Jefferies_SocialAnxiety_7Countries_2020',
    'QT36_Korea_GAD_ML_2025': 'QT36_Moon_Korea_GAD_ML_2025',
    'QT37_Praptomojati_CA-CBT_SEA_2024': 'QT37_Praptomojati_CA-CBT_SEA_AJP_2024',
    'QT38_DeSilva_SriLanka_RCT_2024': 'QT38_DeSilva_SriLanka_RCT_2024',
    'QT39_Xian_NMA_SAD_2024': 'QT39_Xian_NMA_SAD_JAD_2024',
    'QT40_Walder_DMHI_SAD_2025': 'QT40_Walder_JMIR_DMHI_SAD_2025',
    'QT41_Zheng_MXH_Stress_Sleep_2025': 'QT41_Zheng_China_MXH_PRBM_2025',
    'QT42_Brown_Carter_UK_2025': 'QT42_BrownCarter_UK_School_JMH_2025',
    'QT43_Bress_Maya_App_2024': 'QT43_Bress_JAMA_Maya_App_2024',
    'QT44_Cai_Resilience_MA_2025': 'QT44_Cai_Resilience_Frontiers_2025',
    'QT45_Sasaki_Japan_iCBT_2024': 'QT45_Sasaki_Japan_iCBT_JMIR_2024',
    'QT46_AcademicStress_SR_2025': 'QT46_Jagiello_AcademicStress_SR_2025',
    'QT47_Dong_PLOS_DASS_2025': 'QT47_Dong_PLOS_DASS_YaAn_2025',
    'QT48_Chen_COVID_Meta141_2025_ABSTRACT': 'QT48_Chen_COVID_Meta141_JAD_2025_ABSTRACT',
    'QT49_Zhang_Bayesian_CBT_2026_ABSTRACT': 'QT49_Zhang_Bayesian_CBT_JCP_2026_ABSTRACT',
    'QT50_Qiaochu_MobileCBT_2025_ABSTRACT': 'QT50_Qiaochu_MobileCBT_CPP_2025_ABSTRACT',
    'QT51_Menon_LMIC_SEA_Pacific_2025_ABSTRACT': 'QT51_Menon_LMIC_SEA_Pacific_APJPH_2025_ABSTRACT',
}

renamed_summary = 0
skipped_summary = 0
for old, new in SUMMARY_RENAMES.items():
    src = os.path.join('Tom-tat-tung-bai', old + '.docx')
    dst = os.path.join('Tom-tat-tung-bai', new + '.docx')
    if not os.path.exists(src):
        print(f'  MISSING: {old}.docx')
        continue
    if os.path.exists(dst):
        print(f'  ALREADY: {new}.docx')
        skipped_summary += 1
        continue
    shutil.move(src, dst)
    renamed_summary += 1

print(f'\nSummary renames: {renamed_summary} done, {skipped_summary} already-OK')

# ============================================================
# STEP 2: Rename 03_Ban-dich/ translations
# ============================================================
print('\n=== STEP 2: Rename 03_Ban-dich/ translations ===')

# Map old translation filename → canonical
# Use the CANONICAL list: (old_num, canonical_id, descriptor)
# Build by finding files matching old_num_*

trans_dir = '03_Ban-dich'
existing_trans = {f for f in os.listdir(trans_dir)
                  if re.match(r'\d{2}_.*\.docx$', f)}

renamed_trans = 0
trans_renames = {}  # old_path → new_path
for old_num, canon, desc in CANONICAL:
    prefix = f'{old_num:02d}_'
    matches = [f for f in existing_trans if f.startswith(prefix)]
    if not matches:
        print(f'  MISSING: trans #{old_num}')
        continue
    if len(matches) > 1:
        print(f'  MULTIPLE: #{old_num}: {matches}')
        continue
    old_name = matches[0]
    new_name = f'{canon}_{desc}.docx'
    src = os.path.join(trans_dir, old_name)
    dst = os.path.join(trans_dir, new_name)
    if os.path.exists(dst):
        print(f'  ALREADY: {new_name}')
        continue
    shutil.move(src, dst)
    trans_renames[old_name] = new_name
    renamed_trans += 1

print(f'\nTranslation renames: {renamed_trans} done')

# ============================================================
# STEP 3: Rename 02_Papers-goc/ originals
# ============================================================
print('\n=== STEP 3: Rename 02_Papers-goc/*/ originals ===')

# These were already renamed in previous script to NN_descriptor format
# Now we rename to canonical format CANONICAL_descriptor
paper_dirs = [
    '02_Papers-goc/Viet-Nam',
    '02_Papers-goc/The-gioi_Au-My-Uc',
    '02_Papers-goc/The-gioi_Khac',
    '02_Papers-goc/Dong-Nam-A',
    '02_Papers-goc/11-bai-ban-dau-va-mo-rong',
]

renamed_pdf = 0
missing_pdf = []
for old_num, canon, desc in CANONICAL:
    prefix = f'{old_num:02d}_'
    found = None
    for d in paper_dirs:
        if not os.path.exists(d):
            continue
        for f in os.listdir(d):
            if f.endswith('.pdf') and f.startswith(prefix):
                found = (d, f)
                break
        if found: break
    if not found:
        # Maybe the file already has canonical prefix
        for d in paper_dirs:
            if not os.path.exists(d):
                continue
            for f in os.listdir(d):
                if f.endswith('.pdf') and f.startswith(canon + '_'):
                    found = (d, f)
                    break
            if found: break
    if not found:
        missing_pdf.append((old_num, canon, desc))
        continue
    d, old_name = found
    if old_name.startswith(canon + '_'):
        continue  # already canonical
    new_name = f'{canon}_{desc}.pdf'
    src = os.path.join(d, old_name)
    dst = os.path.join(d, new_name)
    if os.path.exists(dst):
        continue
    shutil.move(src, dst)
    renamed_pdf += 1

print(f'\nPDF renames: {renamed_pdf} done')
if missing_pdf:
    print(f'Missing PDFs: {len(missing_pdf)}')
    for o, c, d in missing_pdf[:10]:
        print(f'  #{o} ({c}) - {d}')

# Save translation rename map for use in script updater
import json
with open(os.path.join(os.path.dirname(__file__), 'translation_renames.json'), 'w', encoding='utf-8') as f:
    json.dump(trans_renames, f, ensure_ascii=False, indent=2)

print('\n=== DONE ===')
