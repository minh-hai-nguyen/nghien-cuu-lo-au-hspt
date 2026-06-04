# -*- coding: utf-8 -*-
"""
RENAME CANONICAL v1 — chuan hoa toan bo ten file thong nhat VN01-29 + QT01-51.

Quy tac:
- VN papers (Viet Nam): VN01-VN29 (theo thu tu lich su, giu nguyen so cu cua tom tat)
- QT papers (Quoc te): QT01-QT11 (orig 11 bai dau) + QT21-QT51 (cac bai sau)
- Co gap QT12-QT20 (de cho 6 bai chuyen tiep nhung khong co tom tat — them moi)

Mapping based on existing Tom-tat-tung-bai/ numbering, with:
- Pad VN1/2/3 -> VN01/02/03
- Add QT prefix to 01-11
- Fill QT12-QT17 for 6 papers without summaries (GBD ASEAN, Zhameden, Anderson, Zhu, Mudunna, Puyat)
"""
import os, sys, io, shutil, re
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
os.chdir(ROOT)

# ============================================================
# CANONICAL MAPPING
# ============================================================
# (translation_number_old, canonical_id, short_descriptor)
# canonical_id is the new prefix like "QT01" or "VN15"

CANONICAL = [
    # (old_trans_num, canonical, descriptor)
    (1,  'QT01', 'Jenkins_2023_USA_SanDiego'),
    (2,  'VN01', 'Hoa_2024_Frontiers_HaNoi'),
    (3,  'QT12', 'GBD_ASEAN_2025_Lancet'),  # NEW QT12 (filling gap)
    (4,  'QT13', 'Zhameden_2025_PLOSONE_LMIC'),  # NEW QT13
    (5,  'QT14', 'Anderson_2025_Wiley_Narrative'),  # NEW QT14
    (6,  'VN02', 'VNAMHS_2022_National'),
    (7,  'QT15', 'Zhu_2025_BMC_China'),  # NEW QT15
    (8,  'QT16', 'Mudunna_2025_LancetSEA_SouthAsia'),  # NEW QT16
    (9,  'QT17', 'Puyat_2025_Filipino_Youth'),  # NEW QT17
    (10, 'VN03', 'Pham_2024_Hue_SocialSupport'),
    (11, 'QT02', 'Saikia_2023_India_Assam'),
    (12, 'QT08', 'Wen_2020_China_Rural_LPA'),
    (13, 'QT10', 'Xu_2021_China_LargestEpi'),
    (14, 'VN14', 'HoangTrungHoc_2025_VN_COVID'),
    (15, 'VN15', 'NgoAnhVinh_2024_LangSon_DTTS'),
    (16, 'VN16', 'BaoQuyen_2025_YHCD'),
    (17, 'VN17', 'NguyenDanhLam_2022_YHVN'),
    (18, 'VN18', 'AnGiang_2025_YHVN'),
    (19, 'VN19', 'HoThiTrucQuynh_2025_Hue_TLH'),  # FIX: actual author is Ho Thi Truc Quynh
    (20, 'VN20', 'TranHoVinhLoc_2024_TPHCM_DASS-Y'),
    (21, 'QT21', 'Norway_Brunborg_2025_SocSciMed'),
    (22, 'QT22', 'ScreenTime_Li_2025_BJCP'),
    (23, 'QT23', 'JAACAP_USA_Mojtabai_Trends_2024'),
    (24, 'QT24', 'WHO_Europe_2025_LancetRegional'),
    (25, 'QT25', 'EpiPsychSci_Crisp_Australia_2025'),
    (26, 'QT26', 'UK_NHS_Parliament_2025'),
    (27, 'QT27', 'NatureHumanBehav_Fassi_SocialMedia_2025'),
    (28, 'QT28', 'AJP_Zugman_PediatricAnxiety_2024'),
    (29, 'QT29', 'BMC_Li_CBT_NMA_2025'),
    (30, 'QT30', 'GBD_Trends_10-24y_2025'),
    (31, 'QT31', 'Islam_59Countries_2025'),
    (32, 'QT32', 'Ireland_Fitzgerald_MyWorld_2024'),
    (33, 'QT33', 'JAMA_SchmidtPersson_ScreenMedia_2024'),
    (34, 'QT34', 'Korea_Cho_MH_Trends_NatSciRep_2024'),
    (35, 'QT35', 'Jefferies_SocialAnxiety_7Countries_2020'),
    (36, 'VN21', 'TranThaoVi_2024_JRuralMed_Hue'),
    (37, 'QT36', 'Moon_Korea_GAD_ML_2025'),
    (38, 'VN22', 'UNICEF_VN_2022_SchoolFactors'),
    (39, 'VN23', 'NguyenLX_VN_COVID_Medicine_2023'),
    (40, 'VN24', 'NguyenThanhTruyen_VinhLong_YHVN_2024'),
    (41, 'QT37', 'Praptomojati_CA-CBT_SEA_AJP_2024'),
    (42, 'QT38', 'DeSilva_SriLanka_RCT_2024'),
    (43, 'QT39', 'Xian_NMA_SAD_JAD_2024'),
    (44, 'QT40', 'Walder_JMIR_DMHI_SAD_2025'),
    (45, 'VN25', 'PhamThiNgoc_HaiPhong_VinhBao_2024'),
    (46, 'QT41', 'Zheng_China_MXH_PRBM_2025'),
    (47, 'VN26', 'TranDucSi_LongAn_PNT_2025'),
    (48, 'QT42', 'BrownCarter_UK_School_JMH_2025'),
    (49, 'QT43', 'Bress_JAMA_Maya_App_2024'),
    (50, 'QT44', 'Cai_Resilience_Frontiers_2025'),
    (51, 'QT45', 'Sasaki_Japan_iCBT_JMIR_2024'),
    (52, 'QT46', 'Jagiello_AcademicStress_SR_2025'),
    (53, 'VN27', 'Dinh_SchoolFactors_VN_2021'),
    (54, 'QT47', 'Dong_PLOS_DASS_YaAn_2025'),
    (55, 'QT48', 'Chen_COVID_Meta141_JAD_2025_ABSTRACT'),
    (56, 'QT49', 'Zhang_Bayesian_CBT_JCP_2026_ABSTRACT'),
    (57, 'QT50', 'Qiaochu_MobileCBT_CPP_2025_ABSTRACT'),
    (58, 'QT51', 'Menon_LMIC_SEA_Pacific_APJPH_2025_ABSTRACT'),
    (59, 'VN28', 'DaoThiNgoan_TCNCYH_SVY4_HMU_2025'),
    (60, 'VN29', 'Duong_SocPsychiatry_2631HS_TPHCM_2025'),
]

# Print mapping for verification
print('=== CANONICAL MAPPING ===')
print(f'{"Old":<6} {"New":<6} {"Descriptor"}')
print('-' * 70)
for old, new, desc in CANONICAL:
    print(f'  {old:<4} {new:<6} {desc}')

# Check no duplicates
new_ids = [n for _, n, _ in CANONICAL]
assert len(new_ids) == len(set(new_ids)), 'Duplicate canonical IDs!'
print(f'\nTotal: {len(CANONICAL)} unique canonical IDs')

# Save mapping to JSON for reference
import json
mapping_path = os.path.join(os.path.dirname(__file__), 'canonical_mapping.json')
with open(mapping_path, 'w', encoding='utf-8') as f:
    json.dump([{'old_num': o, 'canonical': n, 'descriptor': d} for o, n, d in CANONICAL],
              f, ensure_ascii=False, indent=2)
print(f'\nMapping saved: {mapping_path}')
