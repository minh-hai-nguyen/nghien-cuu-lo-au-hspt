# -*- coding: utf-8 -*-
"""Final comprehensive matching — author + co-author + year smart matching."""
import os, sys, io, re, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
with open(os.path.join(ROOT, '02_Papers-goc', 'canonical_index.json'), encoding='utf-8') as f:
    canon = json.load(f)

# Full manual map based on context reading of report
# Each tuple: (citation_in_report, canonical_id, paper_pdf, notes)
MAP = [
    # === Exact first-author + year matches ===
    ('Bress 2024 (Maya app)', 'QT043', 'QT043_Bress_JAMA_Maya_App_2024.pdf', 'd=1,04 RCT Maya app'),
    ('Brown & Carter 2025 (BESST UK)', 'QT042', 'QT042_BrownCarter_UK_School_JMH_2025.pdf', 'BESST (CBT self-referral) 900 HS'),
    ('Brunborg 2025 (Norway social media)', 'QT021', 'QT021_Norway_Brunborg_2025_SocSciMed.pdf', 'Longitudinal Norway'),
    ('Cho 2024 (Korea trends)', 'QT034', 'QT034_Korea_Cho_MH_Trends_NatSciRep_2024.pdf', 'Korea mental health trends'),
    ('Crisp 2025 (Australia epi)', 'QT025', 'QT025_EpiPsychSci_Crisp_Australia_2025.pdf', 'Australia epi panel study'),
    ('Dinh 2021 (VN school factors)', 'VN027', 'VN027_Dinh_SchoolFactors_VN_2021.pdf', 'School factors VN'),
    ('Dong, Wang & Lin 2025 (YaAn)', 'QT047', 'QT047_Dong_PLOS_DASS_YaAn_2025.pdf', 'DASS Ya An China'),
    ('Duong 2025 (TPHCM DASS)', 'VN029', 'VN029_Duong_SocPsychiatry_2631HS_TPHCM_2025.pdf', 'TPHCM DASS-21 + YBRS'),
    ('Erskine 2021 (NAMHS protocol)', None, None, 'Protocol NAMHS 3 countries — cited via VN002 refs'),
    ('Fitzgerald 2024 (Ireland)', 'QT032', 'QT032_Ireland_Fitzgerald_MyWorld_2024.pdf', 'Ireland MyWorld study'),
    ('Fassi 2025 (social media)', 'QT027', 'QT027_NatureHumanBehav_Fassi_SocialMedia_2025.pdf', 'Nature Human Behav'),
    ('GBD 2025 (ASEAN)', 'QT012', 'QT012_GBD_ASEAN_2025_Lancet.pdf', 'GBD ASEAN Lancet'),
    ('Hoa 2024 (Hà Nội GAD-7)', 'VN001', 'VN001_Hoa_2024_Frontiers_HaNoi.pdf', 'GAD-7 Hà Nội'),
    ('Hoàng Trung Học 2025 (COVID VN)', 'VN014', 'VN014_HoangTrungHoc_2025_VN_COVID.pdf', 'VN post-COVID'),
    ('Islam 2025 (59 countries)', 'QT031', 'QT031_Islam_59Countries_2025.pdf', 'Multi-country GSHS'),
    ('Jefferies 2020 (social anxiety 7 countries)', 'QT035', 'QT035_Jefferies_SocialAnxiety_7Countries_2020.pdf', '7 countries social anxiety'),
    ('Jenkins 2023 (USA)', 'QT001', 'QT001_Jenkins_2023_USA_SanDiego.pdf', 'San Diego post-COVID'),
    ('Lộc 2024 (TPHCM DASS-Y)', 'VN020', 'VN020_TranHoVinhLoc_2024_TPHCM_DASS-Y.pdf', 'TPHCM DASS-Y'),
    ('Menon 2025 (LMIC scoping)', 'QT051', 'QT051_Menon_LMIC_SEA_Pacific_APJPH_2025_ABSTRACT', '69 NCs 12 LMICs (abstract only)'),
    ('Moon 2025 (Korea ML GAD)', 'QT036', 'QT036_Moon_Korea_GAD_ML_2025.pdf', 'Korea ML anxiety'),
    ('Ngô Anh Vĩnh 2024 (Lạng Sơn DTTS)', 'VN015', 'VN015_NgoAnhVinh_2024_LangSon_DTTS.pdf', 'Lạng Sơn DTTS'),
    ('Nguyễn Danh Lâm 2022 (YHVN)', 'VN017', 'VN017_NguyenDanhLam_2022_YHVN.pdf', 'YHVN 2022'),
    ('Praptomojati & Hartanto 2024 (SEA SR)', 'QT037', 'QT037 (Praptomojati_CA-CBT_SEA_AJP_2024 — no PDF yet?)', '7 NC CA-CBT SEA; 0 from VN'),
    ('Puyat 2025 (Philippines)', 'QT017', 'QT017_Puyat_2025_Filipino_Youth.pdf', 'Filipino youth'),
    ('Qiaochu & Wang 2025 (mobile CBT)', 'QT050', 'QT050 (Qiaochu_MobileCBT_CPP_2025_ABSTRACT — abstract only)', 'Mobile CBT SR — abstract only'),
    ('Sasaki 2024 (Japan iCBT)', 'QT045', 'QT045_Sasaki_Japan_iCBT_JMIR_2024.pdf', 'Japan iCBT RCT'),
    ('De Silva 2024 (Sri Lanka RCT)', 'QT038', 'QT038_DeSilva_SriLanka_RCT_2024.pdf', 'Sri Lanka cluster RCT'),
    ('Trần Nguyễn Ngọc 2018 (luận án Bạch Mai)', None, '(raw) TranNguyenNgoc_2018_LuanAn_TSYH_BachMai.pdf', 'Luận án TSYH Bạch Mai — chưa canonical hoá'),
    ('Trần Thảo Vi 2025 (Huế)', 'VN021', 'VN021_TranThaoVi_2024_JRuralMed_Hue.pdf', 'Note: cite là 2025 nhưng paper là 2024; cần xác nhận'),
    ('VNAMHS 2022', 'VN002', 'VN002_VNAMHS_2022_National.pdf', 'National survey'),
    ('Walder 2025 (JMIR SAD)', 'QT040', 'QT040_Walder_JMIR_DMHI_SAD_2025.pdf', 'JMIR DMHI for SAD'),
    ('Wen 2020 (China rural)', 'QT008', 'QT008_Wen_et_al_2020_China_Rural.pdf', 'China rural'),
    ('Xian, Zhang & Jiang 2024 (NMA SAD)', 'QT039', 'QT039_Xian_NMA_SAD_JAD_2024.pdf', 'NMA SAD ranking'),
    ('Xu 2021 (China largest epi)', 'QT010', 'QT010_Xu_et_al_2021_China_LargestEpi.pdf', 'China N=~76k'),
    ('Zhang, Liang & Kang 2026 (Bayesian MA)', 'QT049', 'QT049 (Zhang_Bayesian_CBT_JCP_2026_ABSTRACT — abstract only)', 'Universal CBT low quality evidence'),
    ('Zhameden 2025 (PLOS LMIC)', 'QT013', 'QT013_Zhameden_2025_PLOSONE_LMIC.pdf', 'LMIC PLOS'),
    ('Zugman 2024 (AJP pediatric anxiety)', 'QT028', 'QT028_AJP_Zugman_PediatricAnxiety_2024.pdf', 'AJP treatment review'),
    ('Bảo Quyên 2025 (YHCD)', 'VN016', 'VN016_BaoQuyen_2025_YHCD.pdf', 'YHCD 2025'),
    ('An Giang 2025 (YHVN)', 'VN018', 'VN018_AnGiang_2025_YHVN.pdf', 'An Giang province'),
    ('Hồ Thị Trúc Quỳnh 2025 (Huế TLH)', 'VN019', 'VN019_HoThiTrucQuynh_2025_Hue_TLH.pdf', 'Huế TLH'),
    ('Nguyễn L.X. 2023 (COVID Medicine)', 'VN023', 'VN023_NguyenLX_VN_COVID_Medicine_2023.pdf', 'VN COVID'),
    ('Nguyễn Thanh Truyền 2024 (Vĩnh Long)', 'VN024', 'VN024_NguyenThanhTruyen_VinhLong_YHVN_2024.pdf', 'Vĩnh Long'),
    ('Phạm Thị Ngọc 2024 (Hải Phòng)', 'VN025', 'VN025_PhamThiNgoc_HaiPhong_VinhBao_2024.pdf', 'Hải Phòng'),
    ('Trần Đức Sĩ 2025 (Long An PNT)', 'VN026', 'VN026_TranDucSi_LongAn_PNT_2025.pdf', 'Long An'),
    ('Đào Thị Ngoan 2025 (HMU)', 'VN028', 'VN028_DaoThiNgoan_TCNCYH_SVY4_HMU_2025.pdf', 'HMU sinh viên năm 4'),
    ('Tran Happy House 2023', 'VN030', '(no PDF in library)', 'Cluster controlled trial Hà Nội'),
    ('Pham 2024 (Huế social support)', 'VN003', 'VN003_Pham_2024_Hue_SocialSupport.pdf', 'Huế social support'),

    # === Likely matches (typo / name variation) ===
    ('Cao 2025 (Frontiers MA resilience)', 'QT044', 'QT044_Cai_Resilience_Frontiers_2025.pdf', 'LIKELY: "Cao 2025" typo → Cai et al. 2025 Frontiers; verify'),
    ('Li 2025 (NMA CBT)', 'QT029', 'QT029_BMC_Li_CBT_NMA_2025.pdf', 'SUCRA 0,66 for CBT'),

    # === Not in library ===
    ('Brown 2024 (BESST pilot)', None, None, 'BESST pilot UK predecessor paper — not in library yet'),
    ('Kuyken 2022 (MYRIAD mindfulness)', None, None, 'MYRIAD trial UK mindfulness — classic, not in library'),
    ('Walkup 2008 (CAMS NEJM)', None, None, 'CAMS classic RCT children anxiety — not in library, seminal'),
    ('Montgomery 2010', None, None, 'Trích gián tiếp qua Trần Nguyễn Ngọc 2018 — không có bản gốc'),
    ('Chen 2025 (COVID meta)', 'QT048', 'QT048 (Chen_COVID_Meta141_JAD_2025_ABSTRACT — abstract only)', 'COVID meta 141 papers'),
    ('Zhu 2025 (BMC China)', 'QT015', 'QT015_Zhu_2025_BMC_China.pdf', 'BMC China'),
]

# Print table
print(f'{"Citation":55s} | {"ID":6s} | PDF / Notes')
print('='*120)
for cite, cid, pdf, note in MAP:
    cid_str = cid if cid else '─'
    pdf_str = pdf if pdf else note[:60]
    print(f'{cite[:55]:55s} | {cid_str:6s} | {pdf_str}')

# Statistics
matched = sum(1 for _, cid, _, _ in MAP if cid)
not_in_lib = sum(1 for _, cid, _, _ in MAP if not cid)
print()
print(f'Total citations mapped: {len(MAP)}')
print(f'  Matched to canonical ID: {matched}')
print(f'  NOT in library (need to download): {not_in_lib}')
print(f'    - Brown 2024 (BESST pilot UK)')
print(f'    - Kuyken 2022 (MYRIAD UK mindfulness)')
print(f'    - Walkup 2008 (CAMS NEJM)')
print(f'    - Trần Nguyễn Ngọc 2018 (exists as raw PDF but not canonicalized)')

# Save JSON
OUT = os.path.join(os.path.dirname(__file__), 'citations_10042026_final.json')
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(
        [{'citation': c, 'canonical_id': i, 'pdf': p, 'notes': n} for c, i, p, n in MAP],
        f, ensure_ascii=False, indent=2)
print(f'\nSaved: {OUT}')
