# -*- coding: utf-8 -*-
"""Rename original PDFs to match translation filenames for easy matching."""
import os, sys, io, shutil
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
os.chdir(ROOT)

# Map translation basename → (current PDF path, new PDF name)
MAPPING = [
    # (current_path, new_name — None means keep)
    ('02_Papers-goc/11-bai-ban-dau-va-mo-rong/01_Jenkins_et_al_2023.pdf', None),
    ('02_Papers-goc/Viet-Nam/Hoa_2024_Frontiers_LoAu_THPT_HaNoi.pdf', '02_Hoa_2024_Frontiers.pdf'),
    ('02_Papers-goc/Dong-Nam-A/GBD_2021_ASEAN_Lancet_2025.pdf', '03_GBD_ASEAN_2025_Lancet.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/Zhameden_2025_School_Interventions_LMIC.pdf', '04_Zhameden_2025_PLOSONE.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/Anderson_2025_Contributing_Factors_Anxiety.pdf', '05_Anderson_2025_Wiley.pdf'),
    ('02_Papers-goc/Viet-Nam/VNAMHS_2022_Report.pdf', '06_VNAMHS_2022.pdf'),
    ('02_Papers-goc/The-gioi_Khac/12889_2024_Article_21252.pdf', '07_Zhu_2025_BMC_China.pdf'),
    ('02_Papers-goc/Dong-Nam-A/PIIS2772368225000034.pdf', '08_Mudunna_2025_LancetSEA.pdf'),
    ('02_Papers-goc/Dong-Nam-A/S2054425125000391a.pdf', '09_Puyat_2025_Filipino.pdf'),
    ('02_Papers-goc/Viet-Nam/pham-et-al-2024-the-correlation-between-quality-of-care-and-mental-health-and-behavioral-problems-among-vietnamese.pdf', '10_Pham_2024_VN_SocialSupport.pdf'),
    ('02_Papers-goc/11-bai-ban-dau-va-mo-rong/02_Saikia_et_al_2023.pdf', '11_Saikia_2023_IJCM.pdf'),
    ('02_Papers-goc/The-gioi_Khac/Wen_2020_LPA_Anxiety_Rural_China.pdf', '12_Wen_2020_IJERPH.pdf'),
    ('02_Papers-goc/11-bai-ban-dau-va-mo-rong/10_Xu_et_al_2021.pdf', '13_Xu_2021_JAffectDisord.pdf'),
    ('02_Papers-goc/Viet-Nam/HoangTrungHoc_2025_COVID_VN.pdf', '14_HoangTrungHoc_2025_AJPR.pdf'),
    ('02_Papers-goc/Viet-Nam/1-s2.0-S2666915324000817-main.pdf', '15_NgoAnhVinh_2024_JAffectDisordReports.pdf'),
    ('02_Papers-goc/Viet-Nam/2617-Văn bản của bài báo-2387-1-10-20250619.pdf', '16_BaoQuyen_2025_YHCD.pdf'),
    ('02_Papers-goc/Viet-Nam/67-70-2948-5466_Văn bản của bài báo.pdf', '17_NguyenDanhLam_2022_YHVN.pdf'),
    ('02_Papers-goc/Viet-Nam/8-13506-23239_Văn bản của bài báo.pdf', '18_AnGiang_2025_YHVN.pdf'),
    # VN19 — actually Hồ Thị Trúc Quỳnh (Hue), the summary VN19 calls her Trần Thảo Vi mistakenly — let me check
    ('02_Papers-goc/Viet-Nam/HoThiTrucQuynh_2025_VN19_Hue_LoAu_LacQuan.pdf', '19_HoThiTrucQuynh_2025_TLH_Hue.pdf'),
    ('02_Papers-goc/Viet-Nam/TranHoVinhLoc_2024_VN20_DASS-Y_TPHCM.pdf', '20_TranHoVinhLoc_2024_YHTPHCM.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/1-s2.0-S0277953625008597-main.pdf', '21_Norway_2025_SocSciMed.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/British J Clinic Psychol - 2025 - Li - Cross‐sectional and longitudinal associations of screen time with adolescent.pdf', '22_ScreenTime_2025_BJCP.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/PIIS0890856724013571.pdf', '23_JAACAP_US_Trends_2024.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/PIIS2666776225002510.pdf', '24_WHO_Europe_2025_LancetRegional.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/S2045796025000083a.pdf', '25_EpiPsychSci_2025.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/SN06988.pdf', '26_UK_NHS_2025_Parliament.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/s41562-025-02134-4.pdf', '27_NatureHumanBehav_SocialMedia_2025.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/zugman-et-al-2024-current-and-future-approaches-to-pediatric-anxiety-disorder-treatment.pdf', '28_AJP_PediatricAnxiety_2024.pdf'),
    ('02_Papers-goc/The-gioi_Khac/s12888-025-07227-y.pdf', '29_Li_CBT_NMA_BMCPsych_2025.pdf'),
    ('02_Papers-goc/The-gioi_Khac/1-s2.0-S0165032725009334-main.pdf', '30_GBD_Trends_10-24y_2025.pdf'),
    ('02_Papers-goc/The-gioi_Khac/1-s2.0-S0165032725017574-main.pdf', '31_59Countries_Anxiety_2025.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/EIP-19-0.pdf', '32_Ireland_MyWorld_2024.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/schmidtpersson_2024_oi_240642_1720187574.86716.pdf', '33_JAMA_ScreenMedia_2024.pdf'),
    ('02_Papers-goc/The-gioi_Khac/Korea_MH_Trends_2006-2022_NatSciRep.pdf', '34_Korea_MH_Trends_NatSciRep_2024.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/Social_Anxiety_Young_People_7Countries_2020_PLOSONE.pdf', '35_Jefferies_SocialAnxiety_7Countries_2020.pdf'),
    ('02_Papers-goc/Viet-Nam/jrm-19-279.pdf', '36_TranThaoVi_2024_JRuralMed.pdf'),
    ('02_Papers-goc/The-gioi_Khac/Korea_GAD_MachineLearning_2025.pdf', '37_Moon_Korea_GAD_ML_2025.pdf'),
    ('02_Papers-goc/Viet-Nam/UNICEF_2022_School_Factors_Vietnam.pdf', '38_UNICEF_VN_2022_SchoolFactors.pdf'),
    ('02_Papers-goc/Viet-Nam/medi-102-e33559.pdf', '39_NguyenLX_VN_COVID_Medicine_2023.pdf'),
    ('02_Papers-goc/Viet-Nam/92-10696-18832_Văn bản của bài báo.pdf', '40_NguyenThanhTruyen_VinhLong_YHVN_2024.pdf'),
    # Note: Praptomojati PDF may not exist in our system; keep as is
    ('02_Papers-goc/The-gioi_Khac/s13034-024-00799-9.pdf', '42_DeSilva_SriLanka_CBT_RCT_2024.pdf'),
    ('02_Papers-goc/The-gioi_Khac/1-s2.0-S0165032724013156-main.pdf', '43_Xian_NMA_SAD_JAD_2024.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/preprint-67067-accepted.pdf', '44_Walder_JMIR_DMHI_SAD_2025.pdf'),
    ('02_Papers-goc/Viet-Nam/O2401115.pdf', '45_PhamThiNgoc_HaiPhong_VinhBao_2024.pdf'),
    ('02_Papers-goc/The-gioi_Khac/prbm-18-1571.pdf', '46_Zheng_MXH_Stress_Sleep_PRBM_2025.pdf'),
    ('02_Papers-goc/Viet-Nam/TranDucSi_2025_LongAn_DASS_PNT.pdf', '47_TranDucSi_LongAn_PNT_2025.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/School based interventions for depression and anxiety in UK.pdf', '48_BrownCarter_UK_School_JMH_2025.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/bress_2024_oi_240871_1723229502.27231.pdf', '49_Bress_JAMA_Maya_App_2024.pdf'),
    ('02_Papers-goc/The-gioi_Khac/fpsyt-16-1594658.pdf', '50_Cai_Resilience_Frontiers_2025.pdf'),
    ('02_Papers-goc/The-gioi_Khac/pediatrics-2024-1-e55786.pdf', '51_Sasaki_Japan_iCBT_JMIR_2024.pdf'),
    ('02_Papers-goc/The-gioi_Au-My-Uc/s10578-024-01667-5.pdf', '52_Jagiello_AcademicStress_SR_2025.pdf'),
    ('02_Papers-goc/Viet-Nam/School_Factors_Causing_Vietnamese_Adolescents_Anx.pdf', '53_Dinh_SchoolFactors_VN_2021.pdf'),
    ('02_Papers-goc/The-gioi_Khac/journal.pone.0328785.pdf', '54_Dong_PLOS_DASS_YaAn_2025.pdf'),
    ('02_Papers-goc/Viet-Nam/TCNCYH_2025_SinhVien_LoAu_TramCam.pdf', '59_DaoThiNgoan_TCNCYH_SVY4_HMU_2025.pdf'),
    ('02_Papers-goc/Viet-Nam/Duong_2025_SocPsychiatry_2631HS_TPHCM.pdf', '60_Duong_SocPsychiatry_2631HS_TPHCM_2025.pdf'),
]

renamed = 0
skipped = 0
missing = 0

for src, new_name in MAPPING:
    if not os.path.exists(src):
        print(f'  MISSING: {os.path.basename(src)[:60]}')
        missing += 1
        continue
    if new_name is None:
        skipped += 1
        continue
    dirname = os.path.dirname(src)
    dst = os.path.join(dirname, new_name)
    if os.path.exists(dst):
        if os.path.getsize(dst) == os.path.getsize(src):
            print(f'  DONE: {new_name[:55]}')
            skipped += 1
        else:
            print(f'  CONFLICT: {new_name[:55]}')
        continue
    try:
        shutil.move(src, dst)
        renamed += 1
        print(f'  [+] {os.path.basename(src)[:40]} -> {new_name[:55]}')
    except Exception as e:
        print(f'  ERR: {new_name[:40]}: {e}')

print(f'\n=== RESULT: {renamed} renamed, {skipped} already OK, {missing} missing ===')
