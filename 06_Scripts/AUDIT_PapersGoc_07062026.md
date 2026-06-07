# AUDIT 02_Papers-goc/ vs CSDL + KG + RAG (07/06/2026)
## Tổng quan
- **Tổng số PDF trong `02_Papers-goc/`:** 160
- **canonical_index.json:** 117 entries (trong đó 100 có file PDF)
- **KG v2 papers:** 75 nodes
- **PDF không có trong canonical:** 61
- **Canonical trỏ tới PDF KHÔNG tồn tại:** 1

## Phân bố PDF theo folder
| Folder | Số PDF |
|---|---|
| `Chua-phan-loai` | 35 |
| `The-gioi_Au-My-Uc` | 35 |
| `The-gioi_Khac` | 25 |
| `Viet-Nam` | 23 |
| `11-bai-ban-dau-va-mo-rong` | 15 |
| `(root)` | 8 |
| `Dong-Nam-A` | 5 |
| `Viet-Nam\_Archive_LoiFile` | 4 |
| `Chua-phan-loai\Viet-nam` | 3 |
| `Chua-phan-loai\tai-them-27052026` | 2 |
| `Coping_Effectiveness` | 2 |
| `GBD_WHO` | 2 |
| `UK_BESST_PLACES` | 2 |
| `The-gioi-moi` | 1 |

## PDF chưa có trong canonical_index (61)
Cần phân loại + thêm vào canonical_index.json:

- `Dong-Nam-A\1-s2.0-S1876201823004537-main.pdf`
- `Chua-phan-loai\1-s2.0-S266691532500054X-main.pdf`
- `Chua-phan-loai\Viet-nam\16.-NguyenThiVan.pdf`
- `Viet-Nam\2617-Văn bản của bài báo-2387-1-10-20250619 (1).pdf`
- `Chua-phan-loai\Bhardwaj_2020_Chandigarh_DASS21.pdf`
- `The-gioi-moi\British J Clinic Psychol - 2025 - Li - Cross‐sectional and longitudinal associations of screen time with adolescent.pdf`
- `UK_BESST_PLACES\Brown_2022_PLACES_IJERPH.pdf`
- `UK_BESST_PLACES\Brown_2024_BESST_Lancet_Psychiatry.pdf`
- `Chua-phan-loai\CBT_Delivery_GAD_TranslPsych_2025.pdf`
- `Chua-phan-loai\Chen_2023_Chinese_SecondarySchool_BMCPsych.pdf`
- `Chua-phan-loai\Child Adoles Psych Nursing - 2025 - Anderson - Contributing Factors to the Rise in Adolescent Anxiety and Associated Mental.pdf`
- `Chua-phan-loai\Chronic_Stress_Neuroinflammation_FrontPsych_2023.pdf`
- `The-gioi_Au-My-Uc\Compas_2017_Coping_MetaAnalysis.pdf`
- `Chua-phan-loai\Viet-nam\CVv443S402020122.pdf`
- `Chua-phan-loai\Digital_MH_SocialAnxiety_MetaAnalysis_2025.pdf`
- `Chua-phan-loai\Epigenetics_Childhood_Maltreatment_TranslPsych_2021.pdf`
- `Chua-phan-loai\fpubh-12-1232856.pdf`
- `GBD_WHO\GBD_2019_Mental_Disorders_Lancet_Psychiatry_2022.pdf`
- `The-gioi_Au-My-Uc\GWAS_122K_GABAergic_NatGenetics_2025.pdf`
- `The-gioi_Au-My-Uc\GWAS_Anxiety_NatHumBehav_2023.pdf`
- `The-gioi_Au-My-Uc\GWAS_MultiAncestry_NatGenetics_2024.pdf`
- `Coping_Effectiveness\Herres_Ohannessian_2015_Adolescent_Coping_Profiles_JAD.pdf`
- `Chua-phan-loai\HoangTrungHoc_2025_COVID_VN.pdf`
- `Dong-Nam-A\Indonesia_Adolescent_MH_2024.pdf`
- `The-gioi_Khac\Japan_Youth_Suicide_Hikikomori_2025.pdf`
- `Chua-phan-loai\journal.pone.0316825.pdf`
- `Chua-phan-loai\Korea_Adolescent_MH_Trends_2024.pdf`
- `Viet-Nam\_Archive_LoiFile\LongBinh_AnGiang_2024_DASS21.pdf`
- `Chua-phan-loai\Microbiota_GutBrain_Anxiety_2024_FrontNeuro.pdf`
- `Chua-phan-loai\Mindfulness_Nature_MentalHealth_2023.pdf`
- `Chua-phan-loai\Nakie_2022_Ethiopia_HighSchool_BMCPsych.pdf`
- `The-gioi_Au-My-Uc\Neural_Circuit_Pathological_Anxiety_NatRevNeuro_2024.pdf`
- `Chua-phan-loai\Neuroinflammation_Neuromodulation_TranslPsych_2022.pdf`
- `Viet-Nam\_Archive_LoiFile\NgoAnhVinh_2024_DTTS_LangSon.pdf`
- `Viet-Nam\_Archive_LoiFile\NguyenDanhLam_2024_YenDinh_ThanhHoa.pdf`
- `Viet-Nam\_Archive_LoiFile\NguyenNgocBaoQuyen_2025_HaNoi_YHCD.pdf`
- `Chua-phan-loai\tai-them-27052026\Pascoe_2020.pdf`
- `pham-et-al-2024-the-correlation-between-quality-of-care-and-mental-health-and-behavioral-problems-among-vietnamese.pdf`
- `Chua-phan-loai\Pham_2024_QualityCare_VN_Adolescents.pdf`
- `Chua-phan-loai\Pharmacotherapy_Anxiety_2020_Frontiers.pdf`
- `Chua-phan-loai\PIIS2468266725000982.pdf`
- `PIIS2772368225000034.pdf`
- `Chua-phan-loai\Qiu_2022_ParentingStyle_Resilience_FrontPsych.pdf`
- `S2054425125000391a.pdf`
- `s41182-025-00697-6.pdf`
- `Chua-phan-loai\SchoolFactors_VN_Anxiety_2021.pdf`
- `Chua-phan-loai\Social_Anxiety_Young_People_7Countries_2020_PLOSONE.pdf`
- `Chua-phan-loai\Stankov_2010_Confucian_Academic_ScienceDirect.pdf`
- `The-gioi_Au-My-Uc\Steare_2023_AcademicPressure_SR.pdf`
- `Coping_Effectiveness\Steinhoff_2023_Longitudinal_Coping_COVID_JEA.pdf`
- `Chua-phan-loai\Study on school-related factors impacting mental health and well-being of adolescents in Viet Nam.pdf`
- `Chua-phan-loai\TCNCYH_2025_LoAu_TramCam.pdf`
- `Chua-phan-loai\Viet-nam\TRANNGUYENNGOC-TamThan.pdf`
- `Chua-phan-loai\tai-them-27052026\V-NAMHS_2022.pdf`
- `Chua-phan-loai\VN_Hue_LoAu_TramCam_VTN_2025.pdf`
- `Chua-phan-loai\VN_Multicenter_2631HS_TPHCM_2025.pdf`
- `Chua-phan-loai\VN_PNT_DASS_THPT_2025.pdf`
- `Chua-phan-loai\VN_TPHCM_DASS_THPT_2024.pdf`
- `Chua-phan-loai\VNAMHS-Report_Eng_15-Feb-2023.pdf`
- `Chua-phan-loai\Wen_2020_LPA_Anxiety_Rural_China.pdf`
- `GBD_WHO\WHO_2022_World_Mental_Health_Report.pdf`

## Canonical entries trỏ tới PDF KHÔNG tồn tại (1)
- `QT002` → `QT002_Saikia_2023_India_Assam.pdf` (không tìm thấy)

## Cross-check 4 paper bài Q1 vừa thêm (07/06/2026)

| Paper | PDF | KG node |
|---|---|---|
| Karasu 1986 AJ Psychotherapy | ✓ có `karasu1986.pdf` | `PA_KARASU_1986_AJPSY` ✓ |
| Karasu Einstein bio sketch | ✓ có `biosketch_karasu_062112.pdf` | `AU_KARASU_TB (trong author node)` ✓ |
| Rose 2002 — Co-rumination | ⏳ cần download `(chưa download)` | `PA_ROSE_2002` ✓ |
| Stankov 2010 — Confucian | ⏳ cần download `(chưa download)` | `PA_STANKOV_2010` ✓ |
| Small & Blanc 2021 — Tam giao | ⏳ cần download `(chưa download)` | `PA_SMALL_BLANC_2021` ✓ |

## Hành động ưu tiên

1. **Thêm `karasu1986.pdf` và `biosketch_karasu_062112.pdf` vào `canonical_index.json`** — em sẽ làm trong vòng tiếp theo.
2. **Download 3 PDF còn thiếu** (Rose 2002, Stankov 2010, Small & Blanc 2021) — Rose + Stankov paywall, Small & Blanc mở qua PMC PMC7820702.
3. **61 PDF chưa categorize** trong 02_Papers-goc/ — chủ yếu là `Chua-phan-loai/` và `11-bai-ban-dau-va-mo-rong/`. Cần script categorize + metadata extraction.
4. **RAG rebuild** sau khi categorize: chạy `tro-ly-nghien-cuu-tam-ly/build_rag.py`.
