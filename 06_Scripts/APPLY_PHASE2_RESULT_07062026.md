# PHASE 2 APPLY RESULT — 07/06/2026

**Script:** `06_Scripts/move_phase2_pdfs.py --apply`
**Working dir:** `c:\Users\OS\OneDrive\read_books\Lo-au\`
**Backup tao truoc apply:**
- `02_Papers-goc/canonical_index.json.bak_phase2_apply_07062026` (manual, pre-run)
- `02_Papers-goc/canonical_index.json.bak_phase2_20260607_135538` (auto, by script)

---

## 1. Stats tu script

| Metric | Count |
|--------|-------|
| moved | 6 |
| skipped | 0 |
| dupe_deleted | 1 |
| entries_added | 2 (via script) + 1 (manual QT083) = **3** |
| entries_updated | 3 |
| errors / warnings | 2 |

---

## 2. Files moved (source -> target)

| # | QT | Source (Chua-phan-loai/) | Target |
|---|----|--------------------------|--------|
| 01 | QT076 | `Small_Blanc_2021_TamGiao_Vietnam.pdf` | `The-gioi_Au-My-Uc/QT076_Small_Blanc_2021_USA_TamGiao_VN.pdf` |
| 02 | QT077 | `Stankov_2010_Confucian_Academic.pdf` | `The-gioi_Au-My-Uc/QT077_Stankov_2010_Australia_ConfucianAcademic.pdf` |
| 03 | QT078 | `Rose_2002_CoRumination.pdf` | `The-gioi_Au-My-Uc/QT078_Rose_2002_USA_CoRumination.pdf` |
| 04 | QT083* | `1-s2.0-S266691532500054X-main.pdf` | `The-gioi_Au-My-Uc/QT083_Li_2025_Australia_Rumination_PTSD_Chinese_EuroAus.pdf` |
| 05 | VN031 | `Viet-nam/CVv443S402020122.pdf` | `Viet-Nam/VN031_TranThiMyLuong_2020_THPT_LoAu.pdf` |
| 06 | VN002b | `tai-them-27052026/V-NAMHS_2022.pdf` | `Viet-Nam/VN002b_VNAMHS_2022_Main_Findings.pdf` |

*QT083 conflict note: script ban dau gan ID = QT079, nhung QT079 da bi Kieling_2024_JAMA da chiem trong index. File da duoc rename sang QT083 va entry moi them voi ghi chu trong field `note`.

## 3. Dupe deleted

| QT | Source deleted | Reason |
|----|----------------|--------|
| VN028 | `Chua-phan-loai/TCNCYH_2025_LoAu_TramCam.pdf` | SHA256 khop 100% voi `Viet-Nam/VN028_DaoThiNgoan_TCNCYH_SVY4_HMU_2025.pdf` |

## 4. Warnings / NOT deleted (manual review)

- **QT040 `Digital_MH_SocialAnxiety_MetaAnalysis_2025.pdf`**: SHA khac (836,299 B vs 833,879 B) so voi `QT040_Walder_JMIR_DMHI_SAD_2025.pdf`. KHONG xoa — can manual review xem co phai typesetter version khac hay paper khac.
- **QT079 ID collision**: handled bang cach reassign sang QT083 (xem muc 2).

---

## 5. Canonical entries moi them

| QT | Descriptor | pdf_folder |
|----|------------|-----------|
| VN031 | TranThiMyLuong_2020_THPT_LoAu | Viet-Nam |
| VN002b | VNAMHS_2022_Main_Findings_Variant | Viet-Nam |
| QT083 | Li_2025_Australia_Rumination_PTSD | The-gioi_Au-My-Uc |

## 6. Canonical entries updated (pdf_folder + pdf)

| QT | folder before | folder after | pdf after |
|----|---------------|--------------|-----------|
| QT076 | None | The-gioi_Au-My-Uc | QT076_Small_Blanc_2021_USA_TamGiao_VN.pdf |
| QT077 | None | The-gioi_Au-My-Uc | QT077_Stankov_2010_Australia_ConfucianAcademic.pdf |
| QT078 | None | The-gioi_Au-My-Uc | QT078_Rose_2002_USA_CoRumination.pdf |

---

## 7. Verification canonical_index.json

| Check | Result |
|-------|--------|
| Total entries before apply | 121 |
| Total entries after apply | **124** (+3 = QT083 + VN031 + VN002b) |
| Duplicate IDs | **None** |
| Entries with valid pdf + pdf_folder pointing to existing file | 122/124 |
| Pre-existing missing PDF entries (not caused by apply) | 2: QT072, QT073 (files exist at `02_Papers-goc/` root, pdf_folder field stuck at value `02_Papers-goc`) |

## 8. Audit re-run (sau apply)

```
06_Scripts/AUDIT_PapersGoc_07062026.md
  161 PDFs total
  123 canonical entries     <-- pre-QT083 add
  59 not in canonical
  0 canonical missing PDF
```
(QT083 duoc them sau khi audit chay -> hien tai 124 entries.)

---

## 9. Files con lai trong Chua-phan-loai/ (sau apply)

**31 PDF + 2 subdir PDF + cac .docx/.md/.txt** chua duoc phan loai. Nhung file nay nam ngoai PHASE 2 plan (MEDIUM/LOW confidence) va se can phase rieng. Subset:

- Bhardwaj_2020_Chandigarh_DASS21.pdf
- CBT_Delivery_GAD_TranslPsych_2025.pdf
- Chen_2023_Chinese_SecondarySchool_BMCPsych.pdf
- Child Adoles Psych Nursing - 2025 - Anderson - ...
- Chronic_Stress_Neuroinflammation_FrontPsych_2023.pdf
- Digital_MH_SocialAnxiety_MetaAnalysis_2025.pdf (giu lai — chua match SHA QT040)
- Epigenetics_Childhood_Maltreatment_TranslPsych_2021.pdf
- HoangTrungHoc_2025_COVID_VN.pdf
- Korea_Adolescent_MH_Trends_2024.pdf
- Microbiota_GutBrain_Anxiety_2024_FrontNeuro.pdf
- Mindfulness_Nature_MentalHealth_2023.pdf
- Nakie_2022_Ethiopia_HighSchool_BMCPsych.pdf
- Neuroinflammation_Neuromodulation_TranslPsych_2022.pdf
- NgoAnhVinh_2024_DTTS_LangSon.pdf
- Niwenahisemo_2024_GAD7_Rwanda.pdf (entry QT080 trong index nhung folder=None — can move sau)
- PIIS2468266725000982.pdf
- Pham_2024_QualityCare_VN_Adolescents.pdf
- Pharmacotherapy_Anxiety_2020_Frontiers.pdf
- Qiu_2022_ParentingStyle_Resilience_FrontPsych.pdf
- SchoolFactors_VN_Anxiety_2021.pdf
- Social_Anxiety_Young_People_7Countries_2020_PLOSONE.pdf
- Stankov_2010_Confucian_Academic_ScienceDirect.pdf (variant cua QT077 — verify)
- Study on school-related factors impacting mental health and well-being of adolescents in Viet Nam.pdf
- VNAMHS-Report_Eng_15-Feb-2023.pdf
- VN_Hue_LoAu_TramCam_VTN_2025.pdf
- VN_Multicenter_2631HS_TPHCM_2025.pdf
- VN_PNT_DASS_THPT_2025.pdf
- VN_TPHCM_DASS_THPT_2024.pdf
- Wen_2020_LPA_Anxiety_Rural_China.pdf
- fpubh-12-1232856.pdf
- journal.pone.0316825.pdf
- Viet-nam/16.-NguyenThiVan.pdf
- Viet-nam/TRANNGUYENNGOC-TamThan.pdf
- tai-them-27052026/Pascoe_2020.pdf

---

## 10. Action items tiep theo

1. **Manual review QT040 conflict**: kiem tra `Digital_MH_SocialAnxiety_MetaAnalysis_2025.pdf` vs `QT040_Walder_JMIR_DMHI_SAD_2025.pdf` — neu cung paper thi xoa, neu khac thi tao QT moi.
2. **Pre-existing index bug**: QT072/QT073 co `pdf_folder = "02_Papers-goc"` (sai). Can sua pdf_folder thanh dung sub-region.
3. **PHASE 3**: phan loai 31 PDF MEDIUM/LOW con lai (gom nhieu paper toan cau ve neuro/CBT, VN regional studies, va variant cua entries da co).
4. **QT080 (Niwenahisemo Rwanda)**: PDF dang nam tai Chua-phan-loai/, can move vao `The-gioi_Khac/`.
