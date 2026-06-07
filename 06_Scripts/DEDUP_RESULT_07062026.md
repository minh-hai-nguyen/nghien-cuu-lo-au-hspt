# Dedup Result — 07/06/2026

**Target folders:** `03_Ban-dich/`, `Tom-tat-tung-bai/`
**Archive folders:** `_Archive/BanDich_dedup_07062026/`, `_Archive/TomTat_dedup_07062026/`

## Per-ID decisions

### 03_Ban-dich

| Keep | Archived |
| --- | --- |
| `DICH_Anderson_2025_FIXED_27052026.docx` | DICH_Anderson_2025.docx |
| `DICH_Hoa_2024_Frontiers_FIXED_27052026.docx` | DICH_Hoa_2024_Frontiers.docx |
| `DICH_VNAMHS_2022_FIXED_27052026.docx` | DICH_VNAMHS_2022.docx |
| `QT001_Jenkins_2023_USA_SanDiego_FIXED_27052026.docx` | QT001_Jenkins_2023_USA_SanDiego.docx |
| `QT010_Xu_2021_China_LargestEpi_FIXED_27052026.docx` | QT010_Xu_2021_China_LargestEpi.docx |
| `QT014_Anderson_2025_Wiley_Narrative_FIXED_27052026.docx` | QT014_Anderson_2025_Wiley_Narrative.docx |
| `QT021_Norway_Brunborg_2025_SocSciMed_FIXED_27052026.docx` | QT021_Norway_Brunborg_2025_SocSciMed.docx |
| `QT022_ScreenTime_Li_2025_BJCP_FIXED_27052026.docx` | QT022_ScreenTime_Li_2025_BJCP.docx |
| `QT026_UK_NHS_Parliament_2025_FIXED_27052026.docx` | QT026_UK_NHS_Parliament_2025.docx |
| `QT028_AJP_Zugman_PediatricAnxiety_2024_FIXED_27052026.docx` | QT028_AJP_Zugman_PediatricAnxiety_2024.docx |
| `QT035_Jefferies_SocialAnxiety_7Countries_2020_FIXED_27052026.docx` | QT035_Jefferies_SocialAnxiety_7Countries_2020.docx |
| `QT036_Moon_Korea_GAD_ML_2025_FIXED_27052026.docx` | QT036_Moon_Korea_GAD_ML_2025.docx |
| `QT043_Bress_JAMA_Maya_App_2024_FIXED_27052026.docx` | QT043_Bress_JAMA_Maya_App_2024.docx |
| `VN001_Hoa_2024_Frontiers_HaNoi_FIXED_27052026.docx` | VN001_Hoa_2024_Frontiers_HaNoi.docx |
| `VN002_VNAMHS_2022_National_FIXED_27052026.docx` | VN002_VNAMHS_2022_National.docx<br>VN002_VNAMHS_2022_National_v1_backup.docx<br>VN002_VNAMHS_2022_National_v1_backup_FIXED_27052026.docx |
| `VN017_NguyenDanhLam_2022_YHVN_FIXED_27052026.docx` | VN017_NguyenDanhLam_2022_YHVN.docx |
| `VN018_AnGiang_2025_YHVN_FIXED_27052026.docx` | VN018_AnGiang_2025_YHVN.docx |
| `VN023_NguyenLX_VN_COVID_Medicine_2023_FIXED_27052026.docx` | VN023_NguyenLX_VN_COVID_Medicine_2023.docx |

### Tom-tat-tung-bai

| Keep | Archived |
| --- | --- |
| `00_Mẫu tóm tắt bài 1_FIXED_27052026.docx` | 00_Mẫu tóm tắt bài 1.docx |
| `QT003_Mandaknalli_Malusare_2021_FIXED_27052026.docx` | QT003_Mandaknalli_Malusare_2021.docx |
| `QT005_Alharbi_et_al_2019_SaudiArabia_FIXED_27052026.docx` | QT005_Alharbi_et_al_2019_SaudiArabia.docx |
| `QT007_Chen_et_al_2023_China_BMCPsych_FIXED_27052026.docx` | QT007_Chen_et_al_2023_China_BMCPsych.docx |
| `QT010_Xu_et_al_2021_China_LargestEpi_FIXED_27052026.docx` | QT010_Xu_et_al_2021_China_LargestEpi.docx |
| `QT011_Bhardwaj_et_al_2020_India_Chandigarh.docx` | QT011_Bhardwaj_et_al_2020_India_Chandigarh_BEFORE_FIX44_01062026.docx |
| `QT035_Jefferies_SocialAnxiety_7Countries_2020_FIXED_27052026.docx` | QT035_Jefferies_SocialAnxiety_7Countries_2020.docx |
| `VN001_Hoa_2024_Frontiers_HaNoi_FIXED_27052026.docx` | VN001_Hoa_2024_Frontiers_HaNoi.docx |
| `VN002_VNAMHS_2022_National.docx` | VN002_VNAMHS_2022_National_v1_backup.docx |
| `VN017_NguyenDanhLam_2022_YHVN_FIXED_27052026.docx` | VN017_NguyenDanhLam_2022_YHVN.docx |
| `VN018_AnGiang_2025_YHVN_FIXED_27052026.docx` | VN018_AnGiang_2025_YHVN.docx |

## Totals

- IDs deduplicated in `03_Ban-dich/`: **18**
- Files archived from `03_Ban-dich/`: **20**
- IDs deduplicated in `Tom-tat-tung-bai/`: **11**
- Files archived from `Tom-tat-tung-bai/`: **11**
- **Total files archived:** 31

## Metadata strip

- OK: 29
- Locked (in use): 0
- Failed: 0

## Manual-review cases

### VN002 in `03_Ban-dich/` — `VN002_VNAMHS_2022_National_FULL.docx`

- Kept as canonical: `VN002_VNAMHS_2022_National_FIXED_27052026.docx` (566 KB, has 27052026 audit fixes).
- **NOT archived:** `VN002_VNAMHS_2022_National_FULL.docx` (1003 KB, ~2x larger than FIXED).
- Reason: FULL is significantly larger, may contain extended content not in FIXED. Needs manual content comparison before deciding whether to keep FULL alongside FIXED, replace FIXED with FULL, or merge.
