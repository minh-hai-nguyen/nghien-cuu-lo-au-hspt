# AUDIT — 03_Ban-dich (07/06/2026)

Read-only audit. Scope: all `.docx` + `.txt` + `.md` under `03_Ban-dich/` (incl. subdirs `Dich-11-bai-dau-tien/`, `Bai_dich_phan_bien/`). Reference: `02_Papers-goc/canonical_index.json` (117 IDs: 78 QT + 24 VN + 15 BB).

## 1. File inventory

| Loại | Số lượng |
|---|---|
| `.docx` | 105 |
| `.txt` | 19 |
| `.md` | 21 |
| **Tổng** | **145** |

Nhóm theo prefix (gộp docx+txt+md):

- `QT###_` : 57 file (gồm bản gốc + bản `_FIXED_27052026` + 1 raw `QT036_...raw.txt`)
- `VN###_` : 30 file (incl. `_FIXED`, `_FULL`, `_v1_backup`, 1 raw `VN021_..._raw.txt`)
- `BB##_` : 0 file (không có file nào dùng prefix BB)
- `A##_/B##_` raw drafts : 17 file (toàn `.txt`/`.docx` cũ)
- `DICH_*` (tên cũ trước khi chuẩn hoá QT/VN) : 16 file
- `TomTat_*` : 3 file
- Khác (hướng dẫn, raw lẻ, dich_phan_bien) : 22 file

## 2. Coverage vs canonical_index

- **Đã có bản dịch: 66 / 117 (56,4%)**
- Phân theo prefix:
  - QT: 46 / 78 có (59%)
  - VN: 20 / 24 có (83%)
  - BB: 0 / 15 có (0%) — toàn bộ 15 sách/báo cáo BB chưa có bản dịch nào
- **17 ID có nhiều bản (versioning):** QT001, QT010, QT014, QT021, QT022, QT026, QT028, QT035, QT036, QT043, VN001, VN002, VN017, VN018, VN021, VN022, VN023 — đa phần là cặp `bản gốc + _FIXED_27052026`. VN002 đặc biệt nhiều: 4 bản (`_National`, `_FIXED`, `_FULL`, `_v1_backup`, `_v1_backup_FIXED`).

### Thiếu (51 ID)

- **QT (32):** QT003-007, QT009, QT011, QT018-020, QT052-061, QT063, QT068-078
- **VN (4):** VN004, VN005, VN006, VN007
- **BB (15):** toàn bộ BB01-03, BB05-12, BB14, BB18, BB19, BB22

## 3. Raw drafts vs polished

Tổng 19 file `_raw.txt`. Đối chiếu nội dung với docx:

- **8 raw đã có docx cùng nhân vật/chủ đề** → có thể là nháp ban đầu, đã hoàn thiện và rename theo QT/VN:
  - `A12_Dinh_SchoolFactors_raw.txt` → `VN027_Dinh_SchoolFactors_VN_2021.docx`
  - `A5_Dong_raw.txt` → `QT047_Dong_PLOS_DASS_YaAn_2025.docx`
  - `A6_Zheng_MXH_raw.txt` → `QT041_Zheng_China_MXH_PRBM_2025.docx`
  - `A7_AcademicStress_raw.txt` → `QT046_Jagiello...` hoặc `QT067_Pascoe...`
  - `A8_VN_COVID_raw.txt` → `VN014` hoặc `VN023` (cần xác định)
  - `A9_VinhLong_raw.txt` → `VN024_NguyenThanhTruyen_VinhLong_YHVN_2024.docx`
  - `A10_HaiPhong_raw.txt` → `VN025_PhamThiNgoc_HaiPhong_VinhBao_2024.docx`
  - `A11_LongAn_raw.txt` → `VN026_TranDucSi_LongAn_PNT_2025.docx`
  - `B5_UK_school_raw.txt` → `QT042_BrownCarter_UK_School_JMH_2025.docx`
  - `B6_resilience_raw.txt` → `QT044_Cai_Resilience_Frontiers_2025.docx`
  - `B7_CACBT_SEA_raw.txt` → `QT037_Praptomojati_CA-CBT_SEA_AJP_2024.docx`
  - `B8_SriLanka_CBT_raw.txt` → `QT038_DeSilva_SriLanka_RCT_2024.docx`
  - `B11_Japan_iCBT_raw.txt` → `QT045_Sasaki_Japan_iCBT_JMIR_2024.docx`
  - `B2_JMIR_Digital_SAD_raw.txt` → `QT040_Walder_JMIR_DMHI_SAD_2025.docx`
  - `B3_JAMA_App_raw.txt` → `QT043_Bress_JAMA_Maya_App_2024.docx`
  - `B9_NMA_SAD_raw.txt` → `QT039_Xian_NMA_SAD_JAD_2024.docx`
  - `QT036_..._raw.txt` → `QT036_Moon_Korea_GAD_ML_2025.docx`
  - `VN021_TranThaoVi_raw.txt` → `VN021_TranThaoVi_2024_JRuralMed_Hue.docx`
- **Raw thật sự WIP (không có docx tương ứng): 1 file** — `UNICEF_VN_raw.txt` (đã có `VN022_UNICEF_VN_2022_SchoolFactors.docx` & `_FULL.docx` → có thể coi là đã hoàn thiện, vậy thực chất **0 WIP bị bỏ dở**).

## 4. Orphan translations

- Không có file nào dùng QT/VN/BB ID nằm ngoài canonical_index. Các file `DICH_*.docx/md` (Anderson, GBD_ASEAN, Hoa_2024, VNAMHS, Zhameden) là **bản cũ trước chuẩn hoá** — tất cả đã có bản chuẩn QT/VN tương ứng (QT014, QT012, VN001, VN002, QT013).

## 5. Quality spot-check (5 file ngẫu nhiên, seed=7)

| File | Words | Ký tự VN | Trích nguồn | Phản biện/Hạn chế | TODO/placeholder |
|---|---:|---:|:---:|:---:|:---:|
| `QT036_Moon_Korea_GAD_ML_2025.docx` | 2 261 | 729 | có (DOI) | Hạn chế | không |
| `QT016_Mudunna_2025_LancetSEA_SouthAsia.docx` | 1 432 | 517 | có | Hạn chế + Đánh giá | không |
| `QT043_Bress_JAMA_Maya_App_2024_FIXED_27052026.docx` | 3 007 | 1 063 | có | Hạn chế + Đánh giá | không |
| `VN020_TranHoVinhLoc_2024_TPHCM_DASS-Y.docx` | **696** | 238 | có | Hạn chế | không |
| `DICH_VNAMHS_2022_FIXED_27052026.docx` | 27 652 | 10 986 | có | Limitation + Đánh giá | không |

Nhận xét: 5/5 có tiếng Việt thực + trích nguồn + dấu hiệu nhận xét. **VN020 chỉ 696 từ → nghi là bản tóm tắt chứ không phải dịch đầy đủ; cần kiểm tra**. Không thấy chuỗi TODO/PLACEHOLDER trong mẫu.

## 6. Khuyến nghị

1. **Ưu tiên dịch BB (15 ID):** toàn bộ sách/báo cáo BB01-22 chưa có bản dịch nào → mảng khung lý thuyết đang trắng.
2. **Dọn version trùng:** 17 ID có ≥2 bản; chọn 1 bản canonical (thường là `_FIXED_27052026`), archive bản cũ vào `_Archive/`. VN002 cần dọn gấp (5 bản).
3. **Lấp khoảng QT003-011 + QT052-078 + VN004-007:** 36 ID còn thiếu, đa phần thuộc batch mở rộng — cần lên kế hoạch dịch theo cụm chủ đề.
4. **Kiểm VN020** (chỉ 696 từ) xem có phải bản nháp/tóm tắt nhầm tên không.
5. **Xoá 19 file `_raw.txt`** sau khi xác nhận bản docx đã ổn — đang gây nhiễu inventory.

## Phụ lục — file đếm chi tiết

- Coverage: 66 / 117 = 56,4%
- Multi-version: 17 ID
- Raw có docx hoàn thiện: 18/19
- Raw thực sự bỏ dở: 0
- Orphan: 0
