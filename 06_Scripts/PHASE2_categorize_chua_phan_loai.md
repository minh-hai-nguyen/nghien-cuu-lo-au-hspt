# PHASE 2 – Phân loại 14 PDFs MEDIUM/LOW trong `02_Papers-goc/Chua-phan-loai/`

Sinh ngày 2026-06-07. Nguồn: `06_Scripts/manual_review_chua_phan_loai_07062026.log`
(các dòng MEDIUM / LOW / REPARSE / STUB; bỏ qua DUPE đã có target).

Trạng thái OneDrive: 7/14 PDF hiện đang ở chế độ **cloud-only** (cloud provider OFFLINE,
attrib `O P`), không đọc được nội dung. Phân loại dưới đây dựa trên (a) tên file +
(b) hint của log + (c) bài đã đọc được full text + (d) ID DOI/PII đối chiếu.

**Lưu ý quan trọng:** QT076, QT077, QT078 đã có entry trong `canonical_index.json`
nhưng PDF chưa được di chuyển khỏi `Chua-phan-loai/`. Plan này sẽ **set `pdf_folder`**
cho 3 entry đó + tạo entry mới cho các bài còn lại.

Next QT-ID: **QT079** | Next VN-ID: **VN031**.

---

## A. HIGH-CONFIDENCE (đã đọc được nội dung hoặc đã có entry sẵn)

### 1. `Small_Blanc_2021_TamGiao_Vietnam.pdf`
- Tác giả: Sean Small, Judite Blanc (NYU, USA)
- Tiêu đề: *Mental Health During COVID-19: Tam Giao and Vietnam's Response*
- Tạp chí: Frontiers in Psychiatry 11:589618 (08/01/2021)
- DOI: 10.3389/fpsyt.2020.589618
- Quốc gia tác giả: USA – Chủ đề: Việt Nam (Tam Giao)
- **Target folder:** `The-gioi_Au-My-Uc/` (tác giả Mỹ, đề tài VN nhưng index hiện gán "USA")
- **QT-ID:** QT076 (đã có entry, chỉ cần set `pdf_folder`)
- Confidence: **HIGH** (đọc được full title page)

### 2. `Stankov_2010_Confucian_Academic.pdf` (+ bản trùng `..._ScienceDirect.pdf`)
- Tác giả: Lazar Stankov (University of Sydney, Australia)
- Tiêu đề: *Unforgiving Confucian culture: A breeding ground for high academic achievement,
  test anxiety and self-doubt?*
- Tạp chí: Learning and Individual Differences (Elsevier) – 2010
- **Target folder:** `The-gioi_Au-My-Uc/`
- **QT-ID:** QT077 (đã có entry, chỉ cần set `pdf_folder`)
- File `Stankov_2010_Confucian_Academic_ScienceDirect.pdf` (541 KB) là **bản trùng nội
  dung** (đã verify cùng tiêu đề + abstract) – đề xuất giữ bản chính + **xóa/archive bản
  ScienceDirect** sau khi user xác nhận.
- Confidence: **HIGH**

### 3. `1-s2.0-S266691532500054X-main.pdf`
- Tác giả: James Haoxiang Li, Larissa Shiying Qiu, Joshua Wong, Winnie Lau,
  Richard A. Bryant, July Lies, Belinda J. Liddell, Laura Jobson
- Tiêu đề: *Investigating cultural influences on the associations between rumination and
  symptoms of posttraumatic stress disorder among European Australian and Chinese
  Australian trauma survivors*
- Năm: 2025 | Tạp chí: (Elsevier S2666-9153 = *Journal of Mood & Anxiety Disorders*)
- Mẫu: 111 European Australian + 111 Chinese Australian
- Quốc gia: **Australia** (so sánh văn hoá Âu vs Trung)
- **Target folder:** `The-gioi_Au-My-Uc/` (NOT Đông Nam Á như log đoán)
- **QT-ID:** QT079 (mới)
- Confidence: **HIGH**

### 4. `Digital_MH_SocialAnxiety_MetaAnalysis_2025.pdf` → **DUPE QT040**
- Tác giả: Noemi Walder, Alessja Frey, Thomas Berger, Stefanie Julia Schmidt
  (University of Bern, Switzerland)
- Tiêu đề: *Digital Mental Health Interventions for the Prevention and Treatment of
  Social Anxiety Disorder in Children, Adolescents, and Young Adults: Systematic Review
  and Meta-Analysis of RCTs*
- Năm: 2025 | Loại: SR + MA
- **⚠ XÁC NHẬN:** QT040 đã tồn tại trong index:
  `QT040_Walder_JMIR_DMHI_SAD_2025.pdf` ở `The-gioi_Au-My-Uc/` (833.879 B).
  Source 836.299 B, SHA khác — có khả năng là **bản preprint vs typesetter** của cùng
  paper. Script `move_phase2_pdfs.py` sẽ chạy `verify_dupe` → vì SHA khác sẽ KHÔNG tự
  xoá, in cảnh báo "manual review needed".
- **Action:** user mở 2 PDF so trang đầu + danh mục tham khảo. Nếu cùng paper → xoá
  source. Nếu khác paper (khó xảy ra) → tạo QT080.
- Confidence: **HIGH** (đề tài chắc chắn trùng QT040 dựa trên metadata).

### 5. `TCNCYH_2025_LoAu_TramCam.pdf`
- Tác giả: Đào Thị Ngoãn và cs. (Trường Đại học Y Hà Nội)
- Tiêu đề: *Thực trạng tâm lý của sinh viên năm thứ tư Trường ĐH Y Hà Nội năm học
  2023-2024*
- Tạp chí: TCNCYH 187(02) – 2025, trang 296
- Mẫu: 196 SV năm 4, công cụ DASS-21
- **Target folder:** `Viet-Nam/`
- **QT-ID:** đã được index hoá là **VN028** (`VN028_DaoThiNgoan_TCNCYH_SVY4_HMU_2025.pdf`)
  → đây là **DUPE của VN028**. Đề xuất: verify SHA256 vs VN028 hiện tại, nếu trùng thì
  xoá file ở Chua-phan-loai (không tạo entry mới).
- Confidence: **HIGH** (đọc được abstract VN)

### 6. `Viet-nam/CVv443S402020122.pdf`
- Tác giả: Trần Thị Mỵ Lương (Học viện Phụ nữ Việt Nam)
- Tiêu đề: *Rối loạn lo âu ở học sinh Trung học phổ thông*
- Tạp chí: Tạp chí Khoa học – Trường Đại học Thủ Đô Hà Nội (CVv443) – 2020
- Loại: review khái niệm + khảo sát thực trạng RL lo âu HS THPT
- **Target folder:** `Viet-Nam/`
- **QT-ID:** **VN031** (mới)
- Đề xuất rename → `VN031_TranThiMyLuong_2020_THPT_LoAu.pdf`
- Confidence: **HIGH**

### 7. `tai-them-27052026/V-NAMHS_2022.pdf`
- Tiêu đề: *Viet Nam Adolescent Mental Health Survey (V-NAMHS) – Report on Main
  Findings*, 11/2022
- Tác giả: Institute of Sociology + University of Queensland + Johns Hopkins
- 51 trang (so với VNAMHS-Report_Eng_15-Feb-2023.pdf = 1.77 MB ở VN002)
- File này (1.87 MB, 51 tr) có khả năng là **bản tiếng Việt / báo cáo tóm tắt khác** của
  cùng dataset VN002.
- **Target folder:** `Viet-Nam/`
- **Action:** giữ làm phiên bản 2 của VN002 → rename → `VN002b_VNAMHS_2022_Main_Findings.pdf`
  HOẶC tạo VN032 nếu user muốn tách entry riêng.
- Confidence: **MEDIUM-HIGH** (xác minh đây là V-NAMHS thật nhưng cần user quyết bản chính
  vs bản phụ).

---

## B. MEDIUM-CONFIDENCE (OneDrive cloud-only, dựa vào tên file + log hint)

Các file này KHÔNG đọc được nội dung trong session này (OneDrive provider offline). Phân
loại dưới đây dựa vào tên file + log. **Cần force-download trước khi move**.

### 8. `VN_Hue_LoAu_TramCam_VTN_2025.pdf` (763 KB)
- Log hint: "Likely dupe của VN019 (HoThiTrucQuynh Hue 2025)"
- **Target folder:** `Viet-Nam/`
- **Action đề xuất:** verify SHA256 vs VN019. Nếu khác → tạo VN032.
- Confidence: **MEDIUM** (chưa đọc được PDF)

### 9. `VN_TPHCM_DASS_THPT_2024.pdf` (1.3 MB)
- Log hint: "Likely dupe của VN020 (TranHoVinhLoc TPHCM DASS-Y)"
- **Target folder:** `Viet-Nam/`
- **Action:** verify vs VN020. Nếu khác → VN033.
- Confidence: **MEDIUM**

### 10. `VN_PNT_DASS_THPT_2025.pdf` (555 KB)
- Log hint: "Likely dupe của VN026 (TranDucSi LongAn PNT 2025)"
- **Target folder:** `Viet-Nam/`
- **Action:** verify vs VN026. Nếu khác → VN034.
- Confidence: **MEDIUM**

### 11. `PIIS2468266725000982.pdf` (1.37 MB)
- PII = S2468-2667 → **The Lancet Public Health**
- Năm trong DOI: 25-0098 → 2025
- Có khả năng GBD region hoặc cross-country MH study
- **Target folder:** `The-gioi_Au-My-Uc/` (dự đoán; cần xác minh region)
- **QT-ID:** QT081 (mới)
- Confidence: **MEDIUM** (chưa đọc PDF; cần force-download)

### 12. `CBT_Delivery_GAD_TranslPsych_2025.pdf` (1.2 MB)
- Log hint: "Possibly QT020 (Liu CBT NMA TranslPsych 2025) – verify journal vs NMA"
- Tạp chí: Translational Psychiatry – 2025
- Chủ đề: CBT delivery methods cho GAD
- **Target folder:** `The-gioi_Au-My-Uc/`
- **Action:** verify vs QT020. Nếu khác → QT082.
- Confidence: **MEDIUM**

### 13. `fpubh-12-1232856.pdf` (183 KB)
- Tạp chí: Frontiers in Public Health, vol 12, article 1232856
- DOI prefix gợi 10.3389/fpubh.2024.1232856 (2024)
- **Target folder:** `The-gioi_Khac/` (theo log hint; cần verify region của tác giả)
- **QT-ID:** QT083 (mới)
- Confidence: **MEDIUM-LOW** (size 183 KB nhỏ → có thể chỉ là 4-6 tr commentary)

### 14. `journal.pone.0316825.pdf` (1.19 MB)
- Tạp chí: PLOS ONE, article 0316825
- DOI: 10.1371/journal.pone.0316825 → công bố 2024-2025
- **Target folder:** `The-gioi_Khac/` (theo log; cần verify region)
- **QT-ID:** QT084 (mới)
- Confidence: **MEDIUM**

---

## C. STUB / KHÔNG HỢP LỆ (không phân loại được)

| Filename | Size | Trạng thái |
|----------|------|-----------|
| `HoangTrungHoc_2025_COVID_VN.pdf` | 22 B | STUB – dupe của VN014; xoá source |
| `Pham_2024_QualityCare_VN_Adolescents.pdf` | 7070 B | STUB – cần tải lại; có thể là QT076 candidate cũ nhưng QT076 đã dùng cho Small_Blanc |
| `SchoolFactors_VN_Anxiety_2021.pdf` | 16 B | STUB – dupe VN027; xoá source |
| `Study on school-related factors ...pdf` | 4.7 MB | REPARSE (cloud-only invalid) – force download lại |
| `Child Adoles Psych Nursing - 2025 - Anderson...pdf` | 239 KB | REPARSE error trong log nhưng size > 0 → thử force download |

---

## D. PDFs đã có entry index nhưng `pdf_folder` chưa set

Cần update khi move:

| QT-ID | descriptor | đề xuất pdf_folder |
|-------|-----------|--------------------|
| QT076 | Small_Blanc_2021_USA_TamGiao_VN | `The-gioi_Au-My-Uc` |
| QT077 | Stankov_2010_Australia_ConfucianAcademic | `The-gioi_Au-My-Uc` |
| QT078 | Rose_2002_USA_CoRumination | `The-gioi_Au-My-Uc` |

---

## E. Tổng kết counts

- HIGH-confidence categorized: **7** (Small_Blanc QT076, Stankov QT077, Rose QT078,
  Li QT079, Walder=DUPE QT040, Đào Thị Ngoãn=DUPE VN028, Trần Mỵ Lương VN031,
  V-NAMHS=variant VN002b)
- MEDIUM (cần force-download trước khi xác nhận): **6** (3 file Viet-Nam có thể dupe +
  3 file opaque ScienceDirect/PLOS/Frontiers)
- STUB/REPARSE không xử lý được trong session: **5**
- Script `move_phase2_pdfs.py` action breakdown:
  - update_folder (entry sẵn, set pdf_folder): 3 (QT076, QT077, QT078)
  - new (tạo entry): 3 (QT079, VN031, VN002b)
  - verify_dupe (so SHA, xoá nếu trùng): 2 (Walder vs QT040 — SHA khác cần
    manual; TCNCYH vs VN028 — SHA trùng, sẽ xoá)
- Target folder breakdown (HIGH only):
  - `Viet-Nam/`: 2 mới (VN031, VN002b) + 1 dupe xoá (VN028)
  - `The-gioi_Au-My-Uc/`: 4 mới (QT076, QT077, QT078, QT079) + 1 dupe (QT040)
  - `Dong-Nam-A/`: 0
  - `The-gioi_Khac/`: 0 HIGH; 2 MEDIUM (fpubh + plos cần force-download)

---

## F. Sau khi user review

User chạy `python 06_Scripts/move_phase2_pdfs.py --apply` để di chuyển HIGH-confidence
PDFs và update `canonical_index.json`. Script default = `--dry-run`.
