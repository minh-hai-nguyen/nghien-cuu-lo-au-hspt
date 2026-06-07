# AUDIT — Outlines bai-bao-Q1 vs Paper gốc

**Ngày:** 07/06/2026
**Phạm vi:** 6 outline đang hoạt động trong `bai-bao-Q1/` (bỏ qua 6 file `_BEFORE_GAD_DOI_` backup).
**Nguyên tắc:** mọi claim/citation trong outline phải truy được PDF gốc trong `02_Papers-goc/` (canonical_index 117 entries).
**Phương pháp:** Trích citation bằng regex Author–Year (cả `(Author, YYYY)`, `Author et al. (YYYY)`, dạng ngoặc), đối khớp với `canonical_index.json`; spot-check 3 claim số liệu bằng cách đọc 3 trang đầu PDF gốc.

---

## 1. Tổng hợp citation per outline

| File | Words | Cites | Verified (có PDF) | Unverified | % PDF |
|------|------:|------:|------------------:|-----------:|------:|
| BaiA_JAD_OUTLINE_v1_30052026.docx | 3.201 | 11 | 2 | 9 | 18% |
| BaiD_StressHealth_OUTLINE_v1_30052026.docx | 2.089 | 11 | 1 | 10 | 9% |
| Outline_Q1_v3_01062026.docx | 1.877 | 5 | 0 | 5 | 0% |
| Outline_Q3_v3_01062026.docx | 2.078 | 7 | 1 | 6 | 14% |
| OutlineBilingual_Q1_01062026.docx | 8.792 | 39 | 11 | 28 | 28% |
| OutlineBilingual_Q3_01062026.docx | 5.698 | 16 | 6 | 10 | 38% |
| **TỔNG** | 23.735 | **89** | **21** | **68** | **24%** |

> Lưu ý: 24% nghe có vẻ thấp nhưng phải xét nature của citation (xem mục 2).

---

## 2. Phân loại citation "unverified"

Khi rà 68 unverified, chia thành 2 nhóm rất khác nhau:

### 2A. Citation phương pháp/thang đo gốc (KHÔNG cần PDF empirical)
Đây là các nguồn KINH ĐIỂN về công cụ/lý thuyết — không thuộc kho `02_Papers-goc/` (kho chỉ chứa paper thực nghiệm trẻ vị thành niên). Việc không có PDF là **chấp nhận được**:

- Rosenberg 1965 (RSES), Nunnally 1978 (α≥0.70), Folkman & Lazarus 1984 (Coping Theory), Carver 1997 (Brief COPE), Olweus 1996 (OBVQ), Goodenow 1993 (PSSM), Zimet 1988 (MSPSS), Chorpita 2000 (RCADS), Hu & Bentler 1999 (CFI/RMSEA cutoffs), Cheung & Rensvold 2002, Braun & Clarke 2006, Triandis 1995, Kitayama 1991, Helsinki 2013 (Declaration), Bronfenbrenner 2006, Skinner 2003/2007, Hayes 2020, Sun et al. 2011 (ESSA), Kwon et al. 2013 (SAS-SV), Allen 2013, Borkovec 2004, Roth 1986, Burton 2013, Morris 2006, Clark 2018, Ungar 2020, Kajastus 2024, Beesdo-Baum 2022.

→ Cần xác minh BIBLIOGRAPHIC (DOI/PMID) khi đưa vào reference list, nhưng KHÔNG bắt buộc PDF trong `02_Papers-goc/`.

### 2B. Citation empirical/dịch tễ có số liệu cụ thể (PHẢI có PDF — risk fabrication)
Đây là claim cần PDF gốc kiểm tra. Trong 68 unverified, chỉ **7 entry** rơi vào nhóm này:

| Citation | Outline xuất hiện | Lý do nghi |
|---|---|---|
| WHO 2022 (anxiety +25% COVID) | BaiA | Slogan WHO; cần ref report cụ thể |
| V-NAMHS 2022 (2.3% DSM-5) | BaiA, Q1, OutlineBilingual_Q1, OutlineBilingual_Q3 | KHÔNG có PDF V-NAMHS trong kho — claim 2.3% lặp 6 lần |
| Hoang Trung Hoc & Nguyen Thuy Dung 2025 (N=8,473 6 tỉnh) | BaiA, Q1, OutlineBilingual_Q1, OutlineBilingual_Q3 | KHÔNG có trong canonical_index; bị regex bắt thành "Hoc 2025" |
| Compas et al. 2017 meta (N=80,850, 212 studies) | BaiA, BaiD, Q1, OutlineBilingual_Q1 | Số liệu lặp nhiều lần, chưa thấy PDF |
| GBD ASEAN 2025 Lancet (10.1% VN, 11.9% region, 16.3% DALY 10-14) | Q1, Q3, OutlineBilingual_Q1, OutlineBilingual_Q3 | Claim 3 con số cụ thể; chưa có PDF GBD trong kho |
| McLean et al. 2011 (female > male anxiety norm) | Q1, Q3, OutlineBilingual_Q1, OutlineBilingual_Q3 | Norm citation — cần verify nhưng risk fabrication thấp |
| Pascoe 2020 + Steare 2023 (academic-stress reviews) | OutlineBilingual_Q1 | Pascoe **verified (QT067)**; Steare 2023 **unverified** — claim "48/52 studies" cần check |

---

## 3. Spot-check claim số liệu vs PDF gốc

| Claim trong outline | PDF gốc | Kết quả |
|---|---|---|
| Xu et al. (2021) JAD, **N=373,216** Chinese, **9.89%** anxiety GAD-7 | QT010 (`02_Papers-goc/11-bai-ban-dau-va-mo-rong/QT010_Xu_2021_China_LargestEpi.pdf`) | **KHỚP CHÍNH XÁC**. Abstract ghi rõ N=373,216, overall prevalence 9.89%. |
| Chen et al. (2023) BMC Psychiatry, **N=63,205**, **13.9%** anxiety | QT007 (`02_Papers-goc/The-gioi_Khac/QT007_Chen_et_al_2023_China_BMCPsych.pdf`) | **KHỚP CHÍNH XÁC**. Abstract: N=63,205, weighted anxiety prevalence 13.9% (95% CI 11.2-17.0). |
| Saikia et al. (2023) Northeast India, **N=287**, **Boy 30.0% vs Girl 18.9%, p=0.049** | `02_Papers-goc/11-bai-ban-dau-va-mo-rong/11_Saikia_2023_IJCM.pdf` | **PHẦN ĐÚNG, PHẦN SAI**. Boy 30.0% / Girl 18.9% / p=0.049 KHỚP CHÍNH XÁC (Table anxiety). NHƯNG sample size sai: paper thực tế **N=360** (180 boy + 180 girl), không phải 287. "287" có thể là sub-sample sau loại NA — outline cần làm rõ hoặc sửa. |

---

## 4. Vấn đề canonical_index phát hiện

- **QT002 mislabel**: `canonical_index.json` ghi `QT002_Saikia_2023_India_Assam.pdf` (file KHÔNG tồn tại). PDF Saikia thực tế nằm tại `11_Saikia_2023_IJCM.pdf`; có file `_MISLABELED_QT002_Saikia_actually_Bhardwaj_2020.pdf` cảnh báo. → Cần sửa `canonical_index.json` để outline chains không bị gãy.
- Một số entry BB02/BB10/BB08 trùng descriptor với QT — không phải lỗi nhưng tăng nguy cơ đếm trùng.

---

## 5. Đánh giá fabrication risk

**Không có claim nào hoàn toàn bịa** trong 3 mẫu spot-check (Xu/Chen khớp 100%; Saikia khớp về số liệu giới nhưng SAI N).

**Risk trung bình** xoay quanh:
1. Sample-size **287 vs 360** ở Saikia trong tất cả outline trích Saikia.
2. **V-NAMHS 2022: 2.3%** lặp 6 lần nhưng KHÔNG có PDF V-NAMHS trong kho → mọi outline đều dùng claim này chưa verify.
3. **GBD ASEAN 2025 Lancet** số chi tiết "10,1% VN / 11,9% khu vực / 16,3% DALY 10-14" lặp ở 4 outline mà chưa có PDF GBD — risk vừa.
4. **Hoàng Trung Học & Nguyễn Thùy Dung 2025 (N=8,473, 6 tỉnh, 41.5% covid / 25.4% post-COVID)** lặp 4 outline — chưa thấy PDF; có thể là LA chính nhưng cần explicit cite.
5. **Compas et al. 2017 meta (N=80,850 / 212 studies)** lặp 4 outline — chưa thấy PDF.

---

## 6. Khuyến nghị — Top 3 outline cần REWORK trước khi viết draft

1. **OutlineBilingual_Q1_01062026.docx** (lớn nhất, 8.792 từ, 39 cites): chỉ 28% có PDF, có ≥5 claim empirical chưa verify (V-NAMHS, GBD, Compas, Hoàng Trung Học, Saikia N). PHẢI tải PDF cho: V-NAMHS 2022 report, GBD ASEAN 2025 Lancet, Compas 2017 meta, Hoàng Trung Học 2025, Steare 2023, trước khi draft.
2. **Outline_Q1_v3_01062026.docx** (1.877 từ, 5 cites): 0% có PDF. Mọi citation đều unverified; cần bổ sung file PDF/bibliographic chính xác trước khi expand sang Bilingual.
3. **BaiA_JAD_OUTLINE_v1_30052026.docx** (3.201 từ, 11 cites): 18%, chứa WHO 2022 "+25%" claim và Hoàng Trung Học 2025 chưa verify. Vì target JAD Q1 nên risk reputational cao.

Outline **BaiD_StressHealth_OUTLINE_v1_30052026.docx** cũng thấp (9%) nhưng chủ yếu là theoretical/coping refs (Folkman, Lazarus, Carver, Skinner) → ít rủi ro fabrication; chấp nhận được nếu reference list có DOI.

---

## 7. Hành động cụ thể đề xuất

1. **Sửa `canonical_index.json` QT002** trỏ về file `11_Saikia_2023_IJCM.pdf` (chuẩn hóa tên thành `QT002_Saikia_2023_IJCM.pdf`).
2. **Tải bổ sung 5 PDF** vào `02_Papers-goc/`:
   - V-NAMHS 2022 final report (UNICEF/Bộ Y tế VN)
   - GBD 2021 ASEAN/Vietnam Lancet supplement
   - Compas et al. 2017 meta-analysis
   - Hoàng Trung Học & Nguyễn Thùy Dung 2025
   - Steare et al. 2023 academic-stress systematic review
3. **Sửa "N=287" Saikia → "N=360"** ở mọi outline trích Saikia 2023.
4. Tạo `canonical_index` entries cho cluster instrument refs (Rosenberg, MSPSS, Chorpita, Sun, Kwon, Olweus, Goodenow, Carver) với DOI để các outline citation Q1/Q3 không còn ở trạng thái "unverified" trên dashboard.

---

## Tham khảo nội bộ (truy vết)

- Source PDFs đã đọc: `02_Papers-goc/11-bai-ban-dau-va-mo-rong/QT010_Xu_2021_China_LargestEpi.pdf`, `02_Papers-goc/The-gioi_Khac/QT007_Chen_et_al_2023_China_BMCPsych.pdf`, `02_Papers-goc/11-bai-ban-dau-va-mo-rong/11_Saikia_2023_IJCM.pdf`.
- Canonical index: `02_Papers-goc/canonical_index.json` (117 entries, 104 author-year keys).
- Outlines audited (6): Glob `bai-bao-Q1/*Outline*.docx` + `bai-bao-Q1/*OUTLINE*.docx` (lọc bỏ `_BEFORE_GAD_DOI_`).
- Citation data: `06_Scripts/_audit_outlines_data.json` (machine-readable).
