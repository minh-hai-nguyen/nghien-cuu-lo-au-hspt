# AUDIT — Sửa fabrication Saikia 2023 N=287 → N=360 trong outlines `bai-bao-Q1/`

**Ngày:** 07/06/2026
**Người thực hiện:** Claude (Opus 4.7), READ-ONLY audit + script generation
**Bối cảnh:** Audit `06_Scripts/AUDIT_Outlines_Q1Q3_07062026.md` (mục 3, mục 5.1) phát hiện outlines viết "Saikia 2023 N=287" trong khi PDF gốc (`02_Papers-goc/11-bai-ban-dau-va-mo-rong/11_Saikia_2023_IJCM.pdf`) khẳng định **N=360 (180 nam + 180 nữ, 10 trường × 36 học sinh, tuổi 14-17)**. "287" KHÔNG xuất hiện trong PDF gốc — đây là FABRICATION.

**Phạm vi:** 6 outline đang hoạt động trong `bai-bao-Q1/`. Bỏ qua 6 file `_BEFORE_GAD_DOI_*` (backup) và bỏ qua `bai-bao-khgdvn/` theo yêu cầu.

---

## 1. Outline chứa fabrication N=287 (Saikia)

Chỉ **1/6 outline** chứa pattern N=287:

| # | File | Vị trí (paragraph) | Pattern xuất hiện |
|---|---|---|---|
| 1 | `bai-bao-Q1/OutlineBilingual_Q1_01062026.docx` | paragraph[48] (VN) | "Saikia ... trên **287 thanh thiếu niên** báo cáo ... (Boy 30,0% vs Girl 18,9%, p = 0,049)" |
| 1 | (cùng file) | paragraph[49] (EN) | "Saikia ... in Northeast India **on 287 adolescents** reported ... (Boy 30.0% vs Girl 18.9%, p = 0.049)" |

5 outline còn lại đề cập Saikia 2023 NHƯNG **KHÔNG ghi sample size**, nên không cần sửa N:

- `BaiA_JAD_OUTLINE_v1_30052026.docx` — chỉ nhắc Saikia cho smartphone-addiction risk cluster + reference list.
- `BaiD_StressHealth_OUTLINE_v1_30052026.docx` — KHÔNG nhắc Saikia.
- `Outline_Q1_v3_01062026.docx` — chỉ liệt kê Saikia 2023 trong context list và reference, kèm số "Boy 30.0% vs Girl 18.9%" (đúng) nhưng KHÔNG có N.
- `Outline_Q3_v3_01062026.docx` — chỉ ghi "Saikia 2023 India (24.4% anxiety)" — số đúng, KHÔNG có N.
- `OutlineBilingual_Q3_01062026.docx` — chỉ liệt kê Saikia trong block reference, KHÔNG có N.

---

## 2. Side-by-side diff cho file cần sửa

### `OutlineBilingual_Q1_01062026.docx` — paragraph[48] (Tiếng Việt)

| OLD (BỊA) | NEW (đúng theo PDF gốc) |
|---|---|
| Kỳ vọng tiêu chuẩn dựa trên y văn quốc tế là tỷ lệ rối loạn lo âu ở nữ cao hơn nam (McLean và cs., 2011 – phân tích meta tổng hợp). Tuy nhiên, bằng chứng mới nổi thách thức giả định này: Wen và cộng sự (2020) trên học sinh nông thôn Trung Quốc ghi nhận nam > nữ; Saikia và cộng sự (2023) tại vùng Đông Bắc Ấn Độ **trên 287 thanh thiếu niên** báo cáo nam giới có tỷ lệ lo âu nặng cao hơn nữ giới (Boy 30,0% vs Girl 18,9%, p = 0,049). | Kỳ vọng tiêu chuẩn dựa trên y văn quốc tế là tỷ lệ rối loạn lo âu ở nữ cao hơn nam (McLean và cs., 2011 – phân tích meta tổng hợp). Tuy nhiên, bằng chứng mới nổi thách thức giả định này: Wen và cộng sự (2020) trên học sinh nông thôn Trung Quốc ghi nhận nam > nữ; Saikia và cộng sự (2023) tại vùng Đông Bắc Ấn Độ **trên 360 thanh thiếu niên (180 nam/180 nữ)** báo cáo nam giới có tỷ lệ lo âu nặng cao hơn nữ giới (Boy 30,0% vs Girl 18,9%, p = 0,049). |

### `OutlineBilingual_Q1_01062026.docx` — paragraph[49] (Tiếng Anh)

| OLD (BỊA) | NEW (đúng theo PDF gốc) |
|---|---|
| Standard expectation based on international literature is higher female-than-male anxiety disorder rates (McLean et al., 2011 – comprehensive meta-analysis). However, emerging evidence challenges this assumption: Wen et al. (2020) on rural Chinese students documented male > female; Saikia et al. (2023) in Northeast India **on 287 adolescents** reported higher severe anxiety rates in boys than girls (Boy 30.0% vs Girl 18.9%, p = 0.049). | Standard expectation based on international literature is higher female-than-male anxiety disorder rates (McLean et al., 2011 – comprehensive meta-analysis). However, emerging evidence challenges this assumption: Wen et al. (2020) on rural Chinese students documented male > female; Saikia et al. (2023) in Northeast India **on 360 adolescents (180 boys/180 girls)** reported higher severe anxiety rates in boys than girls (Boy 30.0% vs Girl 18.9%, p = 0.049). |

**Lưu ý:** tỷ lệ "Boy 30.0% / Girl 18.9% / p = 0.049" trong outline VẪN ĐÚNG theo Table 2 PDF gốc — chỉ sample size bị bịa. Sau khi sửa N=360, đoạn còn lại không cần đổi.

---

## 3. Quét N= khác để phát hiện fabrication tiềm ẩn

Quét toàn bộ pattern `N\s*=\s*[\d.,]+` trong 6 outline, đối chiếu với FACTCHECK đã có (`06_Scripts/FACTCHECK_*_07062026.md`).

| Citation | N trong outline | N canonical (PDF/FACTCHECK) | Trạng thái | Ghi chú |
|---|---|---|---|---|
| Xu et al. 2021 | N=373,216 | N=373,216 | OK | FACTCHECK_AsianStudies §1 — khớp |
| Chen et al. 2023 | N=63,205 | N=63,205 | OK | FACTCHECK_AsianStudies §2 — khớp |
| Wen et al. 2020 | (không ghi N trong outline) | N=900 | n/a | Outline không trích N — không có rủi ro fabrication N |
| **Saikia et al. 2023** | **N=287** (chỉ ở OutlineBilingual_Q1) | **N=360** | **FABRICATION** | FACTCHECK_AsianStudies §4 — đã quyết định sửa |
| Compas et al. 2017 | N=80,850 (5 outline) | N=80,850 (212 studies) | OK | FACTCHECK_Compas2017 §1.4 — khớp tuyệt đối |
| V-NAMHS 2022 | N=5,996 (BaiA) | (chưa có PDF V-NAMHS trong kho) | UNVERIFIED | Cần tải PDF report UNICEF/Bộ Y tế VN để xác minh; đã flag trong AUDIT_Q1Q3 §5.2 |
| Hoàng Trung Học & Nguyễn Thùy Dung 2025 | N=8,473 (4 outline) | (chưa có PDF trong kho) | UNVERIFIED | Flag trong AUDIT_Q1Q3 §5.4 |
| Bhardwaj et al. 2020 | N=288 (OutlineBilingual_Q1) | (chưa có PDF trong kho, chỉ có mention QT002 mislabel) | UNVERIFIED | Số 288 không trùng số bịa nào đã biết, nhưng cần verify PDF |
| Wang & Qin 2025 | N=500 (BaiA) | (chưa có PDF) | UNVERIFIED | Q1 example citation cần verify |
| Steare et al. 2023 | "52 studies, 48 confirming" | 52 studies, 48/52 positive | OK | FACTCHECK_Steare2023 §3 — khớp |
| Pascoe et al. 2020 | (paper type) | narrative review | OK nhưng cần chú thích | FACTCHECK_Pascoe2020 §3 — outline gọi đúng narrative review |
| Bài chính (Nhật Tân + Tây Mỗ) | N=1,352 | N=1,352 (614 nam + 738 nữ) | OK | Số tổng + breakdown khớp giữa 5 outline |

**Tổng số fabricated N values tìm thấy: 1** (Saikia N=287 — duy nhất 1 file, 2 vị trí).

**Tổng số N= UNVERIFIED (cần PDF gốc để fact-check): 4** (V-NAMHS, Hoàng Trung Học, Bhardwaj, Wang & Qin).

---

## 4. Khuyến nghị

### 4.1 Hành động ngay (READ-ONLY pass này KHÔNG thực hiện)

1. **Chạy script** `06_Scripts/fix_outlines_saikia_07062026.py --apply` để xuất `OutlineBilingual_Q1_01062026_FIXED_07062026.docx`. Script đã được dry-run thành công, đúng 2 vị trí (paragraph[48] VN + paragraph[49] EN). File gốc được GIỮ NGUYÊN.
2. Sau khi review file `_FIXED_07062026.docx`, anh có thể đổi tên thay thế file gốc.

### 4.2 Hành động kế tiếp (vượt scope pass này)

1. Tải PDF cho 4 nguồn UNVERIFIED (V-NAMHS 2022, Hoàng Trung Học 2025, Bhardwaj 2020, Wang & Qin 2025) → tạo FACTCHECK report mới tương tự `FACTCHECK_Compas2017_07062026.md`.
2. Sửa `02_Papers-goc/canonical_index.json` entry QT002 trỏ về `11_Saikia_2023_IJCM.pdf` (đã được flag trong AUDIT_Q1Q3 §4).
3. Review 5 outline còn lại để bổ sung N=360 cho Saikia khi outline được mở rộng thành draft — tránh thiếu sót khi viết.

---

## 5. Tham khảo (verification trail)

### PDF gốc đã đối chiếu
- `02_Papers-goc/11-bai-ban-dau-va-mo-rong/11_Saikia_2023_IJCM.pdf` — Saikia AM, Das H, Rajendran V (2023), Indian J Community Med 48(6):835–840. N=360 verified từ Methods (Sampling: "10 schools × 36 students = 360"; gender 180 nam / 180 nữ).

### FACTCHECK reports nội bộ
- `06_Scripts/FACTCHECK_AsianStudies_07062026.md` §4 (Saikia 2023 — N=360 verified)
- `06_Scripts/FACTCHECK_Compas2017_07062026.md` §1.4 (Compas 2017 — N=80,850 verified)
- `06_Scripts/FACTCHECK_Steare2023_07062026.md` §3 (Steare 2023 — 52 studies/48 positive verified)
- `06_Scripts/FACTCHECK_Pascoe2020_07062026.md` §1-§3 (Pascoe 2020 — narrative review verified)

### Audit trước đó
- `06_Scripts/AUDIT_Outlines_Q1Q3_07062026.md` — báo cáo gốc phát hiện vấn đề Saikia N (§3, §5.1)
- `06_Scripts/_saikia_287_scan.json` — output máy đọc của lần quét hôm nay (paragraph index + line)

### DOI/PMID verify cho Saikia
- DOI: 10.4103/ijcm.ijcm_201_23
- PMID: 38249689 (Indian J Community Med 2023 Nov-Dec; 48(6): 835-840)
- URL: https://journals.lww.com/ijcm/fulltext/2023/48060/mental_health_morbidities_and_their_correlates.10.aspx

### Script fix đã tạo
- `06_Scripts/fix_outlines_saikia_07062026.py` — review-before-run, mặc định dry-run, tạo file `_FIXED_07062026.docx`, strip metadata, KHÔNG ghi đè gốc.
