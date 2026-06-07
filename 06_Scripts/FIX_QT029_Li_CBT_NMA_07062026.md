# FIX REPORT — QT029 (Li et al. 2025, BMC Psychiatry, NMA)

**Ngày sửa:** 07/06/2026
**File gốc (lỗi, giữ nguyên):** `Tom-tat-tung-bai/QT029_BMC_Li_CBT_NMA_2025.docx`
**File sửa:** `Tom-tat-tung-bai/QT029_Li_CBT_NMA_2025_FIXED_07062026.docx`
**PDF nguồn:** `02_Papers-goc/The-gioi_Khac/QT029_BMC_Li_CBT_NMA_2025.pdf`
**Phát hiện trong audit:** `06_Scripts/AUDIT_TomTat_FullFactCheck_MetaAnalysis_07062026.md`

---

## Lỗi factual đã phát hiện và sửa

### Lỗi #1 (HIGH) — Xếp hạng can thiệp HOÀN TOÀN SAI

**Tóm tắt cũ viết:** "*CBT cá nhân xếp hạng 1.* Network meta Bayesian xác nhận CBT cá nhân hiệu quả nhất trong 7+ loại can thiệp."

**PDF trang 7–8 (Relative effects, ranking of effectiveness and SUCRA):**
> "The ACT intervention was the most effective intervention for adolescents with anxiety disorders, with a MD of −3.83 [95% CrI: −9.33, 1.51]. CBT was the second most effective (MD = −3.64 [95% CrI: −7.36, −0.48]), followed by VRET (MD = −2.53 [95% CrI: −8.23, 3.32]) and PE (MD = −2.16 [95% CrI: −9.99, 5.52])."

**Đã sửa:** Hạng 1 = ACT (MD=−3,83; SUCRA=0,69); Hạng 2 = CBT (MD=−3,64; SUCRA=0,66); Hạng 3 = VRET (MD=−2,53; SUCRA=0,51); Hạng 4 = PE (MD=−2,16; SUCRA=0,51).

---

### Lỗi #2 (HIGH) — Hiểu sai khung so sánh

**Tóm tắt cũ:** Trình bày như thể bài so sánh các ĐỊNH DẠNG của CBT ("CBT cá nhân", "CBT nhóm", "iCBT").

**PDF trang 3 (Introduction, cuối) + trang 8 (Discussion mở đầu):**
> "this study aimed to compare and rank the effectiveness of ACT, CBT, PE, and VRET in treating anxiety disorders in this population using Bayesian network meta-analysis."

**Đã sửa:** Bài so sánh 4 LOẠI can thiệp (ACT, CBT, PE, VRET) — không phân tách CBT theo định dạng. Phân bố RCTs: CBT n=16, ACT n=6, VRET n=6, PE n=2 (PDF trang 5).

---

### Lỗi #3 (MEDIUM) — Sai cơ sở dữ liệu tìm kiếm

**Tóm tắt cũ:** "Tìm kiếm trong PubMed, PsychINFO, Web of Science đến 24/01/2025."

**PDF trang 3 (Search strategy):**
> "We conducted a comprehensive search of five major databases — Cochrane Library, Embase, PubMed, Scopus, and Web of Science — for studies published between January 1, 1976 and September 1, 2024."

**Đã sửa:** 5 CSDL (Cochrane, Embase, PubMed, Scopus, Web of Science), không có PsychINFO; thời điểm cutoff 01/09/2024 (không phải 24/01/2025).

---

### Lỗi #4 (MEDIUM) — Sai công cụ đánh giá thiên lệch

**Tóm tắt cũ:** "Đánh giá nguy cơ thiên lệch bằng Cochrane RoB 2.0."

**PDF trang 4 (Assessment of quality):**
> "The risk of bias was assessed for each study using Cochrane Review Manager 5.4, based on the methodological quality assessment criteria outlined in the Cochrane Handbook 5.1.0."

**Đã sửa:** Cochrane Review Manager 5.4 + Cochrane Handbook 5.1.0 (phiên bản CŨ — KHÔNG phải RoB 2.0).

---

### Lỗi #5 (LOW) — Tiêu chí RCT bị mô tả thiếu chính xác

**Tóm tắt cũ:** "RCT, trẻ em/VTN có rối loạn lo âu, so sánh ít nhất 2 can thiệp."

**PDF trang 3 (PICOS, Intervention):**
> "Included studies included one or more of the following interventions: cognitive behavioral therapy (CBT), acceptance and commitment therapy (ACT), physical exercise (PE), and virtual reality exposure therapy (VRET)."

**Đã sửa:** Tiêu chí can thiệp là "một hoặc nhiều trong 4 loại"; đối chứng có thể là chăm sóc thông thường/không can thiệp/danh sách chờ.

---

### Lỗi #6 (LOW) — Thiếu thông tin tác giả và xuất bản

**Tóm tắt cũ:** Không liệt kê tác giả, không có DOI, không có volume.

**PDF trang 1:**
> "Longhui Li, Qiner Li, Jingyi Wang, Quan Fu and Meng Chi"
> "Li *et al. BMC Psychiatry* (2025) 25:809"
> "https://doi.org/10.1186/s12888-025-07227-y"

**Đã sửa:** Bổ sung đầy đủ 5 tác giả, affiliations, BMC Psychiatry 2025;25:809, DOI 10.1186/s12888-025-07227-y, PROSPERO CRD42024587910.

---

### Lỗi #7 (LOW) — Đối chiếu liên bài bịa đặt

**Tóm tắt cũ:** "Ba tổng quan can thiệp (AJP 2024, BMC 2025, Zhameden 2025) đều xác nhận CBT hiệu quả nhất."

**Vấn đề:** Suy diễn sai dựa trên giả định rằng bài đang xét xếp CBT hạng 1 (thực tế là ACT). "Zhameden 2025" không xuất hiện trong PDF QT029 và không xác minh được như nguồn tham khảo trong bài này.

**Đã sửa:** Loại bỏ đoạn đối chiếu liên bài chưa kiểm chứng; thay bằng phân tích nội tại từ PDF (so sánh ACT vs CBT về SUCRA, CrI, số RCT).

---

## Phát hiện bổ sung (KHÔNG sửa — chỉ ghi nhận)

- **Số hình 6 (đúng):** Fig 1 PRISMA, Fig 2 RoB, Fig 3 network, Fig 4 relative effects, Fig 5 SUCRA, Fig 6 funnel — đếm trực tiếp từ PDF trang 5–10.
- **Số bảng:** Tóm tắt cũ ghi "3 bảng". PDF chỉ có Table 1 trong văn bản chính (đặc điểm 30 RCTs, trang 6). Supplementary Table 2 và Table 3 nằm trong phụ lục, không phải trong văn bản chính. ĐÃ SỬA thành "1 bảng".
- **30 RCTs, 1.711 trẻ — ĐÚNG** (PDF trang 1 abstract + trang 4 Result).
- **Tuổi 6–18 — ĐÚNG** (PDF trang 3 PICOS).
- **τ = 5,87, p Egger = 0,91, kappa = 0,83 — ĐÚNG** (PDF trang 6, 8).

---

## Mục bất định (đã đánh dấu trong tóm tắt)

Không có mục TBD. Tất cả số liệu và can thiệp trong bản FIXED đều có trích dẫn trực tiếp từ PDF trang 1–10.

---

## Định dạng

- Times New Roman 12pt, line spacing 1,5
- Metadata: author, title, subject, keywords, comments, category, last_modified_by — đều rỗng
- Footer: "Đã sửa lỗi factual (07/06/2026) — đối chiếu trực tiếp PDF gốc"

---

**Trạng thái:** Hoàn thành. File gốc lỗi vẫn được GIỮ NGUYÊN để truy vết audit.
