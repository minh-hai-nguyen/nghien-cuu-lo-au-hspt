# AUDIT — Fact-check 10 tóm tắt vs PDF gốc (Tom-tat-tung-bai)

**Ngày:** 07/06/2026
**Phạm vi:** Spot-check 10 bài (8 QT + 2 VN) ở `Tom-tat-tung-bai/` so với `02_Papers-goc/`
**Phương pháp:** Đọc summary `.docx` + 3 trang đầu PDF gốc. Kiểm 5 yếu tố: N, năm, tác giả đầu, quốc gia/địa bàn, prev/effect-size cốt lõi.
**Severity:** HIGH = fabrication/sai số liệu; MEDIUM = wrong attribution / nhầm efficacy↔acceptability; LOW = formatting / phiên âm tên / thiếu minor.

---

## Bảng tổng hợp

| # | ID | Trạng thái | Tóm phát hiện chính |
|---|---|---|---|
| 1 | QT001 Jenkins 2023 USA | ✓ OK | N=75 ✓, năm ✓, tác giả ✓, San Diego ✓, prev PHQ 44% / GAD 50.6% khớp PDF (49.3% "none") |
| 2 | QT010 Xu 2021 China | ✓ OK | N=373,216 ✓, JAD 288:17-22 ✓, Henan (Zhengzhou+Xinxiang+Xinyang) ✓, prev 9.89% ✓, nam>nữ 10.11/9.66% ✓ |
| 3 | QT020 Liu CBT NMA 2025 | ⚠ MINOR + 1 MEDIUM | 52 RCT N=4361 ✓, Transl Psychiatry ✓. **MEDIUM:** Summary ghi "Individual CBT > Remote CBT: SMD 0,96" như EFFICACY — nhưng PDF: SMD 0.96 là **ACCEPTABILITY** (Individual *acceptable hơn* Remote), không phải efficacy. Nhầm 2 khái niệm. **MINOR:** tuổi TB ghi 43, PDF body nói 40.8 (abstract: 43); nữ 69.7% (abstract) vs 67.1% (body) — summary copy từ abstract. |
| 4 | QT030 GBD Trends 2025 | ✓ OK | Zhang Dongjun + Wu Mingyue ✓, JAD 387:119491 ✓, Xinxiang Medical Univ ✓ (summary phiên âm "Tân Hương" — chấp nhận), GBD 2021, 204 nước, AAPC 0.84-0.97% ✓, Joinpoint+APC ✓ |
| 5 | QT040 Walder DMHI SAD 2025 | ⚠ LOW | 21 RCT (k=21 MA, k=22 review) ✓, g=0.508 ✓, SAD-specific g=0.878 ✓, guided g=0.825 ✓, PROSPERO ✓. **LOW:** Summary ghi 2025 — PDF là preprint submitted Oct 2024 (chưa published chính thức). |
| 6 | QT052 Mindfulness Nature MH 2023 | ✓ OK (skim) | Summary là "bonus note" rất ngắn, không có claim số liệu cụ thể để verify. PDF khớp: IPD MA 13 trials N=2,371, 8 nước, Nature Mental Health 2023. |
| 7 | QT060 Bie GBD 2024 | ⚠ HIGH | Bie Fengsai ✓, Front Psychiatry 2024 ✓, +52% ✓, 708→883/100k ✓, India highest # cases ✓, 20-24y +28.33% ✓. **HIGH:** Summary khẳng định "Latin Mỹ Tropical cao nhất; Đông Á giảm nhiều nhất" — abstract PDF KHÔNG nêu Latin Tropical highest, mà nói "middle SDI highest incidence/prevalence, high SDI largest increases". "Đông Á giảm" cũng KHÔNG xuất hiện trong abstract — đây là khu vực **chậm tăng nhất**, không phải giảm. Sai vùng + sai chiều xu hướng. |
| 8 | QT072 Lee Cyberbullying 2025 | ⚠ HIGH ×2 | Jungup Lee ✓, TVA 27(2):391-406 ✓, năm 2025 ✓, 27 longitudinal studies ✓. **HIGH-1:** Summary ghi **"N = 27.133"** — PDF rõ ràng "27 studies encompassing **13,497** children and adolescents". Số 27,133 dường như do nhầm "27 studies × ... " → bịa N. **HIGH-2:** Summary ghi 6 authors **thiếu 2 người**: bỏ Hyekyung Choo và Yijing Zhang (PDF: Lee, Choo, Zhang Y, Cheung, Zhang Q, Ang). |
| 9 | VN001 Hoa 2024 Hanoi | ✓ OK | Pham Thi Thu Hoa + 3 đồng tác giả ✓, Front Public Health 12:1232856 ✓, DOI ✓, N=3,910 ở 13 trường THPT Hà Nội ✓, GAD-7 α=0.916 ✓, prev 40.6% (mild 23.9 / mod 10.9 / severe 5.8) ✓ |
| 10 | VN002 V-NAMHS 2022 | ✓ OK | Tất cả 14 tác giả (4 institutions) khớp chính xác trang bìa PDF; tháng 11/2022 ✓; DISC-5 trên 10-17y ✓; N=5,996 ✓ |

---

## Tổng số lỗi theo severity

| Severity | Số bài | Bài cụ thể |
|---|---|---|
| ✗ HIGH (fabrication / sai số liệu / sai chiều) | **2** | QT060 (sai vùng + sai chiều xu hướng); QT072 (N bịa 27.133 vs thực 13.497 + thiếu 2 tác giả) |
| ⚠ MEDIUM (wrong attribution / nhầm khái niệm) | **1** | QT020 (nhầm efficacy ↔ acceptability cho SMD 0.96) |
| ⚠ LOW (formatting / nguồn preprint vs final / phiên âm) | **2** | QT020 (số liệu copy từ abstract chứ không phải body); QT040 (preprint 2024 ghi 2025) |
| ✓ OK | **5** | QT001, QT010, QT030, QT052, VN001, VN002 (= 6 thực tế OK) |

**Tỷ lệ lỗi đáng kể (HIGH+MEDIUM): 3/10 = 30%.**

---

## Mẫu lỗi (pattern)

1. **Confabulation số liệu trong bài tổng quan/meta-analysis lớn** — QT072 N=27.133 và QT060 "Latin Tropical / Đông Á giảm" đều là loại lỗi "fluent fabrication": câu nghe trôi chảy, có số cụ thể, nhưng KHÔNG có trong PDF. Cả 2 đều ở bài *meta-analysis / GBD* — nơi có nhiều con số dễ nhầm.
2. **Nhầm 2 outcome song song (efficacy vs acceptability)** — QT020 cho SMD 0.96 sai chiều ý nghĩa.
3. **Copy từ abstract khi body khác** — QT020 (tuổi 43 vs 40.8; nữ 69.7 vs 67.1) — hay gặp ở bài Q1 có nhiều khác biệt abstract-body do version.
4. **Tác giả thiếu/cắt ngắn** — QT072 thiếu 2/6, QT040 cắt "et al." — không hệ thống nhưng có rủi ro tích lũy.
5. **Bài Việt Nam và bài đầu danh sách (QT001, QT010)** chính xác cao hơn — có khả năng được rà kỹ hơn (có file `_FIXED_27052026`).

---

## Khuyến nghị

**Tỷ lệ lỗi 30% (HIGH+MEDIUM) > ngưỡng 20% → CẦN AUDIT TOÀN BỘ.**

Đề xuất hành động ưu tiên cho NCS Công Thị Hằng trước khi dùng trích dẫn:

1. **Audit ưu tiên (P0):** Toàn bộ bài *meta-analysis / GBD / NMA* (QT012, QT018, QT019, QT020, QT028-QT030, QT040, QT052, QT060, QT070-QT072 và các bài tương đương) — đây là nhóm dễ confabulate N và effect size.
2. **Audit P1:** Mọi summary KHÔNG có hậu tố `_FIXED_27052026` — nhóm `_FIXED` (QT003, QT005, QT007, QT010, QT011, VN001, VN017, VN018) đã được rà 1 lượt và cho thấy chất lượng cao hơn.
3. **Quy tắc bắt buộc:** Mọi N, SMD/OR/HR, % prev, tên tác giả, năm xuất bản trước khi vào LA phải kiểm chứng lại với PDF gốc (không tin summary).
4. **Sửa ngay 3 bài flagged HIGH/MEDIUM:**
   - QT020: sửa "SMD 0.96 efficacy" → "SMD 0.96 acceptability (Individual *được chấp nhận hơn* Remote)"; ghi rõ 2 cặp số (abstract vs body).
   - QT060: bỏ "Latin Tropical cao nhất / Đông Á giảm"; thay bằng "middle SDI cao nhất prev+incidence; high SDI tăng mạnh nhất; India số ca tuyệt đối cao nhất; Mexico tăng mạnh nhất".
   - QT072: sửa N **13.497** (không phải 27.133); bổ sung 2 tác giả thiếu (Choo, Zhang Y).

---

## Tham khảo PDF đã đối chiếu

- `02_Papers-goc/11-bai-ban-dau-va-mo-rong/QT001_Jenkins_2023_USA_SanDiego.pdf` (Int J Soc Psychiatry 2023, 69(3):784-794)
- `02_Papers-goc/11-bai-ban-dau-va-mo-rong/QT010_Xu_2021_China_LargestEpi.pdf` (J Affect Disord 2021, 288:17-22)
- `02_Papers-goc/The-gioi_Khac/QT020_Liu_CBT_Delivery_GAD_NMA_2025.pdf` (Transl Psychiatry 2025, 15:197)
- `02_Papers-goc/The-gioi_Khac/QT030_GBD_Trends_10-24y_2025.pdf` (J Affect Disord 2025, 387:119491)
- `02_Papers-goc/The-gioi_Au-My-Uc/QT040_Walder_JMIR_DMHI_SAD_2025.pdf` (JMIR preprint Oct 2024, DOI 10.2196/preprints.67067)
- `02_Papers-goc/The-gioi_Au-My-Uc/QT052_Mindfulness_NatureMH_IPD_MA_2023.pdf` (Nature Mental Health 2023, 1:462-476)
- `02_Papers-goc/The-gioi_Au-My-Uc/QT060_Bie_GlobalAnxiety_GBD_10-24y_1990_2021_FrontPsych_2024.pdf` (Front Psychiatry 2024, 15:1489427)
- `02_Papers-goc/QT072_Lee_2025_CyberbullyingMeta_TVA.pdf` (Trauma Violence Abuse 2025, 27(2):391-406)
- `02_Papers-goc/Viet-Nam/VN001_Hoa_2024_Frontiers_HaNoi.pdf` (Front Public Health 2024, 12:1232856)
- `02_Papers-goc/Viet-Nam/VN002_VNAMHS_2022_National.pdf` (V-NAMHS Report, Institute of Sociology Hanoi, Nov 2022)
