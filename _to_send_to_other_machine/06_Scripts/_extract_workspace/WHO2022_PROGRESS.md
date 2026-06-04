# WHO 2022 Full Translation — Progress Tracker

**Last update:** 2026-04-30 (phiên bị ngắt)
**Plan file:** `C:\Users\OS\.claude\plans\wise-churning-lamport.md`

## Đã thống nhất với thầy

- **Mục tiêu**: Dịch FULL chi tiết 296 trang WHO 2022 *World Mental Health Report* (Open Access CC BY-NC-SA 3.0 IGO)
- **Tách thành 9 doc** (vì quá dài cho 1 file)
- **Thứ tự ưu tiên**: 3 → 7 → 2 → 6 → 5 → 4 → 1+8 (1 doc/phiên)
- **Workflow**: theo `feedback_research_workflow.md` — page markers cam, abbr lần đầu đỏ đậm, bảng Word tái tạo Box/Table/Figure, dịch 1-1 không bỏ câu, phản biện có dẫn chứng, áp dụng VN cụ thể

## Trạng thái các doc

| Doc | Chương | PDF idx | Tình trạng |
|---|---|---|---|
| Doc 0 | Index + Foreword + Exec Summary | 1-22 | ✅ ĐÃ CÓ — file `WHO_2022_World_Mental_Health_Report_dich_phan_bien_30042026.docx` (V3 30/04, dùng làm Doc 0) |
| **Doc 3** | **Ch 3 — World mental health today** | **57-90** | 🟡 **ĐANG LÀM (~30%)** — script viết được Part A-D (cover/bibliography/abbreviations/opener/Box 3.1), chưa chạy |
| Doc 7 | Ch 7 — Restructuring care | 187-245 | ⏳ chờ |
| Doc 2 | Ch 2 — Principles + drivers | 9-34 | ⏳ chờ |
| Doc 6 | Ch 6 — Promotion + prevention | 147-186 | ⏳ chờ |
| Doc 5 | Ch 5 — Foundations | 103-145 | ⏳ chờ |
| Doc 4 | Ch 4 — Benefits of change | 69-101 | ⏳ chờ |
| Doc 1+8 | Ch 1 + Ch 8 + tổng phản biện | 1-8 + 247-258 | ⏳ chờ |

## Phiên này đã làm

### 1. Sửa GBD 2019 doc (đã xong, đã chạy)
- File: `03_Ban-dich/Bai_dich_phan_bien/GBD_2019_Mental_Disorders_dich_phan_bien_30042026.docx` (~1,4 MB, 8 bảng, 2 hình)
- Fix lỗi LABEL: tỉ lệ HIỆN MẮC bị gán nhầm thành "tỉ lệ DALY" ở Bảng 3 + phản biện điểm 3
- Thay Hình 2 PNG bằng 4 BẢNG CÓ MÀU (YLD-EN/YLD-VI/DALY-EN/DALY-VI) theo yêu cầu thầy "tô màu đúng như bản gốc"
- Script: `06_Scripts/sinh_doc_GBD2019_v3_30042026.py` — có helper `fig2_ranking_table()` + `lighten_hex()` (tái dùng được cho Doc khác)

### 2. Tạo Doc 0 / V3 rút gọn WHO 2022 (đã chạy)
- File: `03_Ban-dich/Bai_dich_phan_bien/WHO_2022_World_Mental_Health_Report_dich_phan_bien_30042026.docx`
- Script: `06_Scripts/sinh_doc_WHO2022_v3_30042026.py`
- Nội dung: bibliography table + 22 abbreviations + Foreword + Chapter 1 brief + 8 Executive Summary cards + Chapter 3 highlights (Bảng 3.1) + Chapter 7 brief + Chapter 8 + Phản biện 6 điểm + Áp dụng VN 6 đề xuất + 15 references

### 3. Chuẩn bị Doc 3 (đang dở)
- **Đã extract text Chapter 3 đầy đủ**: `06_Scripts/_extract_workspace/who_ch3_full.txt` (81.909 ký tự, 2356 dòng)
- **Đã render 12 PNG figures** ở 200 DPI vào `02_Papers-goc/GBD_WHO/figures/`:
  - WHO2022_Fig3_1_service_gap.png (PDF idx 59)
  - WHO2022_Fig3_2_prevalence.png (62)
  - WHO2022_Fig3_3_regional_prev.png (67)
  - WHO2022_Fig3_4_suicide.png (69)
  - WHO2022_Fig3_5_6_DALY_YLD.png (71)
  - WHO2022_Fig3_7_top10_YLD.png (72)
  - WHO2022_Fig3_8_key_gaps.png (74)
  - WHO2022_Fig3_9_research_focus.png (75)
  - WHO2022_Fig3_10_research_basic.png (76)
  - WHO2022_Fig3_11_national_plans.png (77)
  - WHO2022_Fig3_12_estimate_vs_allocate.png (79)
  - WHO2022_Fig3_13_workforce.png (81)
- **Đã viết Part A-D của script Doc 3**: `06_Scripts/sinh_doc_WHO2022_Doc3_Chapter3_30042026.py` (~270 dòng):
  - Trang bìa Doc 3
  - Bảng thông tin thư mục Chương 3
  - Bảng từ viết tắt (22 thuật ngữ)
  - Part A: Trang mở chương (PDF idx 57)
  - Part B: Tóm tắt chương (PDF idx 58)
  - Part C: Mở đầu chương + Hình 3.1 (PDF idx 59)
  - Part D: Box 3.1 Data assessing (PDF idx 60)
- **CHƯA CHẠY** script (vì chưa viết xong)

## Phần còn lại của Doc 3 (cần viết tiếp)

| Part | Nội dung | PDF idx | Sections gốc |
|---|---|---|---|
| E | 3.1.1 Prevalence + Hình 3.2 + Bảng 3.1 đầy đủ + Box 3.2 (COVID) + Box 3.3 (Severity) | 61-66 | trang 39-44 |
| F | Geographical disparities + Hình 3.3 + 3.1.2 Mortality + Suicide + Hình 3.4 + Marie's narrative | 66-70 | trang 44-48 |
| G | 3.1.3 Burden + Hình 3.5 + 3.6 + 3.7 (Top 10 YLD — TÁI TẠO BẢNG CÓ MÀU như Hình 2 GBD) | 70-72 | trang 48-50 |
| H | 3.2 Economic consequences (2,5 → 6 nghìn tỉ USD, Philippines 1,3 tỉ USD) | 72-73 | trang 50-51 |
| I | 3.3 Gaps overview + Hình 3.8 + 3.3.1 Information + Hình 3.9, 3.10 + 3.3.2 Governance + Hình 3.11 + 3.3.3 Resources + Hình 3.12, 3.13 + Eleni narrative + 3.3.4 Services + Mrs BN narrative | 73-86 | trang 51-64 |
| J | 3.4 Barriers + 3.4.1 Poor supply + 3.4.2 Health literacy + Steven narrative + 3.4.3 Stigma + Odireleng narrative | 87-90 | trang 65-68 |
| K | Danh mục tham khảo Chương 3 (~60 refs từ số 96-158) | — | tổng hợp |
| L | Phản biện chi tiết (5-7 điểm, chữ đỏ) | — | mới |
| M | Áp dụng VN (4-5 đề xuất) | — | mới |

## Phiên tiếp theo — TODO khi quay lại

1. **Đọc lại** `06_Scripts/sinh_doc_WHO2022_Doc3_Chapter3_30042026.py` (đã viết Part A-D, dòng cuối có `d.save(OUT); print("Wrote (Part 1/3)")`)
2. **Đọc lại** text Chapter 3: `06_Scripts/_extract_workspace/who_ch3_full.txt` (offset 600+ trở đi cho phần còn lại)
3. **Viết tiếp Part E → M** vào script (dùng helper đã có sẵn — KHÔNG cần định nghĩa lại)
4. **Đối với Hình 3.7** (Top 10 YLDs): tái tạo BẢNG CÓ MÀU theo cách của Hình 2 GBD đã làm (mỗi hàng = 1 nguyên nhân với màu riêng) — dùng `fig2_ranking_table()` HOẶC viết bảng đơn giản với shade
5. **Đối với Hình 3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.8, 3.9, 3.10, 3.11, 3.12, 3.13**: chèn ảnh PNG đã render với caption + key data summary tiếng Việt
6. **Chạy script**, verify 4 vòng (số liệu / thuật ngữ / phản biện / tổng thể)
7. **Đối chiếu key stats** với PDF: 970 triệu, 1/8, 13%, 703.000, 9/100k, 77%, 36% drop, 5,1% DALY, 15,6% YLD, 25%, 28%, 26%, 2,5 → 6 nghìn tỉ USD, 35%, 0,3% dev assistance, <2% ngân sách, 70% spending psych hospital, 67% community, 80%, 71% psychosis untreated, 99% research HIC, 56% basic, 21% rights compliant policies, 51% policies, 38% laws compliant, 8% LIC allocate

## Helpers reusable từ script GBD 2019

```python
# Đã có sẵn trong sinh_doc_WHO2022_Doc3_Chapter3_30042026.py:
shade(cell, color)          # tô màu cell
set_w(cell, width_cm)       # set width cell
tbl(headers, rows, widths)  # bảng chuẩn
title/H1/H2/H3/nr/page_marker/acro/crit_para
insert_image(fname, caption, width_cm=15)
box_open/box_para/box_close                # cho Box 3.x
narrative_open/narrative_para              # cho câu chuyện cá nhân (Marie/Eleni/...)
```

## Memory liên quan

- `feedback_research_workflow.md` — workflow chuẩn
- `feedback_doc_tieng_viet_thuan.md` — tiếng Việt thuần
- `feedback_doc_phai_co_reference.md` — luôn có Tham khảo cuối doc
- Plan file: `C:\Users\OS\.claude\plans\wise-churning-lamport.md`
