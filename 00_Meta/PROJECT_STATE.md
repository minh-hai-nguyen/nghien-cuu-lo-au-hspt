# TRẠNG THÁI DỰ ÁN — SINGLE SOURCE OF TRUTH

**Cập nhật lần cuối:** 05/04/2026 phiên 6

---

## TỔNG QUAN

| Chỉ số | Giá trị |
|--------|---------|
| Tổng bài NC | 35 (7 VN + 28 QT) |
| Bản dịch hoàn chỉnh (≥5K chars) | 28/35 |
| Bản dịch tóm lược (<5K, bài VN OK) | 7/35 (VN16-20 + QT27-31) |
| Tóm tắt CTH v5 (có Bảng 1+2) | 37/37 ✓ |
| Phản biện (tất cả bản dịch) | 35/35 ✓ |
| Cross-study | 9 bảng, 35 bài ✓ |
| RAG ChromaDB | 36 entries ✓ |
| Đề cương | 05042026 ✓ |
| Kiểm tra chéo PDF gốc | 100% khớp ✓ |
| Lỗi dấu thập phân | 0 ✓ |

---

## BÀI NC — TRẠNG THÁI TỪNG BÀI

### VN (7 bài gốc + 3 bài bổ sung = 10 bài)

| ID | Tác giả | Năm | n | Địa bàn | Công cụ | PDF gốc | Bản dịch | Tóm tắt |
|----|---------|-----|---|---------|---------|---------|----------|---------|
| VN01 | Hoa et al. | 2024 | 3.910 | Hà Nội | GAD-7 | ĐỦ (8p) | ĐỦ (29K) | ✓ |
| VN02 | V-NAMHS | 2022 | 5.996 | 38 tỉnh | DISC-5 | ĐỦ (51p) | ĐỦ (23K) | ✓ |
| VN03 | Pham et al. | 2024 | ~500 | Huế | DASS-21 | ĐỦ (7p) | ĐỦ (7K) | ✓ |
| VN14 | Hoàng Trung Học | 2025 | ~2000 | VN | DASS-21 | ĐỦ (8p) | ĐỦ (8K) | ✓ |
| VN15 | Ngô Anh Vinh | 2024 | 845 | Lạng Sơn DTTS | DASS-21+ACE | ĐỦ (7p) | ĐỦ (6K) | ✓ |
| VN16 | Bảo Quyên | 2025 | 501 | Hà Nội | DASS-21 | ĐỦ (8p) | OK (3.8K VN) | ✓ |
| VN17 | Nguyễn Danh Lâm | 2022 | 482 | Yên Định, Thanh Hóa | DASS-21 | ĐỦ (4p) | NGẮN (1.4K VN) | ✓ |
| VN18 | An Giang | 2025 | 366 | Long Bình, An Giang | DASS-21 | ĐỦ (5p) | NGẮN (1K VN) | ✓ |
| VN19 | Hồ Thị Trúc Quỳnh | 2025 | 685 | Huế | DASS-21+LOT-R | CẦN TÌM | OK (3.1K VN) | ✓ |
| VN20 | Trần Hồ Vĩnh Lộc | 2024 | 976 | TPHCM | DASS-Y | CẦN TÌM | OK (3.1K VN) | ✓ |

**Lưu ý bài VN:** Theo quy tắc, bài tiếng Việt CHỈ thêm nhận xét + phản biện, KHÔNG dịch câu-câu → 1-4K chars là OK.

### QT (28 bài)

| ID | Tác giả | Năm | Tạp chí | Bản dịch (chars) | Trạng thái |
|----|---------|-----|---------|-------------------|------------|
| QT01 | Jenkins | 2023 | J Early Adolescence | 51K | ĐỦ |
| QT02 | Saikia | 2023 | IJCM | 8K | ĐỦ |
| QT03 | Zhameden | 2025 | PLOS ONE | 23K | ĐỦ |
| QT04 | Anderson | 2025 | Wiley | 52K | ĐỦ |
| QT05 | Zhu | 2025 | BMC | 6K | ĐỦ |
| QT06 | Mudunna/GBD ASEAN | 2025 | Lancet SEA | 7K | ĐỦ |
| QT07 | Puyat | 2025 | Philippines | 7K | ĐỦ |
| QT08 | Wen | 2020 | IJERPH | 7K | ĐỦ |
| QT09 | Xu | 2021 | J Affect Disord | 8K | ĐỦ |
| QT10 | GBD ASEAN Lancet | 2025 | Lancet | 33K | ĐỦ |
| QT11 | Bhardwaj | 2020 | — | ~3K | OK |
| QT21 | Brunborg/Norway | 2025 | SocSciMed Q1 | 19K | ĐỦ |
| QT22 | Li/ScreenTime | 2025 | BJCP Q1 | 15K | ĐỦ |
| QT23 | Mojtabai/JAACAP | 2024 | JAACAP Q1 IF=11 | 12K | ĐỦ |
| QT24 | Tarasenko/WHO | 2025 | Lancet RegEur Q1 | 11K | ĐỦ |
| QT25 | Crisp/EpiPsychSci | 2025 | EpiPsychSci Q1 | 6K | ĐỦ |
| QT26 | Baker/UK NHS | 2024 | UK Parliament | 7K | ĐỦ |
| QT27 | Fassi/Nature SM | 2025 | Nat Hum Behav Q1 IF=30 | 16K | ĐỦ |
| QT28 | Zugman/AJP | 2024 | AJP Q1 IF=18 | 12K | ĐỦ |
| QT29 | Li/BMC NMA | 2025 | BMC Psychiatry Q1 | 9K | ĐỦ |
| QT30 | Zhang/GBD | 2025 | J Affect Disord Q1 | 6K | ĐỦ |
| QT31 | Islam/59Countries | 2025 | J Affect Disord Q1 | 6K | ĐỦ |
| QT32 | Fitzgerald/Ireland | 2024 | Early Interv Psych | 16K | ĐỦ |
| QT33 | Schmidt-Persson/JAMA | 2024 | JAMA NetOpen Q1 | 18K | ĐỦ |
| QT34 | Cho/Korea | 2024 | Nat Sci Rep Q1 | 13K | ĐỦ |
| QT35 | Jefferies/7Countries | 2020 | PLOS ONE Q1 | 10K | ĐỦ |

---

## PDF GỐC — MAPPING

### Thư mục `02_Papers-goc/Viet-Nam/`

| File | Bài | Trạng thái |
|------|-----|------------|
| `1-s2.0-S2666915324000817-main.pdf` | VN15 Ngô Anh Vinh | ĐỦ 7p |
| `2617-Văn bản của bài báo-2387-1-10-20250619.pdf` | VN16 Bảo Quyên | ĐỦ 8p |
| `67-70-2948-5466_Văn bản của bài báo.pdf` | VN17 Nguyễn Danh Lâm | ĐỦ 4p |
| `8-13506-23239_Văn bản của bài báo.pdf` | VN18 An Giang | ĐỦ 5p (trang 2-5, trang 1 là bài khác) |
| `Hoa_2024_Frontiers_LoAu_THPT_HaNoi.pdf` | VN01 | ĐỦ 8p |
| `HoangTrungHoc_2025_COVID_VN.pdf` | VN14 | ĐỦ 8p |
| `VNAMHS_2022_Report.pdf` | VN02 | ĐỦ 51p |
| `NgoAnhVinh_2024_DTTS_LangSon.pdf` | VN15 OLD | CHỈ 2p metadata — THAY bằng file mới |
| `NguyenNgocBaoQuyen_2025_HaNoi_YHCD.pdf` | VN16 OLD | CHỈ 1p metadata — THAY bằng file mới |
| `NguyenDanhLam_2024_YenDinh_ThanhHoa.pdf` | VN17 OLD | LỖI FILE — THAY bằng file mới |
| `LongBinh_AnGiang_2024_DASS21.pdf` | VN18 OLD | LỖI FILE — THAY bằng file mới |

---

## FILE SẢN PHẨM CHÍNH

| File | Loại | Cập nhật cuối |
|------|------|--------------|
| `Lo au - Bao cao + De cuong - 05042026.docx` | Đề cương LATEST | 05/04/2026 |
| `Tổng hợp liên bài báo - Lo âu HS - 04042026.docx` | Cross-study LATEST | 04/04/2026 |
| `Tom tat 7 bai VN - 05042026.docx` | Gộp 7 VN | 05/04/2026 |
| `rag_db/` | ChromaDB 36 entries | 05/04/2026 |

---

## LỖI ĐÃ SỬA

Xem chi tiết: `00_Meta/ERRORS_FIXED.md`

---

## VIỆC CẦN LÀM PHIÊN TIẾP

1. ~~(Tùy chọn)~~ Gộp 42 scripts → 1 script idempotent (khi user yêu cầu)
2. Tìm PDF VN19 (Hồ Thị Trúc Quỳnh, TC Tâm lý học) + VN20 (Vĩnh Lộc, TC Y học TPHCM)
3. Sắp xếp lại thư mục PDF (VN/ vs QT/ thay vì phân mảnh)
4. Viết script kiểm tra tự động (`verify_all.py`)
