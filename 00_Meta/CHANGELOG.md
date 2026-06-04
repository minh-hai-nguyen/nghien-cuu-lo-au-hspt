# NHẬT KÝ THAY ĐỔI (CHANGELOG)

## 2026-04-05 — Phiên 6: Kiểm tra sạch lỗi + Tổ chức dữ liệu

### Thay đổi
- Thêm phản biện trích dẫn dày cho 12 bài cũ (02-13)
- Dịch đầy đủ QT32 Ireland (2.749 → 15.804 chars), QT33 JAMA (18.481), QT34 Korea (13.447), QT27 Nature (15.946), QT28 AJP (11.902), QT29 BMC (8.785), QT30 GBD (5.500), QT31 59Countries (5.995)
- Sửa Wen 2020 OR: 11,6 → 11,58 (trong đề cương + cross-study)
- Sửa CAMS attribution: Zugman 2024 → Walkup 2008 NEJM
- Tìm + xác minh 4 PDF VN mới đầy đủ (VN15-18)
- Gộp 7 tóm tắt VN14-20 → "Tom tat 7 bai VN - 05042026.docx"
- RAG cập nhật VN15-18 với full PDF content
- Tạo `00_Meta/` (PROJECT_STATE, ERRORS_FIXED, CHANGELOG, PAPERS_INDEX)
- Viết chiến lược tổ chức dữ liệu (`04_Co-so-du-lieu/CHIEN_LUOC_TO_CHUC_DU_LIEU.md`)

### Kiểm tra
- 35 bản dịch: 0 lỗi
- Đề cương: 15/15 trích dẫn khớp PDF gốc
- Cross-study: 15/15 khớp
- VN15-18 vs PDF mới: 19/19 khớp

---

## 2026-04-04 — Phiên 2-5: Dịch QT21-35 + Fix labels + RAG

### Thay đổi
- Dịch đầy đủ QT21-25 (6-19K chars mỗi bài, hình crop từ PDF)
- Dịch tóm lược QT26-34 (2.5-4.6K chars)
- Thêm bài QT35 Social Anxiety 7 Countries (VN SAD = 30,7%)
- Fix labels Bảng 1/2 cho 21 tóm tắt
- Cross-study synthesis 9 bảng 35 bài
- RAG ChromaDB 36 entries
- Quy tắc trích dẫn dày (Mục D trong QUY_TAC)

---

## 2026-03-29 — Phiên 1: Khởi tạo + 11 bài đầu

### Thay đổi
- Dịch + tóm tắt 11 bài đầu tiên
- Cross-study synthesis 10 bài
- Phát hiện 3 lỗi (Wen giới tính, Mandaknalli GAD, hút thuốc)

---

## 2026-03-27 — Phiên 0: Thiết lập dự án

### Thay đổi
- Tạo cấu trúc thư mục
- Đề cương CTH v5 + báo cáo
