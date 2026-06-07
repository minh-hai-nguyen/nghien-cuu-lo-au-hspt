# AUDIT — Tom-tat-tung-bai (07/06/2026)

Phạm vi: thư mục `Tom-tat-tung-bai/` (read-only). Tham chiếu: `02_Papers-goc/canonical_index.json` (117 entries: 78 QT + 24 VN + 15 BB).

## 1. Tổng quan

- Tổng số file .docx tóm tắt: **111**
- File khớp canonical ID (QT/VN/BB prefix): **107**
- File không khớp ID (template/bảng tổng hợp): **4**
- Số canonical ID có ≥1 tóm tắt: **97 / 117** (82,9%)
  - QT có tóm tắt: 73 / 78
  - VN có tóm tắt: 24 / 24 (100%)
  - BB có tóm tắt: 0 / 15

## 2. Canonical IDs KHÔNG có tóm tắt (20)

- **QT thiếu (5):** QT074, QT075, QT076, QT077, QT078
- **VN thiếu (0):** không có
- **BB thiếu (15):** BB01, BB02, BB03, BB05, BB06, BB07, BB08, BB09, BB10, BB11, BB12, BB14, BB18, BB19, BB22 (toàn bộ chuỗi BB chưa có file tóm tắt)

## 3. Cặp trùng lặp cần dedup (10)

So sánh theo mtime + size; cột "GIỮ" là bản nên giữ (mới hơn / không phải backup).

| ID | File A | File B | GIỮ |
|---|---|---|---|
| QT003 | …_2021.docx (04/04) | …_FIXED_27052026.docx (02/06, 41.314 B) | **B** |
| QT005 | …_SaudiArabia.docx (04/04) | …_FIXED_27052026.docx (02/06) | **B** |
| QT007 | …_BMCPsych.docx (04/04) | …_FIXED_27052026.docx (02/06) | **B** |
| QT010 | …_LargestEpi.docx (04/04) | …_FIXED_27052026.docx (02/06) | **B** |
| QT011 | …_Chandigarh.docx (01/06, 39.111 B) | …_BEFORE_FIX44_01062026.docx (01/06, 39.100 B) | **A** (B là snapshot trước fix) |
| QT035 | …_2020.docx (04/04) | …_FIXED_27052026.docx (02/06) | **B** |
| VN001 | …_HaNoi.docx (04/04) | …_FIXED_27052026.docx (02/06) | **B** |
| VN002 | …_National.docx (02/06, 42.091 B) | …_v1_backup.docx (14/04, 38.845 B) | **A** (B là backup) |
| VN017 | …_YHVN.docx (04/04) | …_FIXED_27052026.docx (27/05) | **B** |
| VN018 | …_YHVN.docx (04/04) | …_FIXED_27052026.docx (27/05) | **B** |

Tổng: **10 file thừa** có thể chuyển vào `_Archive/`.

## 4. File orphan (không khớp canonical ID) — 4

- `00_Mẫu tóm tắt bài 1.docx` — template, không phải bài
- `00_Mẫu tóm tắt bài 1_FIXED_27052026.docx` — template phiên bản fix
- `BANG_GIOI_TINH_LIEN_BAI.docx` — bảng tổng hợp giới tính
- `BANG_LOI_DA_SUA.docx` — log lỗi đã sửa

Tất cả đều là tệp phụ trợ hợp lệ, không phải tóm tắt bài; **không cần xoá**.

## 5. Spot-check chất lượng (5 mẫu)

| File | Para | Có Title | Author/Journal | Tóm tắt VN | Reference/DOI | Ghi chú |
|---|---|---|---|---|---|---|
| QT001_Jenkins | 19 | có | có | có (313 ký tự VN) | **THIẾU** | Không thấy mục TLTK/DOI |
| QT020_Liu | 27 | có | có | có | có | Đầy đủ |
| QT040_Walder | 12 | có | có | có | có | Đầy đủ |
| QT060_Bie | 32 | có | **không rõ** | có | **THIẾU** | Ghi chú "cào từ web"; thiếu mục tác giả/tạp chí và DOI |
| QT072_Lee | 8 | có | có | có | có | Đầy đủ nhưng ngắn (8 đoạn) |

Có **2/5 mẫu (40%) thiếu phần Reference/DOI** — QT001 và QT060. Cần kiểm tra hàng loạt mục tham khảo trên toàn bộ 107 tóm tắt.

## 6. Hành động đề xuất

1. **Bổ sung 20 tóm tắt thiếu:** ưu tiên 5 QT mới (QT074–QT078) đã có PDF trong `02_Papers-goc/`; sau đó toàn bộ 15 BB (Báo cáo).
2. **Dedup 10 cặp file:** chuyển 9 file cũ + 1 backup (`VN002_..._v1_backup.docx`, `QT011_..._BEFORE_FIX44_01062026.docx`) vào `_Archive/TomTat_dedup_07062026/`.
3. **Audit Reference toàn bộ:** chạy script duyệt 107 file kiểm tra sự có mặt của mục "Tài liệu tham khảo"/"DOI"; QT001 và QT060 đã xác định thiếu — bổ sung trước.
4. **Chuẩn hoá template:** dùng `00_Mẫu tóm tắt bài 1_FIXED_27052026.docx` làm template chính cho 20 tóm tắt sắp tạo.

---
Nguồn dữ liệu: `02_Papers-goc/canonical_index.json` (117), `Tom-tat-tung-bai/*.docx` (111). Phương pháp: liệt kê file, regex prefix `^(QT\d{3}|VN\d{3}|BB\d{2})`, so sánh mtime/size, python-docx đọc 5 mẫu.
