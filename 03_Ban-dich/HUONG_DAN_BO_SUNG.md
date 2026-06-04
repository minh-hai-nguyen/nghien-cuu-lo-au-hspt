# HƯỚNG DẪN BỔ SUNG CHO TẤT CẢ BẢN DỊCH
Cập nhật: 03/04/2026

## Cấu trúc bài dịch theo thứ tự:

1. Link bài gốc + QR code
2. Tiêu đề bài báo (Heading 1, song ngữ Việt/Anh)
3. THÔNG TIN THƯ MỤC — info table (tác giả, tạp chí, DOI, mẫu, công cụ...)
4. --- Trang đầu, Tạp chí, Vol --- (màu cam #FF6600, in đậm, cỡ 9pt)
5. **TÓM TẮT** (dịch Abstract gốc) — dòng Phương pháp và Kết quả tô đỏ #CC0000
6. **TÓM TẮT ĐÁNH GIÁ NHANH** ← ngay sau TÓM TẮT, trước GIỚI THIỆU
   - Điểm nổi bật (3-5 câu)
   - Hạn chế (3-5 câu)
   - Hướng cải thiện (3-5 câu)
7. GIỚI THIỆU → PHƯƠNG PHÁP → KẾT QUẢ → THẢO LUẬN → KẾT LUẬN
8. TÀI LIỆU THAM KHẢO (giữ nguyên tiếng Anh)
9. BẢNG VIẾT TẮT (2 cột)
10. **PHẢN BIỆN CHI TIẾT** (cuối bài, tô đỏ #FF0000)
    - Điểm mạnh (3-5 điểm)
    - Hạn chế chi tiết (5-7 điểm)
    - Khoảng trống nghiên cứu / gap (4-6 điểm)

## Màu sắc:
- Page ref: màu cam #FF6600, in đậm, cỡ 9pt, in nghiêng
- Phương pháp/Kết quả quan trọng trong TÓM TẮT: tô đỏ #CC0000
- Phản biện cuối bài: tô đỏ #FF0000
- Header bảng: nền xanh nhạt #D9E2F3

## Bảng:
- Table Grid có viền đầy đủ
- **Tự căn chiều rộng cột phù hợp với dung lượng nội dung** (cột tên dài → rộng, cột số → hẹp)
- Header: in đậm, căn giữa, nền #D9E2F3
- Font bảng: Times New Roman 10pt

## Ảnh/biểu đồ:
- Cắt ảnh từ PDF gốc bằng PyMuPDF (fitz)
- **Ảnh cần cắt gần hết chiều ngang trang** (width ≈ 14-15cm) cho trọn vẹn
- **KHÔNG cắt hết cả trang** — chỉ cắt vùng chứa biểu đồ/hình
- **"Thừa hơn thiếu"** — lấy rộng hơn một chút còn hơn bị cắt mất nội dung
- Chèn đúng vị trí trong bài (ngay sau đoạn text tương ứng)

## Caption ảnh:
- Dưới mỗi ảnh: dịch NGUYÊN VẸN caption gốc tiếng Anh
- Format: in nghiêng, cỡ 9pt, căn giữa
- Giữ cả tên gốc: "Hình 2: ... (Figure 2: ...)"

## Viết tắt:
- Lần đầu: đầy đủ tiếng Việt + (viết tắt — tên đầy đủ tiếng Anh)
- Ví dụ: "Thang Rối loạn Lo âu Tổng quát 7 mục (GAD-7 — Generalized Anxiety Disorder 7-item Scale)"
- Bảng viết tắt cuối bài

## Số liệu:
- Dấu phẩy thập phân: 22,6% (không phải 22.6%)
- Giữ nguyên tất cả: %, p-value, CI, n, mean, SD, OR, AOR
- **Đối chiếu từng con số với PDF gốc** trước khi đưa vào
- **Giới tính: PHẢI xác nhận từ PDF gốc** hướng nào (nam>nữ hay nữ>nam)
- Nếu bài gốc có mâu thuẫn nội bộ (VD Abstract vs Results): ghi chú rõ trong phản biện

## Font và lề:
- Font: Times New Roman 12pt, bảng 10pt, giãn dòng 1.5
- Lề: trên/dưới 2.5cm, trái 3cm, phải 2cm

## Phân loại bài:
- Bài tiếng Anh → dịch đầy đủ theo format trên
- Bài tiếng Việt → giữ nguyên nội dung, chỉ thêm nhận xét đầu bài + phản biện cuối bài
- Chỉ lấy bài từ 2024 trở đi (trừ khi có giá trị đặc biệt)
