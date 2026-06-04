# QUY TẮC DỊCH VÀ TÓM TẮT BÀI BÁO NGHIÊN CỨU
Cập nhật: 04/04/2026

---

## A. BẢN DỊCH (thư mục 03_Ban-dich/)

### Nguyên tắc cốt lõi
> **Bản dịch PHẢI tuân theo cấu trúc bài gốc 100%.**
> Bài gốc bao nhiêu bảng → dịch bấy nhiêu bảng.
> Bài gốc bao nhiêu hình → chèn bấy nhiêu hình.
> Bài gốc bao nhiêu phần → dịch bấy nhiêu phần.
> KHÔNG viết theo template cứng. KHÔNG tóm tắt thay dịch.

### Cấu trúc file dịch

1. **Link bài gốc + QR code**
2. **Tiêu đề song ngữ** (Heading 1 Việt, Heading 2 Anh)
3. **THÔNG TIN THƯ MỤC** — bảng 2 cột (Mục | Chi tiết)
4. **--- Trang [X], [Tạp chí], Vol [Y] ---** (cam #FF6600, đậm, 9pt, nghiêng)
5. **TÓM TẮT ĐÁNH GIÁ NHANH** (phần mình viết thêm — dùng CTH v5):
   - Điểm nổi bật (3-5 câu)
   - Hạn chế (3-5 câu)
   - Hướng cải thiện (3-5 câu)
6. **Nội dung dịch đầy đủ câu-câu:**
   - GIỚI THIỆU → PHƯƠNG PHÁP → KẾT QUẢ → THẢO LUẬN → KẾT LUẬN
   - Xen kẽ dòng tham chiếu trang gốc giữa các phần
   - **TẤT CẢ bảng** trong bài gốc → dịch thành Table Grid
   - **TẤT CẢ hình/biểu đồ** → crop từ PDF gốc, chèn đúng vị trí
7. **TÀI LIỆU THAM KHẢO** (giữ nguyên tiếng Anh)
8. **BẢNG VIẾT TẮT** (2 cột: Viết tắt | Đầy đủ)
9. **QUAN ĐIỂM PHẢN BIỆN / CRITICAL REVIEW** (Heading 2, đỏ #FF0000):
   - Điểm mạnh (3-5 điểm)
   - Hạn chế chi tiết (5-7 điểm)
   - Khoảng trống nghiên cứu / Gap (4-6 điểm)

### Phân loại theo ngôn ngữ
- **Bài tiếng Anh** → dịch toàn văn câu-câu theo format trên
- **Bài tiếng Việt** → giữ nguyên nội dung gốc, chỉ thêm Tóm tắt đánh giá nhanh (đầu) + Phản biện (cuối)

---

## B. TÓM TẮT (thư mục Tom-tat-tung-bai/)

### Nguyên tắc
> Tóm tắt ≤ 2,5 trang, theo 13 bước CTH v5.
> Đủ bảng như bài gốc nếu có thể. Nếu bài gốc có 5+ bảng thì chọn tất cả bảng quan trọng.
> KHÔNG giới hạn cứng 2 bảng — phụ thuộc nội dung bài gốc.
> Mỗi bài khác nhau về số bảng, số đoạn — KHÔNG viết theo công thức máy móc.

### 13 bước CTH v5

| Bước | Tên | Nội dung |
|------|-----|----------|
| 1 | ĐỊNH DANH | Ai, ở đâu, với ai? (1 câu dài) |
| 2 | TỔNG QUAN PP | Phương pháp chính + "Nói cách khác" (KT1) |
| 3 | ĐỊNH NGHĨA | Từng công cụ = 1 đoạn, phả hệ (KT5) |
| 4 | BIỆN MINH | "Tổng quan tài liệu... cho thấy" (KT4) |
| 5 | LIỆT KÊ DỮ LIỆU | Nhân khẩu — câu ngắn |
| 6 | MÔ TẢ QUY TRÌNH | Thiết kế → chọn mẫu → phân tích |
| 7-8 | CHỨNG THỰC | Giá trị phương pháp |
| 9 | ĐỊNH LƯỢNG | Bảng kết quả chính |
| 10 | ĐỐI CHIẾU | KT3: Y văn → Thực tế → Xác nhận (3 bước) |
| 11-12 | NHẬN XÉT | Phát hiện nổi bật + liên bài |
| 13 | KẾT LUẬN | Kiến trúc B: "Dữ liệu... gợi ý rằng..." |

### Sau 13 bước — thêm:
- **Phản biện** (Heading 2, đỏ)
- **Hướng NC tiếp theo** (Heading 2, đỏ)
- **Đánh giá chất lượng** (⭐ — ⭐⭐⭐⭐⭐)

### Chuẩn chất lượng tóm tắt
- Đoạn: 25-30 (tùy bài, KHÔNG cố ép)
- Bảng: đủ bảng quan trọng của bài gốc
- Ký tự: ~4.000-6.000 (tùy bài)
- Font: Times New Roman 12pt, bảng 10pt
- Màu: đỏ #FF0000 (heading phản biện), xanh #0070C0 (nội dung)

---

## C. QUY TẮC CHUNG

### Bảng
- Table Grid có viền đầy đủ
- **Tự căn chiều rộng cột phù hợp nội dung** (cột tên → rộng, cột số → hẹp)
- Header: đậm, căn giữa, nền xanh nhạt #D9E2F3
- Font bảng: Times New Roman 10pt

### Ảnh/biểu đồ
- Crop từ PDF gốc bằng PyMuPDF (fitz)
- **Width ≈ 14-15cm** — gần hết chiều ngang trang
- **KHÔNG cắt hết cả trang** — chỉ cắt vùng chứa hình
- **"Thừa hơn thiếu"** — lấy rộng hơn một chút
- Caption: dịch nguyên vẹn, giữ tên gốc "Hình 2: ... (Figure 2: ...)", nghiêng, 9pt, giữa

### Số liệu
- Dấu phẩy thập phân: 22,6% (KHÔNG 22.6%)
- Giữ nguyên: %, p-value, CI, n, mean, SD, OR, AOR, Beta
- **Đối chiếu từng con số với PDF gốc** trước khi đưa vào
- **Giới tính: PHẢI xác nhận từ PDF gốc** (bài học Wen 2020)

### Viết tắt
- Lần đầu: đầy đủ TV + (viết tắt — tên Anh đầy đủ)
- VD: "Thang đo Rối loạn Lo âu Tổng quát 7 mục (GAD-7 — Generalized Anxiety Disorder 7-item Scale)"

### Phong cách CTH v5
- **Chỉ áp dụng cho phần mình viết** (nhận xét, phản biện, kết luận)
- **KHÔNG sửa cấu trúc/nội dung bài gốc**
- KT1 "Nói cách khác" — giải thích phương pháp đơn giản
- KT3 đối chiếu 3 bước — y văn → thực tế → xác nhận
- KT5 phả hệ — định nghĩa công cụ
- Kiến trúc B — "Dữ liệu... gợi ý rằng..."
- Hedging — "có thể", "gợi ý", "dường như"

### Tiếng Việt
- **Có dấu đầy đủ** — KHÔNG chấp nhận tiếng Việt không dấu
- Kiểm tra bằng regex trước khi xuất file

### Kiểm tra
- So khớp PDF gốc: mọi số liệu
- Kiểm tra không dấu: regex
- Kiểm tra cấu trúc: heading, bảng, text đỏ, viết tắt
- **Tối thiểu 2 vòng kiểm tra liên tiếp không lỗi**

### Font và lề
- Font: Times New Roman 12pt, bảng 10pt
- Giãn dòng: 1,5
- Lề: trên/dưới 2,5cm, trái 3cm, phải 2cm

---

## D. TRÍCH DẪN TRONG PHẢN BIỆN VÀ ĐỀ XUẤT

> **NGUYÊN TẮC QUAN TRỌNG:** Mọi phản biện, đề xuất hướng nghiên cứu, gap analysis trong bản dịch, tóm tắt, đề cương, và cross-study synthesis PHẢI có lượng trích dẫn tương đối dày.

### Yêu cầu trích dẫn
- **Phản biện (Critical Review):** Mỗi điểm mạnh/hạn chế nên trích dẫn ít nhất 1 NC liên quan. VD: "Thiếu biến giấc ngủ — Zhu 2025 cho thấy AOR = 13,71 cho <5h ngủ."
- **Hướng NC tiếp theo:** Mỗi đề xuất cần nêu bằng chứng từ bài khác. VD: "Cần RCT CBT tại trường VN — BMC NMA 2025 (Li et al.) xếp CBT hạng 2 (SUCRA 0,66), AJP 2024 (Zugman et al.) báo cáo CBT+SSRI 80,7% đáp ứng."
- **Gap analysis:** Nêu CỤ THỂ bài nào phát hiện gap, số liệu liên quan. VD: "VN chưa có phân tích SKTT theo thu nhập — Korea 2024 (Cho et al.) cho thấy chênh 22,7 điểm stress giữa nhóm giàu/nghèo."
- **Đối chiếu liên bài:** Luôn so sánh với ít nhất 2-3 bài khác trong Đề tài. VD: "Tỷ lệ 40,6% (Hoa 2024, GAD-7) vs 2,3% (V-NAMHS, DISC-5) vs 86,2% (Bảo Quyên 2025, DASS-21) — chênh lệch do công cụ đo."

### Cách trích dẫn
- Trong bản dịch/tóm tắt: Tên tác giả + năm + số liệu cụ thể. VD: "Jenkins 2023 (n=92, PHQ-9A: 44%)"
- Trong cross-study: Tên tác giả + năm + mã bài (VN1, QT21...) + chỉ số. VD: "Norway 2025 (QT21, n=979K): bất mãn trường giải thích chính"
- Trong đề cương: Trích dẫn đầy đủ hơn, có thể thêm tạp chí. VD: "Zugman et al. (2024, AJP Q1 IF=18): CBT+SSRI đáp ứng 80,7% (CAMS)"

### Mật độ trích dẫn tối thiểu
- Phần Phản biện: ≥ 5 trích dẫn liên bài
- Phần Hướng NC: ≥ 3 trích dẫn liên bài
- Phần Gap: ≥ 2 trích dẫn mỗi gap
- Phần Cross-study: ≥ 3 trích dẫn mỗi nhận định

---

## E. LƯU Ý ĐÃ HỌC TỪ LỖI

| Lỗi | Bài học | Cách tránh |
|-----|---------|------------|
| Wen 2020 nhầm nam>nữ | PHẢI đọc PDF gốc xác nhận giới tính | Tìm OR/% chính xác, không dựa vào tóm tắt cũ |
| Mandaknalli nhãn GAD sai | PDF gốc có thể mâu thuẫn nội bộ (Abstract vs Table) | Ưu tiên Table data, ghi chú mâu thuẫn |
| Mandaknalli hút thuốc đảo | 14,8% là KHÔNG hút, 85,2% là CÓ hút | Đọc kỹ hướng dữ liệu (Yes/No) |
| Tóm tắt QT sơ sài | 8 đoạn, 454 ký tự vs chuẩn 27 đoạn, 5.000 | Đọc PDF trước, viết theo nội dung |
| Bảng máy móc 2/bài | Bài gốc 5 bảng mà tóm tắt chỉ 2 | Đếm bảng/hình trong PDF trước |
| Dịch = tóm tắt | Sếp yêu cầu dịch câu-câu, KHÔNG tóm tắt | Phân biệt rõ bản dịch vs tóm tắt |
