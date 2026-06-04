# MẪU PROMPT PHÂN TÍCH BÀI BÁO NGHIÊN CỨU

> Lưu lại từ phiên làm việc ngày 24–25/03/2026. Dùng để tái sử dụng khi cần phân tích các bài báo nghiên cứu tương tự.

---

## 1. PROMPT TÌM BÀI BÁO GỐC

```
Bạn là một học giả nghiên cứu. Hãy tìm các bài báo nghiên cứu/chương sách bằng các thứ tiếng phổ biến có uy tín về [CHỦ ĐỀ]. Cơ sở dữ liệu gồm có: tên bài báo, tác giả, năm xuất bản, tóm tắt nội dung song ngữ Anh/Việt, đánh giá chất lượng dựa vào lượt cite từ các nguồn uy tín và chính của tạp chí xuất bản. Nếu có thể thì tải về và lưu bản gốc của bài báo. Chỉ quét các bài từ 2020 lại đây, xa nhất là 2017 nếu công trình thật sự giá trị và đang mở ra các hướng nghiên cứu có giá trị cho tới hiện tại.
```

---

## 2. PROMPT GỢI Ý BÀI THAY THẾ CHẤT LƯỢNG CAO HƠN

```
Sếp đang quan tâm đến [SỐ] bài này: [LIỆT KÊ]. Hãy tìm thêm bài về cùng chủ đề/nội dung có chất lượng cao hơn và mới hơn. Với mỗi bài gốc, phân tích hạn chế (IF tạp chí, mẫu, phương pháp) rồi gợi ý 1–3 bài thay thế/bổ sung từ tạp chí IF cao hơn, mẫu lớn hơn, phương pháp mạnh hơn (meta-analysis, RCT, khảo sát quốc gia).
```

---

## 3. PROMPT ĐỌC & VIẾT BÁO CÁO TỔNG HỢP

```
Đọc các bài trong [THƯ MỤC], viết báo cáo tóm tắt, điểm nổi bật, phản biện, liệt kê một số hướng nghiên cứu để hoàn thiện và phát triển thêm, đề xuất danh sách các bài cần đọc liên quan. Mỗi đoạn cần chi tiết 3–7 câu. Tách câu dài thành câu ngắn cho dễ đọc. Viết thành file doc, bảng phải kẻ thành bảng. Viết với Unicode đầy đủ.
```

---

## 4. YÊU CẦU KỸ THUẬT CHO FILE DOCX

### 4.1 Định dạng chung
- **Font:** Times New Roman
- **Cỡ chữ nội dung:** 12pt
- **Cỡ chữ trong bảng:** 10–11pt
- **Giãn dòng:** 1.5
- **Lề trang:** Trên 2.5cm, Dưới 2.5cm, Trái 3cm, Phải 2cm
- **Ngôn ngữ:** Unicode tiếng Việt đầy đủ có dấu

### 4.2 Bảng (Table)
- **Style:** Table Grid (có kẻ viền đầy đủ)
- **Header:** In đậm, căn giữa, tô nền màu xanh nhạt (#D9E2F3)
- **Căn lề bảng:** Căn giữa trang

### 4.3 Tỷ lệ chiều rộng cột (cho bảng tham khảo bài báo)

| Cột | Chiều rộng | Tỷ lệ | Ghi chú |
|-----|-----------|-------|---------|
| # (STT) | 0.6 cm | ~4% | Chỉ chứa số |
| Tác giả | 2.5 cm | ~16% | Tên ngắn gọn |
| Năm | 0.8 cm | ~5% | 4 chữ số |
| Tạp chí | 2.8 cm | ~17% | Tên viết tắt + IF |
| Nội dung | 9.3 cm | ~58% | Chi tiết 3–5 câu |
| **Tổng** | **16.0 cm** | **100%** | Khổ A4, lề trái 3cm + phải 2cm |

### 4.4 Tỷ lệ chiều rộng cột (cho bảng thông tin bài báo — 2 cột)

| Cột | Chiều rộng | Ghi chú |
|-----|-----------|---------|
| Mục | 3.0 cm | Tên trường (Tác giả, Tạp chí, Mẫu...) |
| Chi tiết | 13.0 cm | Nội dung chi tiết |

### 4.5 Tỷ lệ chiều rộng cột (cho bảng kết quả điều trị — 5 cột)

| Cột | Chiều rộng | Ghi chú |
|-----|-----------|---------|
| Chỉ số | 4.0 cm | Tên triệu chứng/chỉ số |
| T0 | 2.5 cm | Giá trị trước điều trị |
| T2 | 2.5 cm | Giá trị tuần 2 |
| T4 | 2.5 cm | Giá trị tuần 4 |
| p | 2.0 cm | Giá trị p |

---

## 5. CẤU TRÚC BÁO CÁO CHUẨN

```
I.   TÓM TẮT TỪNG BÀI
     - Bảng thông tin (2 cột: Mục | Chi tiết)
     - Phương pháp (1 đoạn, 3–5 câu)
     - Kết quả chính (2–3 đoạn, mỗi đoạn 3–5 câu)
     - Bảng kết quả nếu có dữ liệu định lượng

II.  ĐIỂM NỔI BẬT
     - 5–8 điểm, mỗi điểm gồm:
       + Tiêu đề in đậm (1 câu)
       + Nội dung (3–5 câu)

III. PHẢN BIỆN
     - Mỗi bài 1 mục, 5–7 câu
     - Tập trung: thiết kế NC, mẫu, công cụ, phân tích, tổng quát hóa

IV.  HƯỚNG NGHIÊN CỨU
     - 10–15 hướng, chia nhóm (A, B, C, D)
     - Mỗi hướng 3–4 câu

V.   DANH SÁCH BÀI CẦN ĐỌC
     - Bảng 5 cột (STT, Tác giả, Năm, Tạp chí, Nội dung)
     - Cột Nội dung: 3–5 câu chi tiết
     - Chiều rộng cột theo tỷ lệ mục 4.3
     - Chia thành 3–4 nhóm chủ đề
```

---

## 6. GHI CHÚ BỔ SUNG

- **Đổi tên file:** Theo mẫu `[Chủ đề] - [Phạm vi] - [DDMMYYYY]`. Ví dụ: `Lo âu - Việt Nam - 25032026.docx`
- **Khi quét bài:** Nên quét ít nhất 2–3 lần để bao phủ đủ các góc độ. Mỗi lần quét tập trung vào các chủ đề chưa bao phủ.
- **Khi gợi ý thay thế:** Luôn so sánh IF tạp chí, cỡ mẫu, phương pháp (meta-analysis > RCT > cắt ngang > mô tả) giữa bài gốc và bài thay thế.
- **Tải PDF:** Ưu tiên nguồn open-access (Frontiers, PLOS, PMC, Nature OA). Kiểm tra file >50KB mới là PDF thật (file <50KB thường là trang CAPTCHA).

---

> *Tạo ngày 25/03/2026.*
