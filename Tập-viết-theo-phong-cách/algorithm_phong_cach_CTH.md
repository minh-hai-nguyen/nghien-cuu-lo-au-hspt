# Thuật toán bắt chước phong cách viết Công Thị Hằng (v5)

> **Triết lý**: Mô hình hóa CÁCH TƯ DUY, không liệt kê quy tắc bề mặt.
> 8 tầng từ sâu (mục đích) ra nông (dấu câu). Mỗi tầng sâu quyết định tầng nông.
>
> **Mới ở v5**: Tích hợp 7 chiều phong cách từ nghiên cứu phong cách học
> (stylometry, metadiscourse, thematic progression, evidentiality, discourse prosody)
> + đo lường định lượng thực tế từ bản gốc.

---

## TẦNG 0 — THỂ LOẠI VÀ MỤC ĐÍCH

```
Thể_loại = "Tóm tắt từng bài để đưa vào Báo cáo tổng hợp nghiên cứu"

Mục_đích:
  - Tóm tắt MỘT nghiên cứu cho người đọc tổng hợp
  - Format NHẤT QUÁN để ghép nhiều tóm tắt
  - Đối tượng đọc: nhà nghiên cứu chuyên ngành

Mô_hình_người_đọc:           ← MỚI v5
  - Giả định người đọc BIẾT: phương pháp thống kê (p-value, Mann-Whitney),
    khái niệm nghiên cứu cơ bản → KHÔNG giải thích
  - Giả định người đọc KHÔNG BIẾT: công cụ đo lường cụ thể
    → PHẢI định nghĩa từng công cụ (phả hệ + nguồn)
  - Giả định người đọc cần ĐÁNH GIÁ chất lượng
    → PHẢI biện minh lựa chọn phương pháp
```

---

## TẦNG 1 — MÔ HÌNH THÔNG TIN 4 LỚP

```
Lớp 1 — SỰ KIỆN:      Cái gì đã được làm / phát hiện?
Lớp 2 — PHƯƠNG PHÁP:   Tại sao làm theo cách này?
Lớp 3 — VỊ TRÍ:        Kết quả nằm ở đâu trong bức tranh lớn?
Lớp 4 — DIỄN GIẢI:     Điều này có thể ngụ ý gì? (luôn thận trọng)

Phân bố theo phần:
  Phương pháp → Lớp 1 + Lớp 2
  Kết quả    → Lớp 3 → Lớp 1 (người khác thấy → mình thấy)
  Nhận xét   → Lớp 1 + mở rộng phạm vi
  Kết luận   → Lớp 1 → Lớp 4 (dữ liệu → diễn giải thận trọng)
```

---

## TẦNG 2 — CHUỖI 13 BƯỚC TU TỪ

```
 1. ĐỊNH DANH         [2]     Ai, ở đâu, với ai? (1 câu dài)
 2. TỔNG QUAN PP      [4]     Phương pháp chính + "Nói cách khác" rất ngắn
 3. ĐỊNH NGHĨA        [5-6]   Từng công cụ = 1 đoạn, phả hệ
 4. BIỆN MINH          [7]     "Tổng quan tài liệu... cho thấy" + chuỗi trích dẫn
 5. LIỆT KÊ DỮ LIỆU  [8]     Nhân khẩu — câu ngắn
 6. MÔ TẢ QUY TRÌNH   [9]     Phỏng vấn: thiết kế → chủ đề → kỹ thuật
 7+8. CHỨNG THỰC + MINH HỌA [10]  Giá trị PP hỗn hợp + "Ví dụ... Ngược lại..."
 9. ĐỊNH LƯỢNG         [15]    Tần suất → tổng hợp
10. ĐỐI CHIẾU         [16]    Dẫn NC → "Thực tế" → "Xác nhận" → "Tương tự"
11+12. KHÁM PHÁ + CẦU NỐI [19-21]  Phát hiện định tính → ý nghĩa rộng
13. HÀM Ý             [23]    "Dữ liệu... gợi ý rằng..."
```

---

## TẦNG 3 — NHỊP THU PHÓNG + NHỊP ĐỘ DÀI CÂU

### 3A. Thu phóng tổng quát ↔ cụ thể (giữ từ v4)

```
Quy tắc: Không quá 2-3 đoạn liên tiếp ở cùng mức phóng đại.
Sau phóng vào → phải phóng ra (hoặc ngược lại).
Câu/đoạn cuối mỗi PHẦN luôn ở mức TỔNG QUÁT.
```

### 3B. Nhịp độ dài câu — Discourse Prosody (MỚI v5)

> Đo từ bản gốc: mean = 29,6 từ, median = 27, stdev = 17, range = 3-72.
> Độ lệch chuẩn CAO (17) nghĩa là CTH THAY ĐỔI ĐỘ DÀI CÂU RẤT MẠNH.
> Đây không phải văn đều đều — mà là nhịp sóng: ngắn-dài-ngắn-dài.

```
Chuỗi độ dài câu thực tế (từ/câu):
  PP: 27→12→28→23→62→20→27→56→32→72→44→33
  KQ: 21→13→61→22→25→37→25
  NX: 47→30→36→20→48
  KL: 62→34→34→43

Mẫu nhịp theo phần:
  PP:  ██░██░█████░██░████░██░█████████  (sóng lớn, biên độ rộng)
  KQ:  ██░░████░██░██░███░██             (sóng vừa)
  NX:  ████░██░███░██░████              (sóng vừa)
  KL:  █████░███░███░███                (đều hơn, ở mức dài)

QUY TẮC NHỊP:
  → Sau 1 câu dài (>40 từ), nên có 1 câu ngắn hơn (<25 từ), hoặc ngược lại
  → Phần Kết luận: câu đều dài hơn trung bình (>30 từ)
  → Phần Định nghĩa: câu đều ngắn (10-28 từ)
  → KHÔNG BAO GIỜ viết 3+ câu liên tiếp cùng độ dài (±5 từ)
```

### 3C. Tiến trình chủ đề — Thematic Progression (MỚI v5)

> Dựa trên lý thuyết Đề-Thuyết (Theme-Rheme) của Halliday.
> CTH dùng HỖN HỢP các mẫu tiến trình, không chỉ một kiểu.

```
4 MẪU TIẾN TRÌNH CTH SỬ DỤNG:

1) TUYẾN TÍNH (Rheme câu trước → Theme câu sau):
   "Phỏng vấn được thiết kế để tìm hiểu [các chủ đề]."
   → "Các chủ đề phỏng vấn bao gồm..."
   → Dùng khi: xây dựng lập luận từng bước

2) DẪN XUẤT (Từ 1 chủ đề lớn → nhiều chủ đề con):
   "Công trình sử dụng PHQ-9A và GAD-10..."
   → "PHQ-9A là..." (đoạn riêng)
   → "GAD-10 là..." (đoạn riêng)
   → Dùng khi: liệt kê công cụ, tiểu mục, kết quả phụ

3) HẰNG SỐ (Cùng chủ đề lặp lại):
   "Khác biệt giới tính : Nữ giới..."
   "Thực tế, sự chênh lệch giới..."
   "Chúng tôi xác nhận..."
   → Dùng khi: xoay quanh 1 chủ đề từ nhiều góc (KỸ THUẬT 3)

4) SONG SONG (Hai chủ đề đặt cạnh nhau):
   "Ví dụ, trong khi [A]..."
   "Ngược lại, [B]..."
   → Dùng khi: so sánh, minh họa đối xứng (KỸ THUẬT 2)

QUY TẮC:
  → Đoạn Phương pháp: chủ yếu TUYẾN TÍNH + DẪN XUẤT
  → Đoạn Kết quả (so sánh): chủ yếu HẰNG SỐ
  → Đoạn Ví dụ đối xứng: SONG SONG
  → KHÔNG dùng mẫu HẰNG SỐ quá 4 câu liên tiếp (sẽ monotone)
```

---

## TẦNG 4 — KIẾN TRÚC CÂU + PHÂN BỐ TÁC TỬ (Agency)

### 4A. Bốn kiến trúc câu (giữ từ v4)

```
A — "Thác đổ sang phải": Main → participial → relative → citation
    Dùng cho: lập luận, biện minh
B — "Bằng chứng chèn giữa": Subject, [evidence], main verb [implication]
    Dùng cho: kết luận
C — "Song song đối chiếu": "trong khi [A], [B]" + "Ngược lại, [A'] mà [B']"
    Dùng cho: minh họa, so sánh
D — "Câu đơn chức năng": 1 mệnh đề, ngắn, thông tin nền
    Dùng cho: định nghĩa, liệt kê
```

### 4B. Phân bổ tác tử — Agency Distribution (MỚI v5)

> Phát hiện quan trọng từ đo lường: "chúng tôi" xuất hiện 5 lần:
> 4 lần SỞ HỮU ("của chúng tôi") + 1 lần TÁC TỬ ("Chúng tôi xác nhận")
> Tỉ lệ 4:1 → CTH dùng "chúng tôi" chủ yếu để SỞ HỮU phương pháp,
> rất hiếm khi dùng làm chủ ngữ hành động.

```
AI LÀM GÌ TRONG CÂU — PHÂN BỐ CỐ ĐỊNH:

1) "của chúng tôi" + danh từ phương pháp → SỞ HỮU (phổ biến nhất)
   VD: "Tổng quan tài liệu của chúng tôi cho thấy..."
       "Phương pháp hỗn hợp của chúng tôi sử dụng..."
       "Dữ liệu của chúng tôi... gợi ý rằng..."
   → Chủ ngữ thực sự là "tổng quan", "phương pháp", "dữ liệu" — không phải "chúng tôi"
   → Hệ quả: PHƯƠNG PHÁP và DỮ LIỆU là tác tử, con người đứng sau

2) "Chúng tôi" làm tác tử trực tiếp → CỰC HIẾM (1 lần trong toàn bài)
   VD: "Chúng tôi xác nhận trong nhóm thanh thiếu niên sớm này rằng..."
   → Chỉ dùng ở bước 3 của KỸ THUẬT 3 (xác nhận kết quả)
   → Đây là khoảnh khắc duy nhất CTH "bước ra" nhận trách nhiệm trực tiếp

3) Chủ ngữ VÔ TRI (inanimate agent) → PHỔ BIẾN
   VD: "Công trình này sử dụng...", "Mức trung bình... ít được quan sát thấy hơn"
   → Dùng khi: mô tả phương pháp, trình bày kết quả

4) Cấu trúc BỊ ĐỘNG "được + V" → PHỔ BIẾN (11 lần, 9,6% mệnh đề)
   VD: "được thiết kế để", "được tiến hành bằng", "được ghi âm để"
   → Dùng khi: mô tả quy trình, công cụ

5) DANH HÓA (biến V → N bằng "sự/tính/việc") → RẤT CAO (15 lần)
   VD: "sự hiện diện", "sự chênh lệch", "tính bổ sung", "việc nhận diện"
   → Hiệu ứng: xóa tác tử, biến hành động thành khái niệm trừu tượng
   → Tạo giọng khách quan, khoa học, phi cá nhân

QUY TẮC TỔNG:
  → "của chúng tôi" + danh từ: NHIỀU (3-5 lần/bài)
  → "Chúng tôi" làm chủ ngữ hành động: TỐI ĐA 1-2 lần/bài
  → Chủ ngữ vô tri + danh hóa + bị động: CHIẾM ĐA SỐ câu
  → Hiệu ứng tổng thể: con người ẩn sau, dữ liệu & phương pháp "tự nói"
```

---

## TẦNG 5 — HỆ THỐNG SIÊU DIỄN NGÔN (Metadiscourse) — MỚI v5

> Dựa trên khung Hyland (2005). v4 chỉ có hedging 3 cấp.
> v5 mở rộng thành hệ thống siêu diễn ngôn đầy đủ.

### 5A. Markers tương tác (Interactive — tổ chức văn bản)

```
Từ nối chuyển tiếp (Transition):
  "Nói cách khác" (diễn giải), "Ngược lại" (đối chiếu),
  "Tương tự" (bổ sung), "Thực tế" (xác nhận)
  → Mỗi từ nối có VỊ TRÍ CỐ ĐỊNH trong bài (xem Tầng 2)

★ TỪ NỐI NHÂN QUẢ — KHÔNG BAO GIỜ DÙNG:
  CTH KHÔNG dùng: "Vì vậy", "Do đó", "Bởi vì", "Cho nên", "Vì thế"
  CTH để liên kết nhân quả NGẦM — người đọc tự suy ra từ trình tự ý.
  VD gốc: [7] biện minh (vấn đề) → [8] dữ liệu → [9] phỏng vấn
          (KHÔNG có "Vì vậy" giữa các đoạn)

Từ nối khung (Frame markers):
  "Về [chủ đề]," (đánh dấu chủ đề mới)
  "Khi gộp lại" (đánh dấu tổng hợp)
  → Dùng ở đầu đoạn kết quả

Tham chiếu nội bộ (Endophoric):
  "(Hình 2)" — tham chiếu hình/bảng
  → Dùng khi có bảng biểu

Giải mã thuật ngữ (Code glosses):
  "Nói cách khác" (reformulation)
  Tên TV (Tên tiếng Anh) — song ngữ giải thích
  "(Bệnh virus corona 2019 - Coronavirus Disease 2019)" — giải thích viết tắt
  → Dùng khi giới thiệu thuật ngữ lần đầu
```

### 5B. Markers liên nhân (Interactional — quan hệ tác giả-người đọc)

```
HEDGES (thận trọng) — TỶ LỆ RẤT CAO:
  "có thể" (7 lần), "cho thấy" (3 lần), "gợi ý rằng" (1 lần)
  → Đo được: 11 hedges trong 1.156 từ = ~1 hedge / 105 từ

BOOSTERS (tăng cường) — CỰC HIẾM:
  "luôn" (1 lần: "luôn báo cáo")
  → Đo được: 1 booster trong 1.156 từ

★ TỶ LỆ HEDGE:BOOSTER = 11:1
  → Đây là DẤU VÂN TAY MẠNH NHẤT của CTH
  → Hệ quả: gần như KHÔNG BAO GIỜ dùng từ tăng cường
  → Ngoại lệ duy nhất: "luôn" khi dẫn nhận định chung của ngành
    ("Nữ giới LUÔN báo cáo mức trầm cảm cao hơn")
    → "luôn" ở đây trích dẫn BÊN NGOÀI, không phải khẳng định của CTH

ATTITUDE MARKERS (đánh giá) — KHÔNG CÓ:
  → CTH không dùng "đáng chú ý", "đáng ngạc nhiên", "thú vị"
  → Hoàn toàn trung tính

ENGAGEMENT MARKERS (kéo người đọc) — KHÔNG CÓ:
  → CTH không dùng "cần lưu ý rằng", "ta thấy rằng"
  → Người đọc không được "mời" — chỉ được "trình bày"

SELF-MENTIONS — PATTERN ĐẶC THÙ:
  → Xem Tầng 4B: chủ yếu sở hữu, cực hiếm tác tử
```

### 5C. Tích hợp nguồn — Evidentiality Pattern (MỚI v5)

```
TRÍCH DẪN TÍCH HỢP (Integral — tác giả trong câu văn):
  "do Janis H. Jenkins và cs. (2023) khảo sát..."
  → Dùng ở: Bước 1 Định danh (1 lần duy nhất trong bài)

TRÍCH DẪN KHÔNG TÍCH HỢP (Non-integral — tác giả chỉ trong ngoặc):
  "Nữ giới luôn báo cáo mức trầm cảm cao hơn (Salk và cộng sự, 2017)"
  → Dùng ở: mọi nơi khác (chiếm đa số)

★ TỶ LỆ: Non-integral chiếm >80% trích dẫn
  → Hệ quả: CTH để NỘI DUNG nổi bật, tên tác giả lui về phía sau
  → Ngoại lệ: chỉ bước Định danh mới nêu tên tác giả trong câu

ĐỘNG TỪ BÁO CÁO (Reporting verbs):
  "cho thấy" (show — trung tính)
  "chứng minh" (demonstrate — mạnh hơn, dùng cho NC trước)
  "xác nhận" (confirm — dùng cho kết quả của mình phù hợp NC trước)
  "báo cáo" (report — trung tính, dùng khi dẫn phát hiện người khác)
  → CTH KHÔNG dùng: "khẳng định" (claim), "tranh luận" (argue)
  → Hệ quả: giọng hòa nhã, không đặt câu hỏi về nghiên cứu khác
```

---

## TẦNG 6 — KỸ THUẬT VIẾT ĐOẠN

### 6.1 Năm kỹ thuật đặc trưng (giữ từ v4)

```
KT1 — "Phức tạp → Đơn giản": Câu dài → "Nói cách khác" câu <15 từ
KT2 — "Ví dụ đối xứng": "Ví dụ, trong khi..." + "Ngược lại,..."
       + "có thể" ở CẢ HAI vế, CẢ HAI câu
KT3 — "Xác nhận 3 bước": NC trước → "Thực tế," → "Chúng tôi xác nhận"
KT4 — "Biện minh bằng tổng quan": + chuỗi 4-6 nguồn
KT5 — "Định nghĩa phả hệ": "là phiên bản [đối tượng] của thang [gốc]"
```

### 6.2 Ba cách mở đầu đoạn nhận xét (giữ từ v4, XOAY VÒNG)

```
A — "Nổi bật đối với [nhóm] là..."
B — "[Cảm xúc] trước nhận thức của [đối tượng] về..."
C — "Các cuộc phỏng vấn [loại] cho thấy..."
```

### 6.3 Cầu nối phạm vi cuối đoạn nhận xét (giữ từ v4)

### 6.4 Báo cáo số liệu kép (giữ từ v4)

```
"khoảng một nửa (50,6%)" — ngôn ngữ tự nhiên + con số chính xác
```

### 6.5 Đa dạng mở đầu câu — Sentence Opener Variety (MỚI v5)

> Đo từ bản gốc: 39 câu, gần như KHÔNG có 2 câu liên tiếp mở đầu giống nhau.
> Đây là đặc điểm mạnh — CTH tránh lặp chủ ngữ mở đầu.

```
PHƯƠNG PHÁP ĐẢM BẢO ĐA DẠNG:
Xoay vòng giữa các kiểu mở đầu:

  a) Danh từ phương pháp: "Công trình này...", "PHQ-9A..."
  b) Sở hữu: "Tổng quan tài liệu của chúng tôi...", "Dữ liệu của chúng tôi..."
  c) Cụm giới từ: "Về tần suất,", "Trong bối cảnh..."
  d) Từ nối: "Nói cách khác,", "Ví dụ,", "Ngược lại,", "Thực tế,", "Tương tự,"
  e) Nhãn + hai chấm: "Khác biệt giới tính :"
  f) Chỉ thị: "Tất cả các cuộc phỏng vấn...", "Các chủ đề phỏng vấn..."
  g) Danh hóa: "Những phát hiện...", "Sự khác biệt..."
  h) Đại từ chỉ định: "Đây là...", "Những vấn đề này..."

QUY TẮC: Không bao giờ dùng cùng 1 kiểu mở đầu cho 2 câu liên tiếp.
```

---

## TẦNG 7 — HỘP CÔNG CỤ TỪ VỰNG

### 7.1 Cụm từ cố định (giữ từ v4 + bổ sung)

```
Phương pháp:
  "kết hợp với"
  "được xây dựng dựa trên"
  "tính bổ sung của nhiều loại bằng chứng nghiên cứu khác nhau
   soi chiếu lẫn nhau trên các lĩnh vực tìm hiểu"
  "tìm hiểu một cách mở, với sự linh hoạt cho các câu hỏi thăm dò"
  "phân tích mã hóa chủ đề định tính bằng phần mềm NVivo"
  "thường được sử dụng không phù hợp với"

Kết quả:
  "rơi vào phạm vi từ [mức] đến [mức]"
  "sự hiện diện của"
  "ít được quan sát thấy hơn"
  "bất kỳ mức độ triệu chứng nào"
  "với chênh lệch trung vị là [X] điểm"

Nhận xét:
  "không nằm trong các mục sàng lọc nhưng được mô tả chi tiết bởi"
  "được [đối tượng] trải nghiệm như những rào cản đối với"
  "liên quan đến trải nghiệm [vấn đề] của [đối tượng]"

Kết luận:
  "có ý nghĩa lâm sàng tiềm năng cho việc"
  "tình trạng dưới ngưỡng hội chứng"
  "mẫu phi lâm sàng"
  "đã có thể quan sát được trong nghiên cứu này"
  "là thiết yếu để tạo điều kiện cho sự tham gia giáo dục"
  "dù ở mức nhẹ hay nặng"
```

### 7.2 Kho danh hóa — Nominalization Inventory (MỚI v5)

> CTH dùng danh hóa rất CAO (15 lần / 1.156 từ = ~1,3%).
> Danh hóa = biến động từ/tính từ → danh từ bằng "sự/tính/việc".

```
"sự hiện diện" (← hiện diện)         — dùng ở: kết quả, kết luận
"sự chênh lệch" (← chênh lệch)      — dùng ở: so sánh giới tính
"sự lo lắng" (← lo lắng)             — dùng ở: nhận xét
"sự gia tăng" (← gia tăng)           — dùng ở: nhận xét
"sự tham gia" (← tham gia)           — dùng ở: nhận xét, kết luận
"sự linh hoạt" (← linh hoạt)         — dùng ở: phương pháp
"tính bổ sung" (← bổ sung)           — dùng ở: phương pháp
"việc nhận diện" (← nhận diện)        — dùng ở: kết luận

QUY TẮC:
  → Ưu tiên danh hóa khi có thể, đặc biệt ở:
    - Chủ ngữ câu: "Sự khác biệt giới tính..." thay vì "Giới tính khác biệt..."
    - Sau "cho thấy": "cho thấy sự hiện diện của..." thay vì "cho thấy X hiện diện"
  → Danh hóa tạo giọng trừu tượng, khoa học, phi cá nhân
  → Mục tiêu: ~10-15 danh hóa / 1.000 từ
```

### 7.3 Liên kết từ vựng — Lexical Cohesion Strategy (MỚI v5)

> CTH ưu tiên LẶP LẠI CHÍNH XÁC thuật ngữ hơn là dùng từ đồng nghĩa.
> Đây là chiến lược phổ biến trong văn hàn lâm: lặp = chính xác.

```
LẶP CHÍNH XÁC (ưu tiên):
  "trầm cảm" luôn là "trầm cảm" — không thay bằng "buồn bã", "suy sụp"
  "thanh thiếu niên" luôn là "thanh thiếu niên" — không thay bằng "trẻ", "giới trẻ"
  "sàng lọc" luôn là "sàng lọc" — không thay bằng "kiểm tra", "đánh giá"

THAY THẾ chỉ khi CẦN:
  "học sinh" ↔ "người tham gia nghiên cứu" ↔ "thanh thiếu niên" — tùy ngữ cảnh
  (học sinh = ở trường, người tham gia = trong nghiên cứu, thanh thiếu niên = nhóm tuổi)

QUY TẮC: Khi đã dùng một thuật ngữ chuyên ngành, KHÔNG thay đổi trong toàn bài.
```

---

## TẦNG 8 — HỆ THỐNG TYPOGRAPHIC (giữ từ v4)

```
« »    Tên công trình. Bên TRONG: `:` KHÔNG cách trước.
""     Thuật ngữ phân loại: mức "nhẹ"
:      TRONG thân bài nhãn: CÓ cách trước → "Nhãn : nội dung"
       TRONG tên công trình: KHÔNG cách
;      TRONG chuỗi trích dẫn: KHÔNG cách
,      Thập phân: 22,6%. KHÔNG Oxford: "A, B, C và D"
*...*  Nhãn nhận xét: *Nhãn.* Nội dung
Nhãn kết quả: Nhãn : nội dung (hai chấm, inline)
Trích dẫn ngoặc: (Tác giả và cộng sự, năm)
Trích dẫn trong câu: Tác giả (tên đầy đủ) và cs. (năm)
```

---

## TẦNG 9 — CHỈ SỐ ĐỊNH LƯỢNG MỤC TIÊU (MỚI v5)

> Đo từ bản gốc. Dùng để kiểm tra bài viết bắt chước có "đúng tầm" không.
> Nếu chỉ số lệch quá xa, phong cách đang trôi.

```
╔══════════════════════════════════════════════╦═══════════════╗
║ Chỉ số                                       ║ Giá trị CTH   ║
╠══════════════════════════════════════════════╬═══════════════╣
║ Độ dài câu trung bình                        ║ ~30 từ/câu    ║
║ Độ dài câu trung vị                          ║ ~27 từ/câu    ║
║ Độ lệch chuẩn độ dài câu                     ║ ~17 (CAO)     ║
║ Phạm vi độ dài câu                           ║ 3-72 từ       ║
║ Số câu trung bình / đoạn (nội dung)          ║ 1-4 câu       ║
║ Tỷ lệ bị động "được + V"                     ║ ~10% mệnh đề  ║
║ Tỷ lệ hedge:booster                          ║ 11:1          ║
║ Mật độ danh hóa (sự/tính/việc)               ║ ~13/1000 từ   ║
║ "của chúng tôi" (sở hữu) / bài              ║ 3-5 lần       ║
║ "Chúng tôi" (tác tử) / bài                  ║ 1-2 lần       ║
║ Trích dẫn non-integral                       ║ >80%          ║
║ Mật độ từ vựng (lexical density)             ║ ~80%          ║
╚══════════════════════════════════════════════╩═══════════════╝

CÁCH DÙNG:
  Sau khi viết xong, đếm các chỉ số trên.
  Nếu lệch >25% so với giá trị CTH → xem xét chỉnh lại.
  Đặc biệt chú ý: hedge:booster — nếu dùng quá nhiều booster, giọng sẽ sai.
```

---

## TẦNG 10 — CÂU HỎI CHẨN ĐOÁN (cập nhật từ v4)

```
=== TRƯỚC khi viết mỗi đoạn ===
□ Bước tu từ nào? (Tầng 2)
□ Lớp thông tin nào? (Tầng 1)
□ Mức thu phóng? Có phá nhịp? (Tầng 3A)
□ Tiến trình chủ đề gì? (Tầng 3C: tuyến tính / dẫn xuất / hằng số / song song)

=== TRƯỚC khi viết mỗi câu ===
□ Kiến trúc câu nào? (Tầng 4A: A/B/C/D)
□ Ai/cái gì làm tác tử? (Tầng 4B: sở hữu / vô tri / bị động / danh hóa?)
□ Mở đầu câu có khác câu trước? (Tầng 6.5)
□ Có cụm từ cố định phù hợp? (Tầng 7.1)
□ Có thể danh hóa không? (Tầng 7.2)
□ Cần hedge hay booster? (Tầng 5B: nhớ tỉ lệ 11:1!)

=== SAU khi viết xong bài ===
□ Nhịp độ dài câu có dạng sóng? (Tầng 3B)
□ Hedge:booster có gần 11:1? (Tầng 5B)
□ "của chúng tôi" 3-5 lần, "Chúng tôi" tác tử ≤2 lần? (Tầng 4B)
□ Danh hóa ~10-15 / 1.000 từ? (Tầng 7.2)
□ Mở đầu câu đa dạng? (Tầng 6.5)
□ Thuật ngữ lặp chính xác, không đồng nghĩa hóa? (Tầng 7.3)
□ Mỗi đoạn NX: mở đầu xoay vòng + cầu nối cuối? (Tầng 6.2, 6.3)
□ KL dùng Kiến trúc B? (Tầng 4A)
```

---

## PHỤ LỤC — TIẾN TRÌNH CẢI THIỆN

| Phiên bản | Điểm | Chiều phong cách đã nắm |
|---|---|---|
| v1 | — | Quy tắc cứng cơ bản |
| v2 | — | + kỹ thuật đoạn, câu linh hoạt |
| v3 | 9.1 | + sửa 11 lỗi format, phả hệ, hệ thống nhãn |
| v4 | 9.76 | + 8 tầng tư duy, 4 kiến trúc câu, nhịp thu phóng |
| v5 | ? | + metadiscourse, thematic progression, agency, |
|    |   |   evidentiality, discourse prosody, quantitative |
|    |   |   baselines, nominalization, lexical cohesion, |
|    |   |   sentence opener variety |

Nguồn học thuật tham khảo:
  - Hyland (2005): Metadiscourse — hedge/booster/attitude/engagement/self-mention
  - Halliday & Hasan: Cohesion — thematic progression, lexical cohesion
  - Swales (1990): CARS model — rhetorical moves in research articles
  - Martin & White (2005): Appraisal Theory — evaluative stance
  - Biber (1988): Multi-dimensional analysis — register variation
  - StyloAI (2024): 31 stylometric features for authorship attribution
