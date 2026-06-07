# TEMPLATE — THIẾT KẾ BÀI TEST KIỂM TRA TRI THỨC

> File mẫu, có thể sao chép sang **dự án khác**. Đúc rút từ dự án nghiên cứu
> "Lo âu học sinh" (RAG + KG + tóm tắt + metadata xây trên 90+ tài liệu nguồn),
> tháng 4–5/2026, sau khi phát hiện hệ thống bị nhiễm dữ liệu bịa/sai.
>
> **Phạm vi áp dụng**: bất kỳ dự án nào có một *cơ sở tri thức* (knowledge base)
> được xây dựng/trích xuất từ tài liệu nguồn — RAG index, đồ thị tri thức (KG),
> bản tóm tắt, metadata, chatbot. Mục tiêu: phát hiện sai lệch giữa cơ sở tri
> thức và **nguồn gốc** trước khi nó lan vào sản phẩm cuối.
>
> **Cập nhật 07/06/2026** — thêm yêu cầu Tham khảo bắt buộc (Mục 7) sau khi
> audit phát hiện 89,8% bản tóm tắt cũ thiếu danh mục Tham khảo vì template
> gốc chưa quy định. Từ 07/06/2026, MỌI bản tóm tắt/test design phải có mục
> "Tham khảo / References" ở cuối file.

---

## 0. NGUYÊN TẮC GỐC

1. **Một cơ sở tri thức "nghe hợp lý" KHÔNG phải là đã đúng.** Nội dung do mô
   hình sinh/trích xuất luôn phải được kiểm chứng ngược về nguồn gốc.
2. **Test phải có ĐÁP ÁN CHUẨN độc lập** (gold answer) lấy từ nguồn gốc, không
   lấy từ chính cơ sở tri thức đang kiểm.
3. **Sai một chỗ → nghi sai nhiều chỗ.** Lỗi thường lan: cùng một bản trích xuất
   sai sinh ra lỗi đồng thời ở RAG, KG, metadata, chatbot.
4. **Test phải bắt được cả lỗi BỊA (fabrication), không chỉ lỗi thiếu.**

---

## 1. SÁU LOẠI BÀI TEST

| # | Loại test | Câu hỏi nó trả lời | Đối tượng |
|---|-----------|--------------------|-----------|
| 1 | **Truy xuất (Retrieval)** | Truy vấn X có trả về đúng tài liệu/đoạn không? | RAG index |
| 2 | **Trung thực nguồn (Source-fidelity audit)** | Mỗi mục tri thức có khớp tài liệu gốc không? | Tóm tắt / RAG / KG |
| 3 | **Nhất quán nội bộ** | Các con số/khẳng định trong KB có mâu thuẫn nhau không? | Toàn KB |
| 4 | **Đồ thị tri thức (KG)** | Node/cạnh/triple có đúng và đúng chủ thể không? | KG |
| 5 | **Lan truyền (Propagation)** | Sau khi sửa 1 sự kiện, mọi tầng dẫn xuất đã đồng bộ chưa? | Mọi tầng |
| 6 | **Hồi quy (Regression)** | Bản sửa có làm hỏng chỗ khác / còn sót không? | Toàn KB |

### 1.1. Test TRUY XUẤT (smoke test)
- Soạn bộ truy vấn mẫu; với mỗi truy vấn, ghi sẵn **tài liệu/đoạn đúng phải nằm
  trong top-k** (k = 3 hoặc 5).
- Chỉ số: tỷ lệ hit@k. Ngưỡng đề xuất: ≥ 95% hit@5 cho bộ smoke test.
- Bẫy: thêm vài truy vấn mà KB **không nên** trả lời tự tin (ngoài phạm vi) — xem
  hệ thống có "bịa" câu trả lời không.

### 1.2. Test TRUNG THỰC NGUỒN — quan trọng nhất
Với **từng mục** tri thức, kiểm 3 mức riêng biệt:
- **(a) Tồn tại** — tài liệu/sự kiện có thật không?
- **(b) Thư mục đúng** — tác giả, năm, nhan đề, nguồn, định danh (DOI/URL) đúng?
- **(c) Khẳng định khớp** — con số/phát hiện gán cho mục có đúng nguồn không?

→ **4 mức phán quyết (verdict)**: `ĐÚNG` / `SAI-THƯ-MỤC` / `LỆCH-KHẲNG-ĐỊNH` /
`KHÔNG-TỒN-TẠI (bịa)`.

Nguồn đối chiếu = **tài liệu gốc**, KHÔNG phải bản tóm tắt nội bộ.

### 1.3. Test NHẤT QUÁN NỘI BỘ
- Cùng một con số xuất hiện nhiều nơi có khớp nhau không?
- Tổng các thành phần = tổng đã ghi? (vd các % cộng lại ≈ 100%).
- Quan hệ toán học đúng? (vd R = 0,177 ⇒ R² ≈ 0,031).
- Thuật ngữ/viết tắt dùng nhất quán.

### 1.4. Test ĐỒ THỊ TRI THỨC (KG)
- Mỗi node có thật, gán đúng loại.
- Mỗi triple `(chủ thể, quan hệ, đối tượng)` đúng — **đặc biệt kiểm chủ thể**:
  fact tự trích dễ gán số liệu của tài liệu B cho tài liệu A khi A chỉ *nhắc tới* B.
- Không có node/cạnh mồ côi; không trùng lặp.

### 1.5. Test LAN TRUYỀN
Khi sửa một sự kiện, lập danh sách **mọi tầng dẫn xuất** và kiểm từng tầng:
RAG index (mọi bản sao!) → KG/triples → metadata/faceted browser → tên file →
tóm tắt/báo cáo dẫn xuất → chatbot. Sửa một bản chưa đủ.

### 1.6. Test HỒI QUY
- Sau mỗi đợt sửa: chạy lại test 1–5.
- Script sửa phải in **"số khớp / số không khớp"**; còn sót phải truy nguyên do.
- Đọc lại file đã sửa để xác nhận thay đổi thật sự xảy ra.

---

## 2. CÁCH THIẾT KẾ BỘ CÂU HỎI TEST

1. **Bộ câu hỏi vàng (gold set)**: mỗi câu gồm `truy vấn` + `đáp án chuẩn` +
   `nguồn của đáp án` + `độ khó`. Đáp án chuẩn lấy từ tài liệu gốc.
2. **Phủ (coverage)**: mỗi tài liệu/chủ đề trong KB có ≥ 1 câu; ưu tiên tài liệu
   trọng yếu nhiều câu hơn.
3. **Phân tầng độ khó**:
   - *Dễ* — tra cứu trực tiếp (1 tài liệu, 1 dữ kiện).
   - *Vừa* — tổng hợp 2–3 tài liệu.
   - *Khó* — suy luận liên tài liệu, so sánh, phát hiện mâu thuẫn.
4. **Câu bẫy (adversarial)** — bắt buộc có:
   - Câu hỏi **ngoài phạm vi** KB → hệ thống phải nói "không biết", không bịa.
   - Câu hỏi về **số liệu dễ bịa** (effect size, năm, cỡ mẫu, tên tác giả).
   - Câu hỏi mà đáp án **trái trực giác** → bắt lỗi "đảo ngược phát hiện".
5. **Câu kiểm chủ thể** — hỏi "con số này của tài liệu nào?" để bắt lỗi gán nhầm.

---

## 3. CHECKLIST LOẠI LỖI CẦN BẮT (đúc rút thực chiến)

| Loại lỗi | Mô tả | Cách test bắt |
|----------|-------|---------------|
| **Bịa hoàn toàn** | Tài liệu/số liệu không tồn tại | Test 1.2(a): tra nguồn độc lập |
| **Sai thư mục** | Tác giả/năm/nguồn/định danh sai | Test 1.2(b) |
| **Lệch khẳng định** | Tài liệu thật nhưng số liệu gán sai | Test 1.2(c) |
| **Gán nhầm chủ thể** | Số của tài liệu B bị gán cho A | Test 1.4 + câu kiểm chủ thể |
| **Sai công cụ/đơn vị** | Ghi sai tên thang đo/phương pháp | Đối chiếu mục phương pháp nguồn |
| **Đảo ngược phát hiện** | Ghi ngược thứ hạng/chiều kết quả | Câu bẫy trái trực giác |
| **Số nhóm con ↔ số tổng** | Lấy số một nhóm nhỏ làm số chung | Hỏi rõ "tổng thể hay nhóm con?" |
| **Nhầm trạng thái** | Bản tiền in (preprint) ghi là đã xuất bản | Kiểm định danh có "preprint" |
| **Lỗi lan truyền** | Đã sửa nguồn nhưng tầng dẫn xuất còn sai | Test 1.5 |
| **Tự tạo lỗi khi sửa** | Script sửa ghi đè nhầm (vd ô bảng gộp) | Test 1.6 + đọc lại file |

---

## 4. QUY TRÌNH CHẠY

1. **Nhiều vòng**: chạy test 1→6, sửa, rồi LẶP LẠI đến khi không còn lỗi.
2. **Chia agent song song** khi khối lượng lớn (mỗi agent một lô tài liệu) —
   vừa nhanh, vừa không mất cả lô nếu một agent lỗi. Mỗi agent phải **trả về
   đường dẫn nguồn** cho mọi khẳng định.
3. **Giữ mọi phiên bản cũ** để đối chứng; script chỉ đọc bản cũ, ghi bản mới.
4. **Ngưỡng đạt**: 0 lỗi bịa; 0 lệch-khẳng-định; hit@k ≥ ngưỡng; nhất quán 100%.

---

## 5. CHỈ SỐ ĐÁNH GIÁ (METRICS)

- **Retrieval**: hit@k, độ chính xác top-1.
- **Trung thực nguồn**: phân bố 4 verdict; **số mục bịa = 0 là bắt buộc**.
- **Nhất quán**: số mâu thuẫn nội bộ.
- **Độ phủ test**: % tài liệu có ≥ 1 câu kiểm.
- **Lan truyền**: số tầng đã đồng bộ / tổng số tầng.

---

## 6. PORT SANG DỰ ÁN KHÁC — phần CHUNG vs phần RIÊNG

**Dùng lại được nguyên (generic):**
- 6 loại test, 4 mức verdict, checklist loại lỗi, quy trình nhiều vòng.
- Cấu trúc bộ câu hỏi vàng, ý tưởng câu bẫy.

**Phải thay theo dự án (domain-specific):**
- Danh sách tài liệu nguồn và nơi lưu.
- Danh sách "tầng dẫn xuất" cụ thể (file index, KG, metadata… của dự án đó).
- Bộ câu hỏi vàng + đáp án chuẩn của lĩnh vực.
- Ngưỡng đạt (tuỳ mức độ rủi ro của dự án).

---

## PHỤ LỤC — Lược đồ gợi ý cho 1 câu hỏi vàng (JSON)

```json
{
  "id": "Q001",
  "query": "Truy vấn người dùng",
  "gold_answer": "Đáp án chuẩn lấy từ tài liệu gốc",
  "gold_source": "Định danh tài liệu + trang/đoạn",
  "expected_docs": ["DOC_A", "DOC_B"],
  "difficulty": "dễ | vừa | khó",
  "type": "tra-cứu | tổng-hợp | suy-luận | bẫy-ngoài-phạm-vi | kiểm-chủ-thể",
  "verdict": "ĐÚNG | SAI-THƯ-MỤC | LỆCH-KHẲNG-ĐỊNH | KHÔNG-TỒN-TẠI"
}
```

*Mỗi đợt audit: xuất một báo cáo gồm phân bố verdict + danh sách mục lỗi
(kèm trích dẫn nguyên văn + vị trí) — đó là đầu ra để sửa.*

---

## 7. MỤC THAM KHẢO BẮT BUỘC — yêu cầu chuẩn từ 07/06/2026

> **CỔNG CỨNG (HARD GATE):**
> **KHÔNG nộp summary / test design / báo cáo phân tích nếu thiếu mục Tham khảo.**
> **Đây là yêu cầu bắt buộc cho mọi summary từ 07/06/2026 trở đi.**
> Bối cảnh: audit 07/06/2026 phát hiện 89,8% bản tóm tắt cũ thiếu danh mục
> Tham khảo. Nguyên nhân gốc: template không quy định → người soạn bỏ qua →
> không truy vết được nguồn → fact-check khó. Quy định này khoá lỗ hổng đó.

### 7.1. Vị trí và tên mục
- Đặt ở **cuối file**, ngay trước mọi phụ lục/changelog.
- Tiêu đề chuẩn: `## Tham khảo / References`.
- Mỗi mục đánh số `[1]`, `[2]`, ... và được trích dẫn trong thân bài bằng số đó.

### 7.2. Yêu cầu định dạng từng entry (APA-7 + định danh số)
Mỗi entry **bắt buộc đủ** các thành phần sau (theo APA-7):

1. **Tác giả** — Họ, Chữ cái đầu tên (vd `Nguyen, M. H.`). Tối đa 20 tác giả;
   từ 21+ ghi tác giả đầu + `... ` + tác giả cuối.
2. **Năm xuất bản** — trong ngoặc, vd `(2024)`. Preprint: `(2024, March 15)`.
3. **Tiêu đề bài báo** — câu thường, kết bằng dấu chấm.
4. **Tên tạp chí** *(in nghiêng)* — đầy đủ, không viết tắt.
5. **Volume(Issue)** — vd `45(3)`. *Volume in nghiêng*, Issue trong ngoặc.
6. **Trang** — vd `123–145`. Bài online-only: ghi `Article e12345`.

**Định danh số (BẮT BUỘC nếu có):**
- **DOI** — dạng URL: `https://doi.org/10.xxxx/xxxxx`. **Bắt buộc nếu paper có DOI.**
- **PMID** — `PMID: 12345678`. **Bắt buộc nếu paper được index trên PubMed.**
- **URL** — chỉ dùng cho nguồn KHÔNG có DOI (preprint server, báo cáo WHO, web).

**Liên kết nội bộ (BẮT BUỘC nếu paper là từ canonical_index):**
- **PDF path** — đường dẫn TƯƠNG ĐỐI tới `02_Papers-goc/`, vd:
  `02_Papers-goc/Smith_2023_Anxiety_Adolescents.pdf`.
- Ưu tiên tra cứu trong `04_Co-so-du-lieu/canonical_index.json` (hoặc tương đương)
  để lấy DOI/PMID/PDF path đã verify, **không tự đoán**.

### 7.3. Ví dụ entry chuẩn

```
[1] Pascoe, M. C., Hetrick, S. E., & Parker, A. G. (2020). The impact of
    stress on students in secondary school and higher education.
    *International Journal of Adolescence and Youth*, *25*(1), 104–112.
    https://doi.org/10.1080/02673843.2019.1596823
    PMID: 31736680
    PDF: 02_Papers-goc/Pascoe_2020_Stress_Students.pdf

[2] World Health Organization. (2022). *World mental health report:
    Transforming mental health for all*. WHO.
    https://www.who.int/publications/i/item/9789240049338
    (Không có DOI/PMID — nguồn WHO; PDF: 02_Papers-goc/WHO_2022_WMHR.pdf)
```

### 7.4. Quy trình verify trước khi đưa vào Tham khảo
- Tra `04_Co-so-du-lieu/canonical_index.json` trước → copy entry đã verify.
- Nếu paper CHƯA có trong canonical_index: verify từng trường (PubMed +
  Crossref + JCR) **trước khi viết vào file** — xem memory
  `feedback_verify_tung_entry_truoc_khi_gui.md`.
- KHÔNG dựa trí nhớ hoặc bản tóm tắt nội bộ để tạo entry.

### 7.5. Liên kết tới memory & tài liệu nội bộ
- `feedback_doc_phai_co_reference.md` — quy định gốc: mọi doc trả lời phải có
  Tham khảo với DOI/PMID/URL kiểm chứng + truy vết RAG/KG/glossary nội bộ.
- `feedback_verify_tung_entry_truoc_khi_gui.md` — HARD GATE verify từng entry.
- `feedback_quy_trinh_fact_check.md` — 17 loại fact-check cho bài báo/luận án.
- `04_Co-so-du-lieu/canonical_index.json` — nguồn chuẩn để copy entry đã verify.

### 7.6. Checklist trước khi nộp summary
- [ ] Có mục `## Tham khảo / References` ở cuối file.
- [ ] Mỗi khẳng định/số liệu trong thân bài có chỉ số `[n]` trỏ về entry.
- [ ] Mỗi entry đủ 6 trường APA-7 (tác giả/năm/tiêu đề/tạp chí/vol(iss)/trang).
- [ ] DOI có mặt cho mọi paper có DOI.
- [ ] PMID có mặt cho mọi paper trên PubMed.
- [ ] PDF path tương đối có mặt cho mọi paper trong canonical_index.
- [ ] URL có mặt cho nguồn không-DOI (WHO, báo cáo, preprint không-DOI).
- [ ] Mỗi entry đã được verify (PubMed/Crossref/JCR HOẶC copy từ canonical_index).
