# UPDATE TEMPLATES — Yêu cầu Tham khảo bắt buộc

**Ngày:** 07/06/2026
**Người thực hiện:** Trợ lý nghiên cứu (Claude)
**Bối cảnh:** Audit 07/06/2026 phát hiện 89,8% bản tóm tắt cũ thiếu danh mục Tham khảo — nguyên nhân gốc là template gốc chưa quy định. Cần khoá lỗ hổng này cho mọi summary tương lai.

---

## 1. Files đã cập nhật

| # | Đường dẫn | Loại thay đổi |
|---|-----------|---------------|
| 1 | `TEMPLATE_thiet_ke_bai_test_kiem_tra_tri_thuc.md` | Thêm Mục 7 + cập nhật header |
| 2 | `PROMPT_MAU_PHAN_TICH_BAI_BAO.md` | Thêm Mục 7 + thêm tiểu mục VI vào cấu trúc báo cáo + cập nhật header |

---

## 2. Tóm tắt thay đổi

### 2.1. Header cả hai file
Thêm dòng ghi chú "Cập nhật 07/06/2026 — thêm yêu cầu Tham khảo bắt buộc" với lý do (89,8% tóm tắt cũ thiếu refs).

### 2.2. Mục 7 mới — "MỤC THAM KHẢO BẮT BUỘC"
Nội dung giống nhau ở cả hai file, gồm:

- **HARD GATE rule** (in đậm): "KHÔNG nộp summary thiếu Tham khảo. Đây là yêu cầu bắt buộc cho mọi summary từ 07/06/2026 trở đi."
- **7.1. Vị trí và tên mục** — cuối file, tiêu đề chuẩn `## Tham khảo / References`, đánh số `[1]`, `[2]`, ...
- **7.2. Định dạng entry (APA-7)** — bắt buộc 6 trường: tác giả / năm / tiêu đề / *tạp chí* / vol(iss) / trang.
  - DOI bắt buộc nếu có
  - PMID bắt buộc nếu paper trên PubMed
  - PDF path tương đối (`02_Papers-goc/...`) bắt buộc nếu paper trong canonical_index
  - URL cho nguồn không-DOI (WHO, preprint, web)
- **7.3. Ví dụ entry chuẩn** — 2 entry mẫu (Pascoe 2020 có DOI+PMID+PDF; WHO 2022 chỉ URL+PDF).
- **7.4. Quy trình verify** — tra `canonical_index.json` trước; nếu chưa có thì verify từng trường qua PubMed/Crossref/JCR.
- **7.5. Liên kết tới memory nội bộ** — `feedback_doc_phai_co_reference.md`, `feedback_verify_tung_entry_truoc_khi_gui.md`, `feedback_quy_trinh_fact_check.md`, `canonical_index.json`.
- **7.6. Checklist 8 mục** trước khi nộp.

### 2.3. Bổ sung riêng cho `PROMPT_MAU_PHAN_TICH_BAI_BAO.md`
Thêm tiểu mục **VI. THAM KHẢO / REFERENCES** vào "Cấu trúc báo cáo chuẩn" (Mục 5) — đảm bảo người dùng template thấy yêu cầu refs ngay khi soạn outline.

---

## 3. Khuyến nghị — các template/prompt liên quan cần update tương tự

Nên rà soát và bổ sung cùng yêu cầu Tham khảo vào:

1. **`Tom-tat-tung-bai/`** — nếu có file template/README mô tả cách viết tóm tắt từng bài, cần đồng bộ.
2. **`00_Meta/`** — kiểm tra các file hướng dẫn quy trình (workflow, SOP) xem có cần thêm dòng "tóm tắt phải có Tham khảo".
3. **`06_Scripts/_extract_workspace/WHO2022_PROGRESS.md`** — nếu có template cho từng Doc dịch WHO, đảm bảo Doc 3-9 sắp tới có Tham khảo.
4. **Prompt sinh tóm tắt tự động** (nếu có script Python/notebook tạo tóm tắt từ RAG) — sửa prompt để bắt buộc sinh kèm mục Tham khảo.
5. **`bai-bao-Q1/`, `paper-may/`, `madam Thao/`** — kiểm tra có template riêng không.

**Bỏ qua theo yêu cầu:** `bai-bao-khgdvn/` (như chỉ thị).

---

## 4. Tác động dự kiến

- **Summary mới (từ 07/06/2026):** 100% có Tham khảo (so với 10,2% trước đó).
- **Summary cũ:** không bị ép sửa ngược; sẽ được bổ sung dần theo audit `AUDIT_TomTat_MissingRefs_07062026.md`.
- **Khả năng fact-check:** tăng đáng kể nhờ có DOI/PMID/PDF path để truy nguyên.

---

## 5. Tham khảo / References

[1] Memory nội bộ — `feedback_doc_phai_co_reference.md` (quy định gốc về Tham khảo bắt buộc với DOI/PMID/URL kiểm chứng).
[2] Memory nội bộ — `feedback_verify_tung_entry_truoc_khi_gui.md` (HARD GATE verify từng entry qua PubMed/Crossref/JCR).
[3] Memory nội bộ — `feedback_quy_trinh_fact_check.md` (17 loại fact-check).
[4] Audit 07/06/2026 — `06_Scripts/AUDIT_TomTat_MissingRefs_07062026.md` (phát hiện 89,8% summaries thiếu refs).
[5] American Psychological Association. (2020). *Publication manual of the American Psychological Association* (7th ed.). APA. https://doi.org/10.1037/0000165-000

---

*Tạo 07/06/2026 — báo cáo tự động sau khi cập nhật hai template chính.*
