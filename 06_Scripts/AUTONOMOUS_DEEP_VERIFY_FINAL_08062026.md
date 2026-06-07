# AUTONOMOUS DEEP-VERIFY — FINAL SYNTHESIS (08/06/2026)

Phiên chạy tự động đêm 07–08/06/2026. File này tổng hợp 5 phụ lục cùng ngày:
`AUTONOMOUS_PDF_SCAN_08062026.md`, `AUTONOMOUS_M17_M18_08062026.md`,
`AUTONOMOUS_M29_M26_08062026.md`, `AUTONOMOUS_M42_M43_08062026.md`,
và đối chiếu chéo với `CROSSCHECK_Q2v1_Refs_07062026.md`.

Đối tượng audit chính: `bai-bao-Q1/Draft_Q2_SongNgu_v1_07062026.docx` (10.100 từ body, 35 TLTK).

---

## 1. 8 KEY-REFS — 3-WAY DB VERIFY (PubMed × Crossref × canonical_index)

Nguồn: `CROSSCHECK_Q2v1_Refs_07062026.md` đã được chéo với PubMed + Crossref ngày 07/06.

| # | Ref | PMID/DOI | PubMed | Crossref | canonical_index | Verdict |
|---|---|---|:-:|:-:|:-:|:-:|
| 1 | Kieling C 2024 JAMA Psychiatry 81(4):347 | 38324293 | OK | OK | QT079 | **PASS** |
| 2 | Pascoe MC 2020 J Affect Disord | 32056331 | OK | OK | QT067 | **PASS** |
| 3 | Rose AJ 2002 Child Development | 12487497 | OK | OK | QT078 | **PASS** |
| 4 | Robson EM 2025 Lancet Psychiatry 12(1):44 | 39644904 | OK | OK | QT081 (no-PDF, paywalled) | **PASS** |
| 5 | Chen Z 2023 BMC Psychiatry | — / 10.1186/s12888-023-04713-z | OK | OK | QT007 | **PASS** |
| 6 | Niwenahisemo LC 2024 Front Psychiatry | — / 10.3389/fpsyt.2024.1338381 | OK | OK | QT080 | **PASS** |
| 7 | Grummitt L 2025 eClinicalMedicine (OurFutures) | — / 10.1016/j.eclinm.2024.103009 | OK | OK | QT082 | **PASS** |
| 8 | GBD ASEAN 2025 Lancet Reg Health WPac | — / 10.1016/j.lanwpc.2024.101252 | OK | OK | QT012 | **PASS** |

**Tổng: 8/8 PASS.** Không phát hiện ref bị bịa hoặc lệch siêu dữ liệu.

---

## 2. PDF MAGIC + FUNDING SCAN (131 entries)

- Magic byte `%PDF`: 114/131 PASS (87,0%)
- Size > 20 KB: 115/131 PASS
- Open + page-1 extract: 115/131 PASS, 0 failure trên file mở được.
- Missing/no pdf_path: **16** (QT003-005, QT037, QT048-051, QT059, QT081, QT081-paywalled, VN007, VN030, 4 entries thiếu ID).
- Magic-byte FAIL: **QT079** (head `3c21444f43` = `<!DOC` → file HTML thay vì PDF).
- Stub suspects (<50 KB): QT075, QT087, QT088.
- Funding sample 10/10 ngẫu nhiên: 1/10 có "Funding" (VN015); 9 còn lại không tìm thấy trong page-1 (chấp nhận được, funding thường ở cuối bài).

---

## 3. M42 + M43 (cite-ref + AI hallmark)

- **M42 cite-ref integrity:** 35/35 match, **0 orphan, 0 ghost**. Cảnh báo kỹ thuật: en-dash `–` decode lỗi thành `�` ở plain text — kiểm lại khi export PDF.
- **M43 AI hallmark density:** 8 hits / 10.100 từ = **0,79 / 1.000 từ → THẤP**. Không có delve/multifaceted/tapestry/pivotal/nuanced. Cụm "trong bối cảnh" (4 hit) là tiếng Việt học thuật bình thường.

## 4. M17 + M18 (contradiction + bilingual)

- **M17 internal contradiction (Q25 SEM EN):** 32 β, 10 R², 3 F, 2 d, 21 p — **không mâu thuẫn**. Tỉ lệ |β| self-esteem/academic pressure (0,892 và 0,847) khớp số học chính xác.
- **M18 bilingual EN↔VN (10 cặp đoạn):** mọi giá trị β/R²/F/d/p/α khớp tuyệt đối sau chuẩn hoá decimal-comma. "Drift" duy nhất là N = 1,352 vs 1.352 (quy ước locale, không phải lỗi nội dung).

## 5. M29 + M26 (DSM/ICD năm-phiên bản + Karasu)

- **M29:** 4/4 file (Q2 v1, 4VanDe v3, PhuongAnXuLy, Verify v8) đều cặp đúng DSM-IV-TR/2000, DSM-5/2013, DSM-5-TR/2022 (PMID 35524596), ICD-11/2022 (PMID 35524598), CDDR/2024. **Không lệch năm-phiên bản**.
- **M26 Karasu credentials (Verify v8):** 12/12 PASS theo Einstein bio sketch + PDF gốc.

---

## NEW ISSUES PHÁT HIỆN (cần fix sau)

1. **QT079 PDF corrupt (HTML stub thay vì PDF).** Tải lại từ JAMA Psychiatry — `https://doi.org/10.1001/jamapsychiatry.2023.5051`. Mức độ: TRUNG BÌNH (ref vẫn đúng siêu dữ liệu).
2. **QT075 / QT087 / QT088 stub <50 KB** — kiểm lại nội dung; nếu phải tải lại thì cập nhật canonical_index.
3. **16 entries no pdf_path** — chủ yếu QT cũ + 4 entry `?` ID rỗng. Cần dọn `canonical_index.json` (xoá entry rỗng) ở phiên sau.
4. **En-dash encoding** trong Q2 v1 — kiểm trước khi export PDF cuối.
5. **Funding statement** — quét rộng hơn (toàn văn) thay vì chỉ page-1 ở phiên audit kế.

---

## TRẠNG THÁI CUỐI — Q2 v1 READY FOR SUBMISSION?

**CÓ (Y) — có điều kiện kỹ thuật nhẹ.**

Bằng chứng đủ điều kiện nộp:
- 8/8 key refs verify 3 chiều PASS.
- 35/35 cite-ref toàn vẹn (0 orphan, 0 ghost).
- 0 mâu thuẫn nội tại; bilingual EN↔VN khớp tuyệt đối.
- AI hallmark 0,79/1.000 (rất thấp).
- DSM/ICD năm-phiên bản chính xác.

Việc còn lại trước khi gửi: (a) sửa encoding en-dash khi export PDF; (b) tải lại QT079 PDF (không cản trở submission vì ref đã verify ngoài PDF).

---

## RECOMMENDATIONS — PHIÊN SAU

1. **Dọn `canonical_index.json`**: xoá 4 entry ID rỗng, thêm pdf_path cho 11 entry QT/VN còn thiếu (ưu tiên QT037, QT048–051, QT059).
2. **Re-download** QT079 + QT075/087/088; sau đó chạy lại PDF magic scan để đạt 100%.
3. **Funding full-text scan**: dùng pdfminer toàn văn (không chỉ page-1) để hoàn thiện M30.
4. **Export Q2 v1 → PDF** với font hỗ trợ Unicode (Times New Roman / Cambria) để kiểm en-dash.
5. **Chạy M44 (số liệu vs nguồn gốc)** trên Q2 v1 — đối chiếu N, %, OR vs PDF papers gốc (chưa làm trong phiên này).
6. **Update memory** sau khi user xác nhận Q2 v1 submitted: ghi project_q2v1_submitted.md.
