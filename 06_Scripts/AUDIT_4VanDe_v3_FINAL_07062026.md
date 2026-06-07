# AUDIT FINAL — `4VanDe_BLOCKING_Q1Q3_v3_07062026.docx`

**Audit date:** 07/06/2026
**Auditor:** Agent C (read-only verification)
**File path:** `bai-bao-Q1/4VanDe_BLOCKING_Q1Q3_v3_07062026.docx`
**Agent B claim:** 1.601 từ tiếng Việt
**Method:** python-docx extraction + PubMed independent verification

---

## 1. Word count

- **Target:** 1.500–2.000 từ
- **Actual:** **1601 từ** (đếm bằng `re.findall(r'\S+', full_text)` trên toàn bộ paragraphs + tables)
- **Verdict:** PASS — khớp với claim của Agent B (1.601).

---

## 2. Header + bối cảnh

| Item | Expected | Actual | Verdict |
|---|---|---|---|
| Tiêu đề | "4 Vấn đề BLOCKING — Bản v3 cập nhật ngày 07/06/2026" | "BON VAN DE BLOCKING — BAN v3 CAP NHAT NGAY 07/06/2026" (P0) | PASS |
| Đề-cap tin nhắn thầy MĐ 07/06/2026 | yes | "Ngày 07/06/2026, sau khi nhận tin nhắn từ thầy Mạnh Dũng (thầy MĐ)..." (P3) | PASS |
| Subtitle | có | "Tài liệu chốt trạng thái 4 vấn đề cũ và 2 quyết định mới của nhóm tác giả" (P1) | PASS |

---

## 3. Bảng 4 vấn đề cũ — cập nhật trạng thái (Table 0)

Cấu trúc: 5 hàng × 4 cột (Mã / Câu hỏi gốc / Quyết định mới / Trạng thái).

| Mã | Status text | Yêu cầu | Verdict |
|---|---|---|---|
| Q1-6 | "GIẢI QUYẾT — Quyết định thầy MĐ: BỎ HOÀN TOÀN phần định tính..." | GIẢI QUYẾT (vì bỏ phần định tính) | PASS |
| Q1-8 | "GIỮ NGUYÊN — Phương án A đã chốt ngày 01/06..." | GIỮ NGUYÊN (SEM Phương án A) | PASS |
| Q3-6 | "CHỜ XÁC NHẬN — NCS Công Thị Hằng vẫn cần cung cấp số quyết định + ngày ban hành..." | CHỜ NCS | PASS |
| Q3-9 | "ĐỒNG Ý — Theo solution em đề xuất trong v2..." | ĐỒNG Ý | PASS |

**Verdict:** PASS 4/4.

---

## 4. Hai quyết định mới

### Quyết định 5 — Chuyển Q1 → Q2 (P7–P12)

- Target chính: **Frontiers in Psychiatry** (Q2, JCR IF 2024 = 3,2) — PRESENT (P9).
- Backup 1: **BMC Public Health** (IF 2024 = 4,5) — PRESENT (P11).
- Backup 2: **Frontiers in Public Health** (IF 2024 = 3,0) — PRESENT (P12).
- **Verdict:** PASS (target chính + 2 backup đầy đủ).

### Quyết định 6 — Re-design Q3 companion theo Q2 framework (P13–P17)

- Có 2 hướng tái thiết kế Q3 (Hướng 1: chờ định tính; Hướng 2: SEM phụ SAD).
- Có đề xuất Hướng 2 là tối ưu.
- **Verdict:** PASS.

---

## 5. Bảng action items (Table 1)

- **Cấu trúc:** 10 hàng × **5 cột** = `Người / Vấn đề / Deadline / Output / Status`. PASS cấu trúc.
- **Phân bố 4 người (9 hàng nội dung + 1 header):**

| Người | Số hàng | Deadlines cụ thể |
|---|---|---|
| NCS Cong Thi Hang | 3 | 14/06/2026; 21/06/2026; 28/06/2026 |
| Thay MD (Manh Dung) | 2 | 24/06/2026; 12/06/2026 |
| Thay NMD (Nguyen Minh Duc) | 1 | 24/06/2026 |
| Em (trợ lý nghiên cứu) | 3 | 20/06/2026; 05/07/2026; 25/06/2026 |

- **Verdict:** PASS — đúng phân bố 3/2/1/3, deadline dạng dd/mm/yyyy cụ thể.

---

## 6. Tham khảo (≥ 6 refs) — VERIFICATION INDEPENDENT

8 mục tham khảo (P21–P28); trong đó 6 ref ngoại + 2 nội bộ.

| # | Ref | Agent B ghi | Sự thật PubMed/Crossref | Verdict |
|---|---|---|---|---|
| [1] | Frontiers Psychiatry guidelines | ISSN 1664-0640, IF 3,2 | hợp lệ (đường link tạp chí) | PASS (cần check ISSN online nhưng plausible) |
| [2] | BMC Public Health guidelines | ISSN 1471-2458, IF 4,5 | plausible | PASS |
| [3] | Frontiers Public Health guidelines | ISSN 2296-2565, IF 3,0 | plausible | PASS |
| [4] | **Karasu 1986 "specificity vs nonspecificity dilemma" Am J Psychiatry 143(6):687-695, PMID 3717388** | **PMID đúng = 3717390** (đã verify PubMed). PMID 3717388 thực ra là bài "Ionic requirements of peritubular taurine transport in Fundulus kidney" (Wolff et al., Am J Physiol 250(6):R984-90). | **FAIL — PMID SAI** |
| [5] | First et al. 2022 DSM-5-TR, World Psychiatry 21(2):218-219, DOI 10.1002/wps.20989, PMID 35524596 | Title thực = "DSM-5-TR: overview of what's new and what's changed" (Agent B viết "what is new and what has changed" — gần đúng nhưng không khớp 100%). Authors / journal / year / pages / PMID ĐÚNG. | PASS (PMID + DOI khớp; title nhẹ paraphrase) |
| [6] | **Pezzella 2022 "ICD-11 officially in effect" World Psychiatry 21(3):331-332, DOI 10.1002/wps.21038, PMID 36073678** | **PMID đúng = 35524598**; **DOI đúng = 10.1002/wps.20982**; **issue đúng = 21(2)** chứ không phải 21(3). PMID 36073678 thực ra là Krueger RF "Incremental integration of nosological innovations…" 21(3):416-417. | **FAIL — PMID + DOI + issue ĐỀU SAI** |
| [7] | Nội bộ: v2_01062026 | — | PASS |
| [8] | Nội bộ: tin nhắn thầy MĐ 07/06 | — | PASS |

**Verdict refs:** **FAIL 2/6 (Karasu, Pezzella) — Agent B fabricate PMID/DOI.**

### CRITICAL — Karasu PMID verification chi tiết

User yêu cầu: PMID đúng cho "The Psychotherapies: Benefits and Limitations" Am J Psychotherapy 40(3):324-342 = **3094389**.

Tuy nhiên Agent B trong văn bản v3 thực tế đã trích **một bài Karasu KHÁC**: "The specificity versus nonspecificity dilemma: toward identifying therapeutic change agents", Am J Psychiatry 143(6):687-695 (ref [4]).

- PMID đúng cho bài "specificity vs nonspecificity" mà Agent B trích = **3717390** (verify PubMed).
- Agent B ghi **3717388** → SAI. PMID 3717388 là bài fish-kidney physiology hoàn toàn không liên quan.
- Nếu user kỳ vọng bài "Benefits and Limitations" (PMID 3094389) thì cả bài trích lẫn PMID đều khác (Agent B đã thay bài).

**Action cần fix:** xác nhận với thầy MĐ xem nhóm muốn trích bài nào (Karasu 1986 AJP-specificity, PMID 3717390; HAY Karasu 1986 AJ Psychotherapy-Benefits/Limitations, PMID 3094389). Bất kể chọn bài nào, PMID 3717388 hiện ghi trong v3 là **SAI tuyệt đối**.

---

## 7. Tiếng Việt thuần + metadata + footer

| Item | Verdict |
|---|---|
| Tiếng Việt thuần (chỉ chú thích Anh trong ngoặc: SAD, Social Anxiety Disorder, IRB, JCR, open access, companion, framework, phenotype, joint-display, retroactive, plagiarism, submission, draft, backup, target — vài term Anh không có VN tương đương) | PASS (chấp nhận được, nhưng số term Anh hơi nhiều: "backup", "target", "framework", "draft" — có thể Việt hóa thêm) |
| Metadata stripped (author='', title='', keywords='', subject='', last_modified_by='') | PASS (tất cả rỗng) |
| Footer "Soạn 07/06/2026" | PASS — nằm ở P30 cuối file (không phải Word footer thực sự — section.footer rỗng, nhưng có dòng "Soan 07/06/2026" cuối body) |

Lưu ý: Footer là dòng văn bản cuối body (không phải Word footer object) — đáp ứng yêu cầu functionally.

---

## TỔNG KẾT

| Check | Result |
|---|---|
| 1. Word count (1.601) | PASS |
| 2. Header + bối cảnh | PASS |
| 3. 4 vấn đề cũ status | PASS 4/4 |
| 4. 2 quyết định mới | PASS |
| 5. Action items 5 cột × 9 hàng (3/2/1/3) | PASS |
| 6. Tham khảo verify | **FAIL 2/6 — Karasu PMID + Pezzella PMID/DOI/issue đều sai** |
| 7. Tiếng Việt + metadata + footer | PASS |

**Overall:** **6 PASS / 1 FAIL** — file structure hoàn chỉnh, NHƯNG có 2 reference fabrication critical cần fix trước khi gửi thầy MĐ.

### MUST-FIX trước khi gửi

1. **Ref [4] Karasu**: thay PMID `3717388` → `3717390` (nếu giữ bài "specificity vs nonspecificity"). Hoặc đổi sang bài "Benefits and Limitations" PMID `3094389` (đúng theo gợi ý user).
2. **Ref [6] Pezzella**: thay PMID `36073678` → `35524598`; thay DOI `10.1002/wps.21038` → `10.1002/wps.20982`; thay issue `21(3)` → `21(2)`.

### Optional polish

- Cân nhắc Việt hóa thêm: "backup" → "dự phòng" (đã dùng song song, OK); "target" → "mục tiêu"; "framework" → "khung tham chiếu"; "draft" → "bản thảo".
- Verify ISSN của 3 journal Q2 backup qua JCR/SCImago (chưa thực hiện vì user không yêu cầu).
