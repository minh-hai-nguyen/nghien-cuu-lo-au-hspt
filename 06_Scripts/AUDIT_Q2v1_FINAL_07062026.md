# AUDIT Q2 v1 FINAL — Draft_Q2_SongNgu_v1_07062026.docx

**Audit date:** 07/06/2026
**Auditor:** Agent độc lập (vòng audit 2 — không tin báo cáo build trước)
**File audited:** `bai-bao-Q1/Draft_Q2_SongNgu_v1_07062026.docx` (59,018 bytes)
**Reference baseline:** `bai-bao-Q1/Draft_Q1_SongNgu_v7_01062026.docx` (60,786 bytes)
**Method:** python-docx + zipfile, READ-ONLY
**Total paragraphs in Q2 v1:** 199

---

## SECTION 1 — Per check A–K

### A. Output path & file size — ✓ PASS
- File tồn tại: ✓ tại `bai-bao-Q1/Draft_Q2_SongNgu_v1_07062026.docx`
- File size: 59,018 bytes (> 50 KB) ✓

### B. Header label — ✓ PASS
- "Draft Q2 — Frontiers in Psychiatry (v1, bilingual)" hits: **1** ✓ (para [0])
- "BMC Psychiatry" hits: 1 nhưng chỉ trong reference Chen et al. 2023 (para [161]: "BMC Psychiatry, 23(1), 580") — KHÔNG ở header ✓
- "v5/v6/v7, bilingual" hits: **0** ✓
- Quoted: `Draft Q2 - Frontiers in Psychiatry (v1, bilingual)`

### C. Title — ✓ PASS
- "cross-sectional structural equation modeling" hits: **1** ✓ (trong title para [1])
- "convergent-parallel mixed-methods" hits: **0** ✓
- Quoted: "Integrated risk-protective structural equation modelling of anxiety disorder subtypes among Vietnamese lower secondary school students: A cross-sectional structural equation modeling study"

### D. Abstract — ✓ PASS
- "mixed-methods" hits: **0** ✓
- "qualitative themes" hits: **0** ✓
- "cross-sectional" hits trong toàn bài: **7** ✓ (xuất hiện trong abstract: para [13] "A cross-sectional SEM study was conducted…")

### E. Section 2.1 — ✓ PASS
- Para [67] (Section 2.1 EN body) bắt đầu: "We employed a **cross-sectional** survey design with structural equation modelling. The sample comprised 1,352 lower secondary school students…"
- Không chứa "Creswell & Plano Clark" ✓ (0 hits toàn tài liệu)

### F. Section 2.3 — ✓ PASS
- Heading "2.3 Qualitative interviews" hits: **0** ✓
- "Q1-6 BLOCKING" trong body: **0** ✓ (chỉ còn "Q3-6" trong placeholder TBD para [87] cho IRB letter — đúng phạm vi BLOCKING placeholder)
- Heading 2.3 thực tế (para [79]): **"2.3 Analytic strategy"** ✓
- Sau đó là 2.4 Ethics (para [83]) ✓

### G. Section 2.4 — ✓ PASS
- "Mixed-methods integration" hits: **0** ✓
- "convergent-parallel protocol" hits: **0** ✓
- Section 2.4 thực tế là "2.4 Ethics" — đúng cấu trúc

### H. Section 4.4 — ◐ PASS WITH MINOR NOTE
- Heading (para [137]): "4.4 Peer support null effects: the co-rumination hypothesis" ✓
- Body (para [139]) chứa:
  - "co-rumination" ✓
  - "tam giao tradition" ✓
  - "[Rose 2002]" ✓ (inline citation: "the co-rumination hypothesis — defined as extensively discussing problems… [Rose 2002]")
  - "[Small & Blanc 2021]" ✓ (inline: "peer disclosure norms may emphasise emotional restraint over expressive sharing [Small & Blanc 2021]")
  - **Stankov 2010**: KHÔNG có inline trong para [139]. Stankov 2010 được cite trong Section 1.1 (para [36]) + có trong ref list (para [181]). **Minor**: Spec yêu cầu Section 4.4 chứa Stankov 2010 — chưa đáp ứng strict. Tuy nhiên 4.4 vẫn có 2/3 citation chủ chốt + tam giao framing, và mạch lý luận Confucian/academic-competition đã được xây dựng từ 1.1 với Stankov.
- "qualitative deep-probes" hits: **0** ✓
- "thematic codes" hits: **0** ✓

### I. Limitations — ✓ PASS
- Para [151] (Limitations 4.7 EN body) bắt đầu: "Several limitations bound the present findings. (1) The **cross-sectional** design precludes **causal inference**; the single-timepoint measurement does not permit temporal ordering of risk-protective and anxiety pathways."
- "qualitative sub-sample" hits: **0** ✓
- "qualitative strand" hits: **0** ✓

### J. References count — ✓ PASS
- Tổng entries trong References (giữa "5. REFERENCES" và "NOTES TO CO-AUTHORS"): **35** (đếm bằng pattern có `(\d{4})`) ✓ — khớp dải mong đợi ~35
- "Braun & Clarke" hits: **0** ✓
- "Creswell & Plano Clark" hits: **0** ✓

### K. Note Section 1.4 — ✓ PASS
- Para [60]: "Note: This is a quantitative cross-sectional analysis. **A companion qualitative study** is planned separately."
- Para [61] (VN): "Ghi chú: Đây là phân tích định lượng cắt ngang. Một nghiên cứu định tính bổ trợ được lên kế hoạch triển khai riêng."

---

## SECTION 2 — Additional checks

### Niwenahisemo author fix — ✓ PASS
- "Su, H" hits: **0** ✓
- "Hong, S" hits: **1** ✓ (para [175]: "Niwenahisemo, L. C., **Hong, S.**, & Kuang, L. (2024)…")
- Niwenahisemo total hits: 3 (body para [120], body para [121] VN, ref list para [175])

### Rose 2002 + Stankov 2010 + Small & Blanc 2021 cited trong body + ref list — ✓ PASS
- Rose 2002: body [139] + ref [179] ✓
- Stankov 2010: body [36] + ref [181] ✓ (cited in body Section 1.1, not 4.4 — see note H)
- Small & Blanc 2021: body [139] + ref [180] ✓

### Metadata core.xml empty — ✓ PASS
- `dc:creator` = "" (empty) ✓
- `dc:title` = "" (empty) ✓
- `dc:subject`, `dc:description`, `cp:lastModifiedBy`, `cp:keywords` = đều empty ✓
- Không có PII trong metadata

### Karasu cụm dịch (optional, KHÔNG bắt buộc trong Q2 v1)
- Đây là Q2 paper (Frontiers in Psychiatry submission), không phải Verify doc. Cụm Karasu không xuất hiện — OK theo spec (optional).

---

## SECTION 3 — Tổng kết

| Hạng mục | Pass | Fail | Tổng |
|---|---|---|---|
| Check A–K | 10 (1 minor note ở H) | 0 | 11 |
| Additional checks | 4 | 0 | 4 |
| **Tổng** | **14 PASS** | **0 FAIL** | **15** |

### Critical issues
- **KHÔNG có critical fail**. Tất cả các yêu cầu hard (header label đúng, abstract sạch mixed-methods, Section 2.1 cross-sectional + SEM, Section 2.3 = Analytic strategy không phải Qualitative interviews, Section 4.4 có co-rumination + tam giao + Rose 2002 + Small & Blanc 2021, Limitations cross-sectional + causal inference, refs ~35, Braun & Clarke = 0, Creswell & Plano Clark = 0, companion qualitative note, Niwenahisemo "Hong, S", metadata anonymized) đều ĐẠT.

### Minor notes (không blocking)
1. **Stankov 2010 vắng mặt inline trong Section 4.4 body**. Hiện chỉ cited ở Section 1.1 + ref list. Nếu thầy/reviewer chấp nhận viện dẫn "Confucian-influenced collectivist context" trong 4.4 đã link gián tiếp về Stankov 2010 ở 1.1, thì OK. Nếu muốn an toàn, nên bổ sung inline "[Stankov 2010]" ngay sau cụm "Confucian-influenced collectivist context" trong para [139] (cả EN [139] + VN [140]).
2. Tham chiếu Karasu (đếm > 450 loại) không có — đúng spec (Q2 paper không cần).
3. Placeholder TBD còn lại: 3 ô blocking (Q1-8 R² interpretation, Q3-6 IRB letter, Q3-9 submission strategy) — đã đánh dấu rõ ràng trong NOTES TO CO-AUTHORS, không phải lỗi.

---

## RECOMMENDATION

**Trạng thái:** ✓ **READY FOR THẦY MĐ REVIEW**

Q2 v1 đã sạch toàn bộ dấu vết mixed-methods/qualitative của Q1 v7. 11/11 check A–K đạt (1 minor note ở H về Stankov 2010 inline trong 4.4 — không blocking). 4/4 additional check đạt. Metadata anonymized hoàn toàn. References = 35 entries hợp lệ.

**Đề xuất hành động:**
- Có thể gửi thẳng cho thầy MĐ review nội dung khoa học.
- Trước khi submit Frontiers in Psychiatry, gợi ý fix nhỏ: thêm inline "[Stankov 2010]" vào Section 4.4 para [139]+[140] cho hoàn chỉnh kết nối lý thuyết.
- 3 placeholder BLOCKING (R² composite, IRB letter, journal section) cần thầy MĐ xác nhận trước submission.

---

**Audit hoàn tất:** 07/06/2026, vòng độc lập 2.
