# BÁO CÁO KIỂM TRA ĐẠO VĂN — Q2 v1 (Frontiers in Psychiatry)

**Đối tượng:** `bai-bao-Q1/Draft_Q2_SongNgu_v1_07062026.docx` (bản polished, Frontiers numbered citations).
**Script nguồn:** `06_Scripts/sinh_Draft_Q2_SongNgu_07062026.py` (đã chỉnh sửa 1 paraphrase trong báo cáo này).
**Ngày soạn:** 07/06/2026.
**Vòng:** Plagiarism check chặt chẽ trước khi nộp Frontiers Crossref Similarity Check (iThenticate).

---

## 0. TÓM TẮT TỔNG QUAN

| Dimension | Kết quả | Mức rủi ro |
|---|---|---|
| **1. Self-plagiarism (Q2 v1 vs LA chính tiếng Việt)** | 7-gram match 1,5%; 10-gram 0,5%; 15-gram 0,0%; tất cả 5 cụm ≥12 từ là technical terms không thể paraphrase | **THẤP** |
| **2. Word-for-word vs cited papers (35 refs)** | 1 cụm HIGH (co-rumination definition trong §4.4) — **đã sửa**; còn lại all clear | **THẤP sau khi sửa** |
| **3. Methods scale descriptions (Crossref hot zone)** | Mô tả ESSA/SAS-SV/OBVQ/PSSM/MSPSS/RSES/RCADS đều paraphrase từ phong cách gốc; không copy abstract | **THẤP** |

**Verdict cuối:** **READY for Frontiers Crossref Similarity Check** sau khi chỉnh 1 cụm co-rumination (đã áp dụng trong commit này). Dự kiến similarity score iThenticate < 15% (target Frontiers < 20%).

---

## 1. SELF-PLAGIARISM — Q2 v1 (VI) vs LA chính (VI)

### Phương pháp
- Trích xuất 87 chuỗi tiếng Việt từ script Q2 (toàn bộ phần VI của các call `PP/H3/H2/H1/NOTE/TBD`) — tổng **6.452 từ VI**.
- LA chính `01_LuanAn_v3_1_FixCoping_28052026.docx`: 77.527 từ.
- N-gram match: 7, 10, 15 từ liên tiếp. Quét long contiguous match (12–60 từ).

### Kết quả

| Đo lường | Số trùng / tổng unique Q2 | Tỷ lệ |
|---|---|---|
| 7-grams shared | 89 / 6.052 | 1,5% |
| 10-grams shared | 29 / 6.223 | 0,5% |
| 15-grams shared | **0 / 6.345** | **0,0%** |
| Long contiguous ≥12 từ | **5 cụm** | n/a |

### Phân tích 5 long contiguous matches (≥12 từ)

| # | Cụm trùng (≥12 từ) | Loại | Đánh giá |
|---|---|---|---|
| 1 | "sổ tay chẩn đoán và thống kê các rối loạn tâm thần phiên bản" (14w) | Tên chính thức DSM-5 tiếng Việt | UNAVOIDABLE — tên chuẩn của Hiệp hội Tâm thần học Hoa Kỳ; bắt buộc giữ nguyên |
| 2 | "rối loạn lo âu là một trong những vấn đề sức khỏe tâm thần" (14w) | Cụm stock mô tả tầm quan trọng | UNAVOIDABLE — xuất hiện trong hàng nghìn bài VI cùng chủ đề |
| 3 | "rối loạn lo âu ở học sinh trung học cơ sở việt nam" (13w) | Mô tả dân số nghiên cứu | UNAVOIDABLE — tên dân số = tên dân số |
| 4 | "áp lực học tập nghiện điện thoại và bắt nạt học đường" (12w) | Liệt kê tên 3 biến nguy cơ (ESSA/SAS-SV/OBVQ) | UNAVOIDABLE — đây là chính tên biến trong nghiên cứu |
| 5 | "gắn bó trường học hỗ trợ cha mẹ hỗ trợ bạn bè" (12w) | Liệt kê tên 3 biến bảo vệ (PSSM/MSPSS-P/MSPSS-Peer) | UNAVOIDABLE — tên biến |

**Tự đạo văn overall:** 0,0% ở mức 15-gram, **không có một câu nào** trong Q2 sao chép từ LA. Toàn bộ 5 cụm flagged là (a) tên chính thức DSM-5, (b) tên biến nghiên cứu, hoặc (c) cụm stock mô tả phổ biến — không paraphraseable và không yêu cầu sửa.

### So sánh với Q1 v4 (báo cáo trước)
Mẫu hình identical với Q1 v4: 5 cụm flag cùng loại, không có cụm prose có nội dung diễn giải bị copy. Confirms Q2 v1 viết lại từ đầu chứ không paste từ LA.

---

## 2. WORD-FOR-WORD vs PAPER GỐC (35 refs)

### Phương pháp
- Quét WebSearch trên 8–10 từ liên tiếp ở các đoạn cite-heavy: §1.1 (Stankov 2010, Markus & Kitayama, Triandis), §1.2 (Compas, Lazarus & Folkman), §1.3 (McLean 2011), §4.4 (Rose 2002 co-rumination, Small & Blanc tam giao), Methods (Chorpita 2000 RCADS).

### Kết quả từng đoạn high-risk

| Đoạn | Cụm test (8–10 từ) | Verbatim hit? | Đánh giá |
|---|---|---|---|
| §1.1 Stankov-style | "Confucian academic culture frames high-stakes examinations as the principal gateway to social mobility" | NO | LOW — paraphrase tốt |
| §1.2 Compas | "joint contribution of risk and protective factors to youth internalising symptoms" | NO | LOW |
| §1.2 Lazarus | "adolescent psychopathology emerges from dynamic interactions between stressors and resources" | NO | LOW — original phrasing |
| §1.3 McLean | "females exhibit higher anxiety than males across subtypes" | NO | LOW — paraphrase tốt |
| §4.2 Bowlby/Triandis | "bullying victimisation in collectivist contexts is preferentially processed as an attachment-related threat" | NO | LOW — original synthesis |
| **§4.4 Rose 2002 (co-rumination)** | **"extensively discussing problems, speculating about them and focusing on negative feelings"** | **CLOSE MATCH với canonical Rose 2002 phrasing** | **HIGH — đã sửa, xem bên dưới** |
| §4.4 Small & Blanc | "tam giao tradition (the coexistence of Confucianism, Buddhism and Taoism)" | NO — tên truyền thống văn hóa, unavoidable | LOW |

### Phát hiện HIGH RISK — và bản sửa đã áp dụng

**Vấn đề:** Định nghĩa co-rumination của Rose (2002, *Child Development*) theo phrasing chuẩn là *"extensively discussing problems frequently, rehashing problems, speculating about problems including causes and consequences, and dwelling on negative affect"*. Bản Q2 v1 viết: *"extensively discussing problems, speculating about them and focusing on negative feelings"* — giữ nguyên 3 keywords liên tiếp (*extensively discussing problems*, *speculating*, *negative*). Một tool như Crossref Similarity Check / iThenticate sẽ flag cụm 8 từ này.

**Bản gốc (line 1038–1040 trong script):**
```
the co-rumination hypothesis — defined as extensively discussing problems,
speculating about them and focusing on negative feelings (34)
```

**Bản sửa đã áp dụng:**
```
the co-rumination construct (34) — a perseverative form of dyadic problem-talk
in which adolescents repeatedly revisit stressful events and their accompanying
negative affect rather than moving toward resolution
```

**Lý do paraphrase work:** thay (a) "hypothesis" → "construct" + cite ngay sau; (b) "extensively discussing" → "perseverative form of dyadic problem-talk"; (c) "speculating about them" → "repeatedly revisit"; (d) "focusing on negative feelings" → "negative affect rather than moving toward resolution". Toàn bộ nội dung lý thuyết giữ nguyên — chỉ thay surface wording.

**Bản tiếng Việt cũng đã được đồng bộ** để giữ tính bilingual.

---

## 3. METHODS SCALE DESCRIPTIONS (Crossref hot zone)

### Phương pháp
Quét §2.2 Measures (8 scale descriptions) so với cách viết abstract của paper validation gốc.

### Kết quả

| Thang đo | Câu Q2 v1 | Verbatim với validation paper? | Đánh giá |
|---|---|---|---|
| RCADS (Chorpita 2000) | "a developmentally calibrated, DSM-5-aligned instrument" | NO — Q2 dùng phrasing original (Chorpita gốc nói "Revised Child Anxiety and Depression Scale") | LOW |
| ESSA (Sun 2011) | "Educational Stress Scale for Adolescents (ESSA), 4 items" | NO — chỉ giữ tên chính thức scale + số items (fact bắt buộc) | LOW |
| SAS-SV (Kwon 2013) | "Smartphone Addiction Scale–Short Version (SAS-SV), 5 items" | NO — tên chính thức + số items | LOW |
| OBVQ (Olweus 1996) | "Olweus Bully/Victim Questionnaire (OBVQ), 8 items capturing physical and verbal victimisation" | NO — phrasing original | LOW |
| PSSM (Goodenow 1993) | "Psychological Sense of School Membership (PSSM), 7 items" | NO — tên scale + số items | LOW |
| MSPSS (Zimet 1988) | "Multidimensional Scale of Perceived Social Support (MSPSS), 8 items split into parental and peer subscales" | NO — phrasing original | LOW |
| RSES (Rosenberg 1965) | "Rosenberg Self-Esteem Scale (RSES), 5 items" | NO — tên scale + số items | LOW |
| GAD-7 (Niwenahisemo 2024) | "procedures established for cross-gender invariance of the GAD-7 in adolescent samples" | NO | LOW |

**Tất cả 8 mô tả thang đo PASS.** Phong cách viết của Q2 v1 chỉ giữ lại "fact bắt buộc" (tên chính thức scale + số items + định dạng câu trả lời) và viết phần mô tả bằng phong cách riêng — đây là chuẩn vàng cho Methods sections trong international peer review.

---

## 4. TỔNG KẾT — 3 DIMENSIONS

| Dimension | HIGH | MEDIUM | LOW | Action |
|---|---|---|---|---|
| 1. Self-plagiarism vs LA | 0 | 0 | 5 (unavoidable) | NONE |
| 2. Word-for-word vs cited papers | 1 (co-rumination) | 0 | 6 | **ĐÃ SỬA** |
| 3. Methods scale descriptions | 0 | 0 | 8 | NONE |
| **TỔNG** | **1 (đã sửa)** | **0** | **19** | **Ready** |

---

## 5. TOP 3 RECOMMENDED PARAPHRASES

### #1 — §4.4 co-rumination definition (ĐÃ ÁP DỤNG)
- **Trước:** *"the co-rumination hypothesis — defined as extensively discussing problems, speculating about them and focusing on negative feelings"*
- **Sau:** *"the co-rumination construct — a perseverative form of dyadic problem-talk in which adolescents repeatedly revisit stressful events and their accompanying negative affect rather than moving toward resolution"*

### #2 — KHUYẾN NGHỊ THÊM (tùy chọn, không bắt buộc): §1.1 Stankov phrasing
- **Hiện tại:** *"a Confucian academic culture (11) frames high-stakes examinations as the principal gateway to social mobility, intensifying chronic academic stress."*
- **Đề xuất tùy chọn:** Cụm "high-stakes examinations as the principal gateway to social mobility" có thể xuất hiện trong Stankov 2010 hoặc derivative papers. Đã test WebSearch → không có verbatim hit. Tuy nhiên nếu muốn an toàn 100% với iThenticate, đổi thành: *"a Confucian academic culture (11) elevates examination performance to a chief mechanism of upward social mobility, generating chronic academic stress."*
- **Rủi ro nếu giữ nguyên:** thấp — phong cách viết tổng quát, chưa thấy verbatim hit trên web.

### #3 — KHUYẾN NGHỊ THÊM (tùy chọn): §1.2 Compas meta-analysis phrasing
- **Hiện tại:** *"A landmark meta-analysis by Compas and colleagues (212 studies, N = 80,850) confirmed the joint contribution of risk and protective factors to youth internalising symptoms"*
- **Đề xuất tùy chọn:** Cụm "joint contribution of risk and protective factors to youth internalising symptoms" — có thể trùng với phrasing trong abstract Compas 2017. Đã test WebSearch → no verbatim hit. Nếu muốn an toàn: *"a meta-analytic synthesis by Compas and colleagues (212 studies, N = 80,850) demonstrated that risk and protective factors jointly shape youth internalising outcomes"*.
- **Rủi ro nếu giữ nguyên:** thấp.

**Lưu ý:** Em đã chỉ apply #1 (HIGH risk thực sự); #2 và #3 chỉ là "extra safety" và **không cần thiết** cho vòng nộp này. Thầy + NCS có thể chỉnh thêm nếu muốn similarity score < 10%.

---

## 6. VERDICT

**Q2 v1 đã sẵn sàng cho Frontiers Crossref Similarity Check.**

- Tự đạo văn vs LA: 0,0% ở 15-gram (xuất sắc).
- Cited paper exposure: 1 HIGH risk → đã sửa.
- Methods scale descriptions: 0 risk.
- Dự kiến iThenticate similarity score: **8–15%** (well below Frontiers cảnh báo ngưỡng 20%; well below auto-reject ngưỡng 25%).

**Hành động bắt buộc trước khi submit:** KHÔNG có. Bản hiện tại sau khi paraphrase co-rumination là sẵn sàng.

**Hành động khuyến nghị tùy chọn (an toàn 100%):** Áp dụng paraphrase #2 và #3 ở trên nếu thầy + NCS muốn similarity score < 10%.

**Hành động ngoài plagiarism (vẫn còn pending):** 3 BLOCKING items đã đánh dấu TBD trong Q2 v1 (Q1-8 SEM re-analysis decision; Q3-6 IRB decision number; Q3-9 Frontiers section choice) — không liên quan đạo văn nhưng cần resolve trước submission.

---

## THAM KHẢO ĐÃ KIỂM CHỨNG TRONG VÒNG NÀY

- **Rose, A. J.** (2002). Co-rumination in the friendships of girls and boys. *Child Development*, 73(6), 1830–1843. PMID: 12487497. Verified canonical phrasing via [PubMed](https://pubmed.ncbi.nlm.nih.gov/12487497/) + [Wiley](https://srcd.onlinelibrary.wiley.com/doi/abs/10.1111/1467-8624.00509).
- **Chorpita, B. F., Yim, L., Moffitt, C., Umemoto, L. A., & Francis, S. E.** (2000). Assessment of symptoms of DSM-IV anxiety and depression in children: A revised child anxiety and depression scale. *Behaviour Research and Therapy*, 38(8), 835–855. Verified abstract via [PubMed 10937431](https://pubmed.ncbi.nlm.nih.gov/10937431/).
- **Stankov, L.** (2010). Unforgiving Confucian culture. *Learning and Individual Differences*, 20(6), 555–563.
- **McLean, C. P., Asnaani, A., Litz, B. T., & Hofmann, S. G.** (2011). Gender differences in anxiety disorders. *J Psychiatr Res*, 45(8), 1027–1035. Verified via [ScienceDirect](https://www.sciencedirect.com/science/article/abs/pii/S0022395611000458).
- **Sun, J., Dunne, M. P., Hou, X. Y., & Xu, A.** (2011). Educational stress scale for adolescents. *J Psychoeduc Assess*, 29(6), 534–546. Verified via [SAGE Journals](https://journals.sagepub.com/doi/abs/10.1177/0734282910394976).

## TRUY VẾT RAG/KG/GLOSSARY NỘI BỘ
- LA chính: `Luận án TS/01_LuanAn_v3_1_FixCoping_28052026.docx` (1820 paragraphs, 77.527 từ VI).
- Q2 script đã chỉnh: `06_Scripts/sinh_Draft_Q2_SongNgu_07062026.py` (line 1038–1046 + bản VI line 996–1003).
- Q2 docx tái tạo: `bai-bao-Q1/Draft_Q2_SongNgu_v1_07062026.docx` (63.897 bytes, metadata đã strip).
- N-gram match raw: `_audit_tmp/q2_vs_la_long_matches.txt`, `_audit_tmp/q2_vs_la_10gram.txt`, `_audit_tmp/q2_vs_la_15gram.txt`.
- Báo cáo Q1 vòng trước (so sánh pattern): `bai-bao-Q1/BaoCao_Plagiarism_Q1_v1_01062026.docx`.
