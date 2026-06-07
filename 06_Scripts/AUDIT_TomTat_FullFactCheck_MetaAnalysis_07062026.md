# AUDIT FACT-CHECK — META-ANALYSIS / GBD / NMA / SYSTEMATIC REVIEW
**Date:** 07/06/2026  
**Scope:** All meta-analysis, GBD, NMA, systematic review, and scoping review summaries in `Tom-tat-tung-bai/`  
**Method:** Compared each summary `.docx` (preferring `_FIXED_27052026` versions) against the first 4 pages of the original PDF (abstract + intro + early results). Used `python-docx` + `PyMuPDF (fitz)`.  
**Constraint:** READ-ONLY — no docx modified.

---

## 1. Target list identified (23 papers)

Pulled from `02_Papers-goc/canonical_index.json` (117 entries) by descriptor keywords (`meta`, `nma`, `gbd`, `sr`, `review`, `bayesian`, `systematic`, `burden`, `lmic`, `trends`, `countries`).

| ID | Descriptor | PDF available |
|---|---|---|
| QT012 | GBD ASEAN 2025 Lancet | Yes |
| QT013 | Zhameden 2025 PLOS ONE LMIC | Yes |
| QT018 | Salari SAD Prevalence Global SR/MA 2024 | Yes |
| QT020 | Liu CBT Delivery GAD NMA 2025 | Yes |
| QT023 | JAACAP Mojtabai Trends 2024 | Yes |
| QT029 | BMC Li CBT NMA 2025 | Yes |
| QT030 | GBD Trends 10-24y 2025 | Yes |
| QT031 | Islam 59Countries 2025 | Yes |
| QT034 | Korea Cho MH Trends NatSciRep 2024 | Yes |
| QT035 | Jefferies SocialAnxiety 7Countries 2020 | Yes |
| QT039 | Xian NMA SAD JAD 2024 | Yes |
| QT046 | Jagiello AcademicStress SR 2025 | Yes |
| QT048 | Chen COVID Meta141 JAD 2025 | NO (abstract-only) |
| QT049 | Zhang Bayesian CBT JCP 2026 | NO (abstract-only) |
| QT051 | Menon LMIC SEA Pacific APJPH 2025 | NO (abstract-only) |
| QT052 | Mindfulness NatureMH IPD MA 2023 | Yes (stub summary) |
| QT053 | Pharmacotherapy Anxiety Review 2020 | Yes (stub summary) |
| QT059 | Cai SchoolResilience SR/MA 2025 | NO (abstract-only) |
| QT060 | Bie GlobalAnxiety GBD 10-24y FrontPsych 2024 | Yes |
| QT064 | Stephens Photovoice ScopingReview 2023 | Yes |
| QT066 | Murphy PeerSupport ScopingReview 2024 | Yes |
| QT070 | Moore 2017 BullyingMeta WorldJPsy | NO (abstract-only) |
| QT072 | Lee 2025 CyberbullyingMeta TVA | Yes |

---

## 2. Per-ID findings

### OK (full match) — 11 papers
- **QT012** GBD ASEAN: All numbers verified (80.4M cases, +70% since 1990, VN 10.1%, Malaysia 13.2%, 10-14yo 16.3% DALYs).
- **QT018** Salari SAD: 4.7% / 8.3% / 17% across children/adolescents/youth, 38 studies, CMA v2.0 all correct.
- **QT020** Liu CBT GAD NMA: 52 trials, N=4361, mean age 43, 69.7% women, all SMDs exact.
- **QT034** Korea Cho KYRBS: 62.77% (poor) vs 40.07% (rich) stress prevalences match exactly; N=1,138,804 not stated in summary but no fabrication.
- **QT035** Jefferies SAD 7-countries: 6,825 / 16-29 / M=22.84 / 36% / 18% / 7 countries — all correct.
- **QT039** Xian SAD NMA: 30 RCTs, N=1547, SUCRA 71.2/68.4/92.2/89.6 exact.
- **QT046** Jagiello AcademicStress SR: 31 studies / 13 countries / CBT-strongest evidence / poor methodology — correct.
- **QT052** Mindfulness IPD MA: stub-only summary (no specific factual claims to fact-check).
- **QT053** Pharmacotherapy review: stub-only summary, no specific claims.
- **QT064** Stephens Photovoice: 12 studies, ~200 participants, 4 themes, 10/12 HIC — correct.
- **QT066** Murphy PeerSupport: 15 studies, Arksey & O'Malley, recovery outcomes, UCD/Jigsaw — correct.

### LOW severity — 2 papers
- **QT030** GBD Trends 10-24y: Summary says AAPC "0.84%" then "0,84–0,97%". PDF abstract: **"0.80% to 0.97%"**. Summary lower-bound is 0.04pp off. Affiliation (Xinxiang Medical University) correct.
- **QT031** Islam 59Countries: Summary: SEA prevalence **3.6%**. PDF: **3.78%**. Summary also omits N=179,937. Rounding/imprecise.

### MEDIUM severity — 2 papers
- **QT013** Zhameden LMIC SR: Summary states authors are at "**University of York UK**". PDF header clearly shows ALL THREE authors at **University College London (UCL)**. Wrong institutional attribution.
- **QT023** Mojtabai JAACAP: Summary tags year **2024** and JAACAP "Vol. 64(8)"; PDF citation: **"J Am Acad Child Adolesc Psychiatry 2025;64(8):906-920"** — published 2025, not 2024. All AOR / N=13,684,154 numbers correct.

### HIGH severity — 3 papers (with fabricated / factually wrong claims)

#### QT029 BMC Li CBT NMA 2025 — WRONG #1 INTERVENTION CLAIM
Summary asserts repeatedly: *"CBT cá nhân xếp hạng 1"*, *"CBT cá nhân hiệu quả nhất trong 7+ loại can thiệp — phù hợp AJP 2024"*, conclusion: *"CBT cá nhân hiệu quả nhất, CBT nhóm xếp thứ 2–3"*.
**PDF abstract is unambiguous:** *"Acceptance and Commitment Therapy was the most effective intervention (MD = −3.83). Cognitive Behavioral Therapy was the second most effective (MD = −3.64), followed by Virtual Reality Exposure Therapy (MD = −2.53) and Physical Exercise (MD = −2.16)."* The paper compares ACT vs CBT vs VRET vs Exercise — NOT individual-CBT vs group-CBT. The summary's central message is wrong; downstream RAG retrieval will mislead the dissertation.

#### QT060 Bie GlobalAnxiety GBD 10-24y FrontPsych 2024 — FABRICATED REGIONAL & RATE CLAIMS
Summary lists specific facts that DO NOT appear in PDF abstract:
- *"Tỷ lệ mới mắc: 708 → 883 / 100.000 dân"* — these exact figures not in abstract.
- *"Nhóm 20–24 tuổi tăng nhiều nhất: +28,33%"* — PDF says **10–14 age group** had the biggest increase in incidence; DALYs rose in 20–24, not incidence.
- *"Latin Mỹ Tropical cao nhất; Đông Á giảm nhiều nhất"* — PDF says **middle-SDI** regions had highest incidence/prevalence; **high-SDI** had largest increases; **India** highest cases, **Mexico** greatest rise. No "Latin Tropical highest" claim, no "East Asia largest decrease" claim.
- *"Nam Á dẫn đầu số ca tuyệt đối: 3,5 triệu (Ấn Độ chủ yếu)"* — India top yes, but "3.5 million Nam Á" figure not in abstract.
- *"+52% số ca lo âu mới"* — this number **is** in PDF (correct).

Confirms the prior audit flag. This summary appears to have been generated from a wrong source or hallucinated regional breakdown.

#### QT072 Lee 2025 CyberbullyingMeta TVA — FABRICATED N AND MISSING AUTHORS
- **Sample size fabricated**: Summary says *"N = 27.133 trẻ em + vị thành niên"*. PDF abstract: *"yielding a sample of 27 studies encompassing 13,497 children and adolescents aged 8 to 19 years old"*. Real N = **13,497**, NOT 27,133. The fabricated number appears to be "27 studies × ~1000" approximation.
- **Authors wrong/missing**: Summary lists *"Jungup Lee, Hoi Shan Cheung, Qiyang Zhang, Rebecca P. Ang"* (4 authors). PDF: **"Jungup Lee, Hyekyung Choo, Yijing Zhang, Hoi Shan Cheung, Qiyang Zhang, Rebecca P. Ang"** (6 authors). Two authors (**Hyekyung Choo** and **Yijing Zhang**) are omitted.
- **Country breakdown** *"Mỹ (12), Trung Quốc (7), Tây Ban Nha (2)…"* not present in PDF abstract — likely fabricated; needs verification from full PDF body.

### Abstract-only (cannot verify against PDF) — 5 papers
QT048, QT049, QT051, QT059, QT070 explicitly flagged as abstract-only in their own headers (`⚠ LƯU Ý: dựa trên ABSTRACT`). Cannot run PDF cross-check. These should be re-audited once full PDFs obtained. QT070 (Moore 2017) is a widely-cited classic and OR=1.77 anxiety is recognizable — likely correct but unverified here.

---

## 3. Pattern analysis

**Journal/topic distribution of errors:**
- **GBD analyses (QT060)** — MOST ERROR-PRONE: fabricated regional rankings + specific incidence rates. QT060 (Front Psychiatry) has 5+ fabricated claims. QT030 (J Affect Disord) has only a 0.04pp rounding error.
- **Chinese-language journals / Chinese authors (QT029 BMC Li, QT060 Bie Frontiers)** — both HIGH severity. Suggests the bilingual layer in the AI summarization workflow may struggle with these papers.
- **NMA papers (QT020, QT029, QT039)** — mixed: QT020 and QT039 OK, QT029 catastrophically wrong (likely confused intervention rankings).
- **GBD ASEAN (QT012)** — fully accurate, possibly because numbers come from a clear, single table.
- **Trauma/Violence journal (QT072)** — N fabrication (27,133 vs 13,497) consistent with prior audit's flag.
- **Scoping reviews (QT064, QT066)** — both OK; lower-stakes qualitative claims tend not to be fabricated.

---

## 4. Recommended fix priority

| Priority | Paper | Action |
|---|---|---|
| 1 (urgent) | **QT060** | Rewrite results section: replace fabricated regional claims with PDF-verified facts (middle-SDI highest incidence, high-SDI largest increases, India top cases, Mexico top rise, 10–14 biggest incidence jump). Remove "708 → 883/100k" and "20–24 +28.33%" unless verified from full body. |
| 2 (urgent) | **QT072** | Fix N (27,133 → **13,497**). Add missing authors (**Hyekyung Choo, Yijing Zhang**). Verify or remove country breakdown "Mỹ(12) TQ(7)…" from full PDF. |
| 3 (urgent) | **QT029** | Rewrite headline finding: ACT ranked #1 (MD=−3.83), CBT #2 (MD=−3.64), VRET #3, Exercise #4. Remove fabricated "CBT cá nhân vs CBT nhóm" framing — paper does not stratify CBT by delivery format. |
| 4 (medium) | **QT023** | Correct publication year 2024 → **2025** (JAACAP 2025;64(8):906–920). |
| 5 (medium) | **QT013** | Correct affiliation **University of York → University College London (UCL)**. |
| 6 (low) | **QT030** | AAPC lower bound 0.84% → **0.80%**. |
| 7 (low) | **QT031** | SEA prevalence 3.6% → **3.78%**; add N=179,937. |
| 8 (deferred) | QT048, QT049, QT051, QT059, QT070 | Re-audit when full PDFs are obtained. Currently self-flagged as abstract-only. |

---

## 5. Files referenced
- Source index: `c:/Users/OS/OneDrive/read_books/Lo-au/02_Papers-goc/canonical_index.json`
- Summary docs: `c:/Users/OS/OneDrive/read_books/Lo-au/Tom-tat-tung-bai/QT{012..072}_*.docx`
- Extract helper: `c:/Users/OS/OneDrive/read_books/Lo-au/06_Scripts/_audit_batch.py`
- Per-ID raw extracts: `c:/Users/OS/OneDrive/read_books/Lo-au/_audit_tmp/QT*.txt`
- Per-ID comparison views: `c:/Users/OS/OneDrive/read_books/Lo-au/_audit_tmp/_view_QT*.txt`
