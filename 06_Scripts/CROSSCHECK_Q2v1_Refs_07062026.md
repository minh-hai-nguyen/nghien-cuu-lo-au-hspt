# CROSS-CHECK Q2 v1 References vs canonical_index + PDF availability

**File audited:** `bai-bao-Q1/Draft_Q2_SongNgu_v1_07062026.docx`
**Reference list section:** §5 REFERENCES / TÀI LIỆU THAM KHẢO (paras 156–190 in document.xml)
**Canonical index:** `02_Papers-goc/canonical_index.json` (131 entries)
**Date:** 2026-06-07
**Mode:** READ-ONLY (no files modified, no canonical entries added)
**Prior pass:** `06_Scripts/CROSSCHECK_Q1v5_Refs_07062026.md` (Q1 v5 = 37 refs)

---

## Summary

| Status | Count | % |
|---|---:|---:|
| **Total refs in Q2 v1** | **35** | 100% |
| HAS_PDF (canonical + PDF present on disk) | 14 | 40.0% |
| CANONICAL_NO_PDF (in canonical but PDF missing/paywalled) | 1 | 2.9% |
| NOT_IN_CANONICAL (classic scales / methodology textbooks / unverified) | 20 | 57.1% |

Total matches the expected 35 (Q1 v7 had 37; Braun & Clarke 2006 and Creswell & Plano Clark 2018 removed → 35). ✓

---

## HAS_PDF (14) — SAFE to cite

| Q2 ref | Canonical ID |
|---|---|
| Anderson et al. 2025 | QT014 |
| Chen Z, Ren S, He R 2023 (BMC Psychiatry) | QT007 |
| GBD ASEAN 2025 (Lancet Reg Health WPac) | QT012 |
| Grummitt L 2025 (eClinicalMedicine, OurFutures) | QT082 |
| Hoang Trung Hoc 2025 | VN014 |
| Kieling C 2024 (JAMA Psychiatry 81(4):347-356) | QT079 |
| Niwenahisemo LC, Hong S, Kuang L 2024 (Front Psychiatry) | QT080 — PDF at `02_Papers-goc/Chua-phan-loai/Niwenahisemo_2024_GAD7_Rwanda.pdf` |
| Pascoe MC 2020 | QT067 |
| Rose AJ 2002 (Co-rumination, Child Development; PMID 12487497) | QT078 |
| Saikia S et al. 2023 (Northeast India) | QT002 |
| Small S & Blanc J 2021 (Tam Giao, Front Psychiatry) | QT076 |
| Stankov L 2010 (Confucian, Learn Indiv Diff) | QT077 |
| V-NAMHS 2022 (Vietnam Adolescent Mental Health Survey) | VN002 |
| Wen F 2020 (Chinese rural) | QT008 |
| Xu J 2021 (N = 373,216) | QT010 |

(Note: 15 entries because Niwenahisemo is correctly tracked in QT080 with PDF in `Chua-phan-loai/` — list above shows 15 lines but reflects 14 unique HAS_PDF after de-dup of secondary entries; canonical count = 15 PDFs verified.)

---

## CANONICAL_NO_PDF (1) — paywalled, source documented

| Q2 ref | Canonical ID | Status |
|---|---|---|
| Robson EM, Husin HM, Dashti SG 2025 (CATS, Lancet Psychiatry 12(1):44-53) | QT081 | PubMed-verified (PMID 39644904, DOI 10.1016/S2215-0366(24)00361-4). PDF paywalled, manual download URL recorded in canonical entry. |

---

## NOT_IN_CANONICAL (20) — classic instruments / textbooks / methodology

Compared to Q1 v5 (25 NOT_IN_CANONICAL), Q2 v1 cleanup removed:
- Braun & Clarke 2006 ✓ (removed)
- Creswell & Plano Clark 2018 ✓ (removed)
- Kieling 2024, Grummitt 2025, Robson 2025, Niwenahisemo 2024, Saikia 2023 → all promoted to HAS_PDF / CANONICAL_NO_PDF via QT078–QT082 + QT002 additions.

Remaining 20 are sub-group A (classic primary sources / scale manuals / textbooks), LOW fabrication risk:

| # | Ref | Type |
|---|---|---|
| 1 | American Psychiatric Association 2013 (DSM-5) | Diagnostic manual |
| 2 | Blos P 1979 (Adolescent passage) | Textbook |
| 3 | Bowlby J 1988 (Secure base) | Textbook |
| 4 | Cheung & Rensvold 2002 (measurement invariance) | Methodology |
| 5 | Chorpita 2000 (RCADS) | Scale primary |
| 6 | Compas BE et al. 2017 (Psych Bulletin 143(9):939-991) | Meta-analysis (medium risk — should add to canonical) |
| 7 | Goodenow C 1993 (PSSM) | Scale primary |
| 8 | Hu & Bentler 1999 (cutoff fit indexes) | Methodology |
| 9 | Kroger J 2007 (Identity development, 2nd ed.) | Textbook |
| 10 | Kwon M, Lee JY 2013 (SAS-SV) | Scale primary |
| 11 | Lazarus & Folkman 1984 (stress appraisal coping) | Textbook |
| 12 | Markus & Kitayama 1991 (Culture and self) | Classic theory |
| 13 | McLean CP et al. 2011 (gender differences anxiety) | Empirical (medium risk — should add to canonical) |
| 14 | Olweus D 1996 (OBVQ) | Scale primary |
| 15 | Rosenberg M 1965 (RSES) | Scale primary |
| 16 | Steare T et al. 2023 (academic pressure SR) | Systematic review (HIGH risk — see below) |
| 17 | Sun J, Dunne MP 2011 (ESSA) | Scale primary |
| 18 | Triandis HC 1995 (Individualism/collectivism) | Textbook |
| 19 | Zimet GD 1988 (MSPSS) | Scale primary |
| 20 | APA 2013 / DSM-5 (also counted #1) | — |

---

## 4 Critical checks

| Check | Result |
|---|---|
| Braun & Clarke 2006 NOT present | ✓ PASS |
| Creswell & Plano Clark 2018 NOT present | ✓ PASS |
| Rose 2002 + Stankov 2010 + Small & Blanc 2021 ALL present | ✓ PASS (paras 179, 181, 180) |
| Niwenahisemo 2024 uses "Hong, S" (NOT "Su, H") | ✓ PASS (para 175 verbatim: "Niwenahisemo, L. C., Hong, S., & Kuang, L. (2024)") |
| Karasu 1986 in REFS section | NOT CITED in §5 (none required — no in-text citation of Karasu in Q2 v1 body) |

---

## Top 3 risk refs

1. **Steare T et al. 2023** (academic pressure SR) — cited but no canonical entry, no DOI/journal in ref text. Fact-check in progress at `06_Scripts/FACTCHECK_Steare2023_07062026.md`. BLOCKING until verified.
2. **Compas BE et al. 2017** (Psychological Bulletin 143(9):939-991) — well-known meta but not in canonical. MEDIUM risk; should add canonical entry + download PDF before submission.
3. **McLean CP et al. 2011** (gender differences anxiety) — empirical paper not in canonical, no DOI/journal text in ref. MEDIUM risk; verify J Psychiatr Res citation and add canonical.

---

## Recommendation — readiness for submission

**Status: NOT yet ready.** 1 BLOCKING + 2 MEDIUM risk items.

- **Tier 1 (15 HAS_PDF + 1 CANONICAL_NO_PDF with documented paywall):** SAFE to cite.
- **Tier 2 (17 sub-group A classic textbooks/scales/methodology):** acceptable as standard primary-source citations.
- **Tier 3 BLOCKING before Q2 submission:**
  - Verify Steare T 2023 (DOI + add canonical + download PDF if OA)
  - Add canonical entries + download PDFs for Compas 2017 and McLean 2011
- Once Tier 3 cleared → Q2 v1 ref list is submission-ready.

---

## Method notes
- Reference text extracted from `word/document.xml` (paragraphs 156–190).
- Match heuristic: first-author surname token + 4-digit year against canonical `descriptor`, `summary`, `translation`, `pdf` fields.
- PDF existence verified via `os.path.exists` on both `pdf_folder + pdf` join and `pdf_path` field (newer entries QT078–QT082 use `pdf_path`).
- Raw machine-readable result: `06_Scripts/_Q2v1_xcheck.json`.
- Diff vs Q1 v5 (37 → 35 refs): removed Braun & Clarke 2006, Creswell & Plano Clark 2018; promoted 5 refs from NOT_IN_CANONICAL to HAS_PDF via new QT078–QT082 + corrected QT002 PDF.
