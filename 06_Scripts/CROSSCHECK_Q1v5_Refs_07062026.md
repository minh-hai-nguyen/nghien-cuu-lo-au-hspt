# CROSS-CHECK Q1 v5 References vs canonical_index + PDF availability

**File audited:** `bai-bao-Q1/Draft_Q1_SongNgu_v5_01062026.docx`
**Reference list section:** §5 REFERENCES / TÀI LIỆU THAM KHẢO (paras 207–243 in document.xml)
**Canonical index:** `02_Papers-goc/canonical_index.json` (117 entries; 81 with PDF on disk, 36 missing PDF)
**Date:** 2026-06-07
**Mode:** READ-ONLY (no files modified, no canonical entries added)

---

## Summary

| Status | Count | % |
|---|---:|---:|
| **Total refs in Q1 v5** | **37** | 100% |
| HAS_PDF (canonical + PDF present) | 8 | 21.6% |
| CANONICAL_NO_PDF (in canonical but PDF missing) | 4 | 10.8% |
| NOT_IN_CANONICAL (uncited in master DB) | 25 | 67.6% |

---

## HAS_PDF (8) — SAFE to cite

| Q1 ref | Canonical ID | Folder |
|---|---|---|
| Anderson et al. 2025 | QT014 Anderson_2025_Wiley_Narrative | – |
| Chen Z, Ren S, He R 2023 (BMC Psychiatry) | QT007 Chen_et_al_2023_China_BMCPsych | – |
| Hoang Trung Hoc 2025 | VN014 HoangTrungHoc_2025_VN_COVID | Viet-Nam |
| GBD ASEAN 2025 (Lancet Reg Health WPac) | QT012 GBD_ASEAN_2025_Lancet | – |
| Pascoe MC 2020 | QT067 Pascoe_AcademicStress_IJAY_2020 | – |
| V-NAMHS 2022 | VN002 VNAMHS_2022_National | Viet-Nam |
| Wen F 2020 (Chinese rural) | QT008 Wen_et_al_2020_China_Rural | – |
| Xu J 2021 (N = 373,216) | QT010 Xu_et_al_2021_China_LargestEpi | – |

---

## CANONICAL_NO_PDF (4) — need to download PDF before cite

| Q1 ref | Canonical ID | Action |
|---|---|---|
| Rose AJ 2002 (Co-rumination, Child Development; PMID 12487497) | QT078 Rose_2002_USA_CoRumination | Download PDF |
| Small S & Blanc J 2021 (Tam Giao, Front Psychiatry; DOI 10.3389/fpsyt.2020.589618) | QT076 Small_Blanc_2021_USA_TamGiao_VN | Download PDF (Front Psych is OA) |
| Stankov L 2010 (Confucian, Learn Indiv Diff; DOI 10.1016/j.lindif.2010.05.003) | QT077 Stankov_2010_Australia_ConfucianAcademic | Download PDF |
| Saikia S et al. 2023 (Northeast India) | QT002 Saikia_et_al_2023_India_Assam | Download PDF |

---

## NOT_IN_CANONICAL (25) — HIGHEST RISK

These references are cited in Q1 v5 but no entry exists in `canonical_index.json`. They split into two sub-groups:

### Sub-group A: Classic methodology / scale primary sources (15)
Expected to be cited from textbooks or original instrument papers; NOT empirical papers tracked in QT/VN/BB canonical. LOW fabrication risk but should still source the original PDF or page citation if available.

- American Psychiatric Association 2013 (DSM-5)
- Blos 1979 (Adolescent passage – book)
- Bowlby 1988 (Secure base – book)
- Braun & Clarke 2006 (Thematic analysis – QRP)
- Cheung & Rensvold 2002 (measurement invariance)
- Chorpita 2000 (RCADS scale)
- Creswell & Plano Clark 2018 (mixed methods textbook)
- Goodenow 1993 (PSSM)
- Hu & Bentler 1999 (cutoff fit indexes)
- Kroger 2007 (Identity development – book)
- Kwon M, Lee JY 2013 (SAS-SV)
- Lazarus & Folkman 1984 (stress appraisal coping – book)
- Markus & Kitayama 1991 (Culture and self)
- Olweus 1996 (OBVQ)
- Rosenberg 1965 (RSES)
- Sun J, Dunne MP 2011 (ESSA scale)
- Triandis 1995 (Individualism/collectivism)
- Zimet 1988 (MSPSS)

### Sub-group B: Recent empirical / review papers — HIGH PRIORITY VERIFY (7)
These are 2017–2025 empirical or review citations that SHOULD be tracked in canonical and have PDFs. Status unverified — possible fabricated cites.

| Q1 ref | Risk | Suggested action |
|---|---|---|
| **Kieling C 2024 (JAMA Psychiatry 81(4):347)** | HIGH — frequently mis-cited | Verify DOI 10.1001/jamapsychiatry.2023.5051; add to canonical; download |
| **Steare T 2023 (academic pressure SR)** | HIGH | Verify (likely Lancet Psychiatry 10:11 or J Affect Disord); fact-check already started at `06_Scripts/FACTCHECK_Steare2023_07062026.md` |
| **Robson EM, Husin HM, Dashti SG 2025 (CATS, Lancet Psychiatry 12:44–53)** | HIGH | Verify Lancet Psychiatry vol/page; add canonical |
| **Niwenahisemo LC, Su H, Kuang L 2024 (GAD-7 Rwandan, Front Psychiatry)** | HIGH | Verify DOI |
| **Grummitt L, O'Dean S 2025 (OurFutures MH, eClinicalMedicine)** | HIGH | Verify DOI |
| **Compas BE et al. 2017 (Psych Bulletin 143(9):939–991)** | MEDIUM — well-known meta but not in DB | Add canonical + PDF |
| **McLean CP et al. 2011 (gender differences anxiety)** | MEDIUM | Verify J Psychiatr Res citation; add canonical |

---

## Recommendation — SAFE-to-submit subset for Q1

**Tier 1 — cite immediately:** the 8 HAS_PDF refs.
**Tier 2 — cite after PDF download:** the 4 CANONICAL_NO_PDF (Rose, Small & Blanc, Stankov, Saikia).
**Tier 3 — sub-group A (scales/textbooks):** acceptable if standard primary-source citation, but flag each in coversheet.
**Tier 4 — BLOCKING:** the 7 sub-group B refs MUST be verified via DOI/PubMed + PDF downloaded + added to `canonical_index.json` before Q1 submission. Per workspace principle "tất cả bài test phải dựa paper gốc," any of these 7 left unverified is a blocking fabrication risk.

---

## Method notes
- Reference text extracted from `word/document.xml` (paragraphs 207–243).
- Match heuristic: first-author surname token + 4-digit year against canonical `descriptor` field; fallback keyword scan over descriptor + summary + pdf field.
- PDF existence verified via filesystem `os.path.exists` on `pdf_folder/pdf` join.
- Raw machine-readable result saved to `06_Scripts/_Q1v5_xcheck.json`.
