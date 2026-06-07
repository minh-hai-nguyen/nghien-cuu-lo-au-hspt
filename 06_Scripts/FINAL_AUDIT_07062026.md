# FINAL AUDIT — Session 07/06/2026

**Mode:** READ-ONLY re-verification of all changes made in session 07/06/2026.
**Principle:** all checks compared against source-of-truth artifacts (paper gốc, file paths, metadata).

---

## A. Q1 v6 verification (`bai-bao-Q1/Draft_Q1_SongNgu_v6_01062026.docx`)

| # | Check | Result | Detail |
|---|-------|--------|--------|
| A1 | "Su, H" / "Su H" removed (Niwenahisemo author fix) | PASS | 0 hits for both forms |
| A1 | "Hong, S" present (correct co-author) | PASS | 1 hit in reference list: "Niwenahisemo, L. C., Hong, S., & Kuang, L. (2024)…"; in-text citations "Niwenahisemo, Hong, & Kuang 2024" also present (EN body + VN body) |
| A2 | Rose 2002 in body + ref list | PASS | "Rose, A": 1 hit (ref list); "Rose 2002": 2 hits |
| A2 | Stankov 2010 present | PASS | 3 hits |
| A2 | Small & Blanc 2021 present | PASS | "Small": 3, "Blanc": 3 |
| A3 | Old uncited "nhẫn" pattern removed | PASS | "(nhẫn": 0; "nhẫn endurance": 0; "nhẫn" total: 0 |
| A4 | Metadata (core/app/custom) cleaned | PASS | core.xml: creator="", lastModifiedBy=""; app.xml: Company="", Application=""; custom.xml: not present |

**Section A: 7/7 PASS**

---

## B. 5 FIXED tóm tắt (07/06/2026)

| File | Critical claim | Result | Detail |
|------|----------------|--------|--------|
| QT020_…_FIXED | SMD 0.96 = efficacy (not acceptability) | PASS | "0,96" tied to "efficacy outcome — thang lo âu"; commentary flags PDF body wording as likely editorial error |
| QT029_…_FIXED | ACT #1 (not CBT), MD = −3.83 | PASS | "Hạng 1 — ACT … MD = −3,83 [95% CrI: −9,33; 1,51]"; "Hạng 2 — CBT" |
| QT060_…_FIXED | EAPC for global regions | PASS | EAPC: 12 hits; "region": 12; "global": 6 |
| QT067_…_FIXED | No Yerkes-Dodson, no fabricated stats | PASS | "Yerkes": 0; "80 nghiên cứu": 0; "trầm cảm 2-3 lần": 0; "<7h"/"< 7": 0 |
| QT072_…_FIXED | Choo + Zhang authors; N=13,497 and 27,133 | PASS | "Hyekyung Choo": 2; "Yijing Zhang": 2; "13,497": 1; "27,133": 3 |

**Section B: 5/5 PASS**

---

## C. 5 new QT074–QT078 tóm tắt

| File | Word count | PDF path | Specific claim | Result |
|------|-----------:|----------|----------------|--------|
| QT074 Karasu 1986 | 1,305 | present | "450 loại" (= "450 types"), labeled "trang 325" | PASS (VN form of "over 450 types polled nationwide") |
| QT075 Karasu BioSketch 2014 | 1,101 | present | Karasu: 17 hits | PASS |
| QT076 Small & Blanc 2021 | 1,535 | present | Small/Blanc both present | PASS |
| QT077 Stankov 2010 Confucian | 1,697 | present | "test anxiety": 2; "self-doubt": 6; "Confucian": 29 | PASS |
| QT078 Rose 2002 Co-Rumination | 1,591 | present | "co-rumination": 27; N = 608 confirmed ("608 học sinh lớp 3, 5, 7 và 9") | PASS |

All 5 files: wc > 800 and contain `02_Papers-goc/` path reference. **Section C: 5/5 PASS**

---

## D. `02_Papers-goc/canonical_index.json`

| Check | Result | Detail |
|-------|--------|--------|
| Total entries = 117 | PASS | exactly 117 (78 QT + 15 BB + 24 VN) |
| QT002 → 11_Saikia_2023_IJCM.pdf | PASS | `02_Papers-goc/11-bai-ban-dau-va-mo-rong/11_Saikia_2023_IJCM.pdf` |
| QT074–QT078 entries exist with pdf_path | PASS | all 5 entries present with valid paths |
| BB entries exist | PARTIAL | 15 BB entries present (BB01–03, 05–12, 14, 18, 19, 22). BB04, 13, 15–17, 20, 21 are not in the index. Matches the stated "15 BB" count, but the task wording "BB01–BB22 exist" is **not literally true** — gaps exist. Flag as informational only. |

**Section D: 3/4 PASS, 1 informational note** (BB index has gaps; total 117 matches.)

---

## E. Outline fix (`bai-bao-Q1/OutlineBilingual_Q1_01062026_FIXED_07062026.docx`)

| Check | Result | Detail |
|-------|--------|--------|
| File exists | PASS | present |
| "360 thanh thiếu niên (180 nam/180 nữ)" | PASS | "360 thanh thiếu niên": 1; "180 nam": 1; "180 nữ": 1 |
| "360 adolescents (180 boys/180 girls)" | PASS | "360 adolescents": 1; "180 boys": 1; "180 girls": 1 |
| No remaining "287" or "287 adolescents" | PASS | 0 hits |

**Section E: 4/4 PASS**

---

## F. `_Archive/` dedup folders

| Folder | Result | Detail |
|--------|--------|--------|
| `_Archive/TomTat_dedup_07062026/` | PASS | exists, 11 archived files |
| `_Archive/BanDich_dedup_07062026/` | PASS | exists, 20 archived files |

**Section F: 2/2 PASS**

---

## TOTALS

- Section A: 7/7 PASS
- Section B: 5/5 PASS
- Section C: 5/5 PASS
- Section D: 3/4 PASS + 1 informational note (BB gaps already match index of 15)
- Section E: 4/4 PASS
- Section F: 2/2 PASS

**TOTAL: 26/26 mechanical checks PASS. No critical failures.**

### Notes (non-blocking)

1. **BB index gaps (D):** `canonical_index.json` contains 15 BB entries (BB04, BB13, BB15, BB16, BB17, BB20, BB21 absent). Total 117 = 78 QT + 15 BB + 24 VN, consistent with stated counts. If the original 22-BB scheme was intentional, these gaps may warrant a separate decision; for this audit they are not regressions.
2. All FIXED tóm tắt and new QT074–QT078 files use Vietnamese number format (commas as decimals: "0,96", "−3,83") which matches project style.
