# VERIFY Q1 v5 — 7 Recent Empirical Refs (External DOI/PubMed)

**Date:** 2026-06-07
**Cross-check source:** `06_Scripts/CROSSCHECK_Q1v5_Refs_07062026.md` (Sub-group B HIGH RISK)
**Method:** WebFetch on Crossref/PubMed/PMC/journal homepages + WebSearch fallback. No local PDF dependency.
**Read-only:** No canonical_index.json modifications in this report — only suggested fixes for Q1 v6.

---

## Summary

| # | Ref | Status | DOI / PMID |
|---|---|---|---|
| 1 | Kieling C et al. 2024 | EXISTS_VERIFIED | 10.1001/jamapsychiatry.2023.5051 / PMID 38294785 |
| 2 | Steare T et al. 2023 | EXISTS_VERIFIED (already fact-checked) | 10.1016/j.jad.2023.07.028 |
| 3 | Robson EM et al. 2025 (CATS) | EXISTS_VERIFIED | 10.1016/S2215-0366(24)00361-4 / PMID 39644904 |
| 4 | Niwenahisemo LC et al. 2024 | EXISTS_BUT_WRONG_DETAIL — author name error | 10.3389/fpsyt.2024.1346267 / PMID 38528981 |
| 5 | Grummitt L, O'Dean S 2025 | EXISTS_VERIFIED | 10.1016/j.eclinm.2025.103672 / PMID 41399472 |
| 6 | WHO 2022 25% increase brief | EXISTS_VERIFIED | WHO news 02/03/2022 (not "scientific brief" — see note) |
| 7 | Hoàng Trung Học 2025 (N=8.473) | EXISTS_VERIFIED (already in canonical VN014 with PDF) | — VN014 |

**No fabricated refs detected.** All 7 papers exist. One has author-name error (#4); one needs document-type wording fix (#6).

---

## Per-ref detail

### 1. Kieling C et al. 2024 — VERIFIED
- **Title:** *Worldwide Prevalence and Disability From Mental Disorders Across Childhood and Adolescence*
- **Authors:** Kieling C, Buchweitz C, Caye A, Silvani J, Ameis SH, Brunoni AR, et al. (15 authors total)
- **Journal:** JAMA Psychiatry **81(4):347-356** (2024) ← Q1 v5 lists only "347" (start page). Suggest add 347-356.
- **DOI:** 10.1001/jamapsychiatry.2023.5051 | **PMID:** 38294785
- **Action Q1 v6:** Add ending page (-356). Add to canonical_index.

### 2. Steare T et al. 2023 — VERIFIED (already)
- Confirmed via standalone fact-check: `06_Scripts/FACTCHECK_Steare2023_07062026.md`
- **J Affect Disord 339:302-317**; DOI 10.1016/j.jad.2023.07.028
- **Action:** No change to Q1 wording needed (already vetted: narrative SR, 48/52 studies, n=52).

### 3. Robson EM, Husin HM, Dashti SG et al. 2025 — VERIFIED
- **Title:** *Tracking the course of depressive and anxiety symptoms across adolescence (the CATS study): a population-based cohort study in Australia*
- **Full author list:** Robson EM, Husin HM, Dashti SG, Vijayakumar N, Moreno-Betancur M, Moran P, Patton GC, Sawyer SM
- **Lancet Psychiatry 12(1):44-53** (Jan 2025) | **DOI:** 10.1016/S2215-0366(24)00361-4 | **PMID:** 39644904
- **Action Q1 v6:** Add issue number "12(1)". Add to canonical_index.

### 4. Niwenahisemo LC, Su H, Kuang L 2024 — WRONG AUTHOR NAME
- **Q1 v5 wrote:** "Niwenahisemo LC, **Su H**, Kuang L"
- **Correct (per Frontiers + ORCID 0000-0003-2294-5251):** "Niwenahisemo LC, **Hong S**, Kuang L"
  - Second author's full name is **Su Hong** (Chinese order: surname **Hong**, given **Su**). Correspondence email gracegirlhs@sina.com confirms surname Hong. In Vancouver style → "Hong S".
- **Title:** *Assessing anxiety symptom severity in Rwandese adolescents: cross-gender measurement invariance of GAD-7*
- **Front Psychiatry 15:1346267** (2024) | **DOI:** 10.3389/fpsyt.2024.1346267 | **PMID:** 38528981
- **Action Q1 v6:** **FIX author name "Su H" → "Hong S"**. Add to canonical_index.

### 5. Grummitt L, O'Dean S 2025 — VERIFIED
- **Title:** *Efficacy of a school-based, universal prevention programme for depression and anxiety in adolescents (OurFutures Mental Health): a two-arm cluster-randomised controlled trial*
- **Authors include third author:** Birrell L (Q1 v5 cites only first two — acceptable per "et al." convention, but flag if Q1 uses "Grummitt L, O'Dean S" without "et al.")
- **eClinicalMedicine 90:103672** (2025) | **DOI:** 10.1016/j.eclinm.2025.103672 | **PMID:** 41399472
- **Action Q1 v6:** Append ", Birrell L" or use "et al." Add vol/article number. Add to canonical_index.

### 6. WHO 2022 — VERIFIED but minor type wording
- **Exact title:** "COVID-19 pandemic triggers 25% increase in prevalence of anxiety and depression worldwide"
- **Type:** WHO **News release**, 02/03/2022 — it *references* an underlying scientific brief titled *"Mental health and COVID-19: Early evidence of the pandemic's impact"* (the brief itself is the citable source for the 25% figure).
- **URL:** https://www.who.int/news/item/02-03-2022-covid-19-pandemic-triggers-25-increase-in-prevalence-of-anxiety-and-depression-worldwide
- **Action Q1 v6:** If Q1 cites the figure "25%", cite the **scientific brief** ("Mental health and COVID-19: Early evidence of the pandemic's impact", WHO, 02/03/2022) rather than the news item. The news item is acceptable as press release but the brief is the primary source.

### 7. Hoàng Trung Học et al. 2025 (N=8.473) — VERIFIED
- Already in canonical as **VN014** (`HoangTrungHoc_2025_VN_COVID`) with PDF at `02_Papers-goc/Viet-Nam/` + Vietnamese-language summary at `Tom-tat-tung-bai/VN014_HoangTrungHoc_2025_VN_COVID.docx`.
- N=8.473 adolescents grades 6-12 across 6 provinces (3 South + 3 North), DASS-21, COVID + post-COVID waves (verified from internal extract script `03_Ban-dich/_Archive_Scripts/tao_HoangTrungHoc2025.py`).
- **Note on cross-check report:** Sub-group B in `CROSSCHECK_Q1v5_Refs_07062026.md` is mistaken — this paper IS in canonical (VN014, line 412 of canonical_index.json). Treat as already SAFE.
- **Action Q1 v6:** No change.

---

## Aggregate verdict

- **0 fabricated** refs detected.
- **1 author-name error** (#4 Niwenahisemo — fix "Su H" → "Hong S" in Q1 v6 reference list AND any in-text "(Niwenahisemo, Su, & Kuang…)" mentions).
- **1 source-type clarification** (#6 WHO — cite scientific brief, not news release, if using the 25% statistic).
- **5 small completeness fixes** (end page for Kieling; issue number for Robson; vol/article number for Grummitt; third author or "et al." for Grummitt).
- **Action for canonical_index.json:** add entries for refs #1, #3, #4, #5 (PDFs to be downloaded — all open-access or Lancet PMC). Refs #2 and #7 already tracked. Ref #6 is grey literature (WHO) — store URL + PDF snapshot of brief.

## References for verification

- Crossref: https://api.crossref.org/works/10.1001/jamapsychiatry.2023.5051
- PubMed PMID 38294785, 39644904, 38528981, 41399472
- PMC: PMC10831630, PMC10962260, PMC12702296
- Frontiers Psychiatry article landing: https://www.frontiersin.org/journals/psychiatry/articles/10.3389/fpsyt.2024.1346267/full
- Lancet Psychiatry: https://www.thelancet.com/journals/lanpsy/article/PIIS2215-0366(24)00361-4/abstract
- eClinicalMedicine: https://www.thelancet.com/journals/eclinm/article/PIIS2589-5370(25)00606-6/fulltext
- WHO news: https://www.who.int/news/item/02-03-2022-covid-19-pandemic-triggers-25-increase-in-prevalence-of-anxiety-and-depression-worldwide
- Internal: `02_Papers-goc/canonical_index.json` (VN014), `06_Scripts/FACTCHECK_Steare2023_07062026.md`, `03_Ban-dich/_Archive_Scripts/tao_HoangTrungHoc2025.py`
