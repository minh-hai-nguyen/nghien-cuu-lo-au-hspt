# FINAL AUDIT — Afternoon Session 07/06/2026 (Q1 → Q2 pivot)

**Mode:** READ-ONLY mechanical re-verification of 3 files built in afternoon session 07/06/2026.
**Companion to:** `06_Scripts/FINAL_AUDIT_07062026.md` (morning session, 26/26 PASS).
**Scope:** 23 mechanical checks across 7 sections (target ≥ 20).

Files audited:
- `bai-bao-Q1/Draft_Q2_SongNgu_v1_07062026.docx` (59,018 bytes)
- `bai-bao-Q1/PhuongAnXuLy_Q1toQ2_07062026.docx` (48,442 bytes)
- `bai-bao-Q1/4VanDe_BLOCKING_Q1Q3_v3_07062026.docx` (41,011 bytes)

---

## Section A — Files exist + valid (4/4 PASS)

| # | Check | Result | Detail |
|---|-------|--------|--------|
| A1 | Draft_Q2_SongNgu_v1 exists + >50KB | PASS | size = 59,018 bytes |
| A2 | PhuongAnXuLy exists + >30KB | PASS | size = 48,442 bytes |
| A3 | 4VanDe v3 exists + >20KB | PASS | size = 41,011 bytes |
| A4 | All 3 readable as valid .docx (python-docx) | PASS | all 3 open without error |

## Section B — Metadata stripped (3/3 PASS)

| # | Check | Result | Detail |
|---|-------|--------|--------|
| B1 | Draft_Q2_SongNgu_v1 core.xml empty | PASS | creator/lastModifiedBy/title/subject/keywords/description all empty |
| B2 | PhuongAnXuLy core.xml empty | PASS | all 6 fields empty |
| B3 | 4VanDe v3 core.xml empty | PASS | all 6 fields empty |

## Section C — Footers / date stamp (2/3 PASS, 1 informational)

| # | Check | Result | Detail |
|---|-------|--------|--------|
| C1 | Draft_Q2 v1: footer/end has "v1" or "07/06/2026" | PASS | "v1" present (footer has "Phien 07/06/2026") |
| C2 | PhuongAnXuLy: has "Soạn 07/06/2026" | PASS | found in footer XML (`docProps/footer*.xml`) |
| C3 | 4VanDe v3: has "Soạn 07/06/2026" | INFO/FAIL | exact string "Soạn 07/06/2026" not present; but last body paragraph reads `"Soan 07/06/2026"` (no diacritics) — date stamp is there, just in non-diacritic form. Body otherwise has 8 "07/06/2026" hits including title "BAN v3 CAP NHAT NGAY 07/06/2026". Non-blocking. |

## Section D — Q2 v1 critical content (5/5 PASS)

| # | Check | Result | Detail |
|---|-------|--------|--------|
| D1 | Title contains "cross-sectional structural equation modeling" | PASS | 1 hit |
| D2 | No "mixed-methods" / "qualitative interviews" / "thematic analysis" / "joint-display" | PASS | all 5 banned terms = 0 hits |
| D3 | Section 4.4 contains "co-rumination" + "tam giao" | PASS | both present in section 4.4 block (35,202 chars) |
| D4 | Refs include Rose 2002 + Stankov 2010 + Small & Blanc 2021 | PASS | all 3 present |
| D5 | Niwenahisemo: "Hong, S" present, "Su, H" absent | PASS | Hong present, Su absent |

## Section E — PhuongAnXuLy content (3/3 PASS)

| # | Check | Result | Detail |
|---|-------|--------|--------|
| E1 | Mentions Frontiers in Psychiatry as TOP | PASS | 16 hits |
| E2 | Mentions BMC Public Health + Frontiers in Public Health as backup | PASS | both present |
| E3 | Word count 3,000–4,000 | PASS | 3,980 words |

## Section F — 4VanDe v3 content (3/3 PASS)

| # | Check | Result | Detail |
|---|-------|--------|--------|
| F1 | Mentions all 4 issues (Q1-6, Q1-8, Q3-6, Q3-9) | PASS | hits: {Q1-6:1, Q1-8:1, Q3-6:2, Q3-9:1} |
| F2 | Mentions Q1→Q2 + Q3 re-design decisions | PASS | both keywords present |
| F3 | Action items table 5 columns visible | PASS | table 2 has 5 cols × 10 rows |

## Section G — References cross-validation (0/2 PASS — CRITICAL)

| # | Check | Result | Detail |
|---|-------|--------|--------|
| G1 | Karasu PMID 3094389 correct (no 3717390 / 3717388) | **FAIL** | 4VanDe v3 ref [4] cites `PMID: 3717388` — wrong: verified as fish-kidney physiology paper (Wolff et al., Am J Physiol 250(6):R984–90). Correct PMIDs: 3094389 (Benefits/Limitations) or 3717390 (specificity vs nonspecificity). Already flagged in `06_Scripts/AUDIT_4VanDe_v3_FINAL_07062026.md`. |
| G2 | Sample 5 refs in Q2 v1 each have DOI/PMID/URL | **FAIL** | 3/5 missing identifier in Draft_Q2 v1: McLean et al. 2011, Compas et al. 2017, Anderson et al. 2025 — only journal/title given, no DOI/PMID/URL. 2/5 OK: Small & Blanc 2021, Chen et al. 2023. |

---

## TOTALS

- Section A: 4/4 PASS
- Section B: 3/3 PASS
- Section C: 2/3 PASS (+ 1 informational, non-blocking)
- Section D: 5/5 PASS
- Section E: 3/3 PASS
- Section F: 3/3 PASS
- Section G: 0/2 PASS — **2 critical findings**

**Total checks run: 23 (target ≥ 20 met).**
**PASS: 20/23 (87%). FAIL: 3 (1 informational + 2 critical).**

---

## Comparison vs morning session

| | Morning (`FINAL_AUDIT_07062026.md`) | Afternoon (this report) |
|---|---|---|
| Total checks | 26 | 23 |
| PASS | 26 | 20 |
| FAIL | 0 | 3 (1 informational, 2 critical) |
| Critical issues | None | 2 (G1 Karasu PMID, G2 Q2 v1 refs missing DOI/PMID/URL) |
| Files cleared to send | Yes | **NO — fix G1 + G2 first** |

---

## Critical issues (must fix before sending)

1. **G1 — Karasu PMID error in `4VanDe_BLOCKING_Q1Q3_v3_07062026.docx`.**
   Ref [4] currently reads: `Karasu, T. B. (1986)… American Journal of Psychiatry, 143(6), 687-695. DOI: 10.1176/ajp.143.6.687. PMID: 3717388.`
   PMID 3717388 = fish-kidney paper (NOT Karasu). Fix to:
   - PMID **3717390** if keeping "specificity vs nonspecificity" (matches the title/journal already in the line); OR
   - PMID **3094389** + change citation to Karasu 1986 "Benefits and Limitations" Am J Psychotherapy 40(3):324-342 (matches canonical_index.json entry and morning session QT074).
   This was already documented in `06_Scripts/AUDIT_4VanDe_v3_FINAL_07062026.md` but the file was not patched after that earlier audit.

2. **G2 — 3 references in `Draft_Q2_SongNgu_v1_07062026.docx` lack DOI/PMID/URL identifiers:**
   - McLean, C. P., et al. (2011). Gender differences in anxiety disorders…
   - Compas, B. E., Jaser, S. S., Bettis, A. H., et al. (2017). Coping, emotion regulation… *Psychological Bulletin*, 143(9), 939–991.
   - Anderson, E. M., et al. (2025). Narrative review of adolescent anxiety…
   Per user memory `feedback_doc_phai_co_reference.md` ("Mọi doc trả lời thầy phải có DOI/PMID/URL kiểm chứng"), these must be supplemented before sending. Also worth verifying Anderson 2025 actually exists (recent year + "et al." + no identifier is a common fabrication red flag — cross-check against memory `project_bai2_khgdvn_audit.md` finding "17/21 TLTK sai/bịa").

## Informational (non-blocking)

3. **C3 — `4VanDe v3` ends with `Soan 07/06/2026` (no diacritics)** instead of `Soạn 07/06/2026`. Date stamp is functionally present; recommend normalising to diacritic form for consistency with PhuongAnXuLy/Q2 v1 styling, but not blocking.

---

## Final recommendation

**NOT ready to send thầy MĐ + chị CH + thầy NMĐ as-is.** Two critical fixes required:

1. Patch ref [4] in `4VanDe_BLOCKING_Q1Q3_v3_07062026.docx` to use correct Karasu PMID (3717390 or 3094389 per author decision).
2. Add DOI/PMID/URL to McLean 2011, Compas 2017, Anderson 2025 in `Draft_Q2_SongNgu_v1_07062026.docx` — **AND verify Anderson 2025 is a real paper** (high fabrication-risk signature).

After those 2 fixes, re-run sections G1 + G2 (and optionally normalise the C3 diacritics in 4VanDe v3) → then files are clear.

Morning session (Q1 v6 + 5 FIXED tóm tắt + 5 new QT074–QT078 + canonical_index.json + Outline + dedup archives) remains 26/26 PASS and is safe to send.

---

*Audit script: `_audit_tmp/afternoon_audit_07062026.py` (read-only, deterministic, reproducible).*
*Raw output: `_audit_tmp/afternoon_audit_raw.txt`.*
