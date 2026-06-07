# VERIFY — Polish Fact-Check Q2 v1 vs Q1 v7 (07/06/2026)

## 1. Scope

- Source A (Q1 v7, baseline): `bai-bao-Q1/Draft_Q1_SongNgu_v7_01062026.docx` — 212 paragraphs, 54 166 chars, mixed-methods version.
- Source B (Q2 v1, polished, target): `bai-bao-Q1/Draft_Q2_SongNgu_v1_07062026.docx` — 223 paragraphs, 62 076 chars, Frontiers in Psychiatry (quant-only SEM) version with 3 polish rounds applied.
- Method: read-only `python-docx` extraction; regex-based fact harvest; diff; specific anti-fabrication ref checks; bilingual spot-check.
- Scripts produced: `06_Scripts/_verify_q2polish_extract.py`, `_verify_q2polish_deeper.py`, `_verify_q2polish_finalize.py` (+ JSON outputs).

## 2. Facts/numbers extracted from Q2 v1

| Category | Count of distinct values |
|---|---|
| Sample sizes (N=…) | 3 distinct strings (1,352 / 1.352 + "1,352 participants" mention) |
| β coefficients | 13 distinct (full set match Q1v7) |
| p-values | 9 distinct |
| R² values | 11 distinct |
| Citation years | 23 distinct (1965–2025) |
| Percentages | 16 distinct |
| DOIs | 6 (vs 2 in Q1v7) |
| PMIDs | 4 (vs 1 in Q1v7) |
| Authors cited | 35 ref entries |

## 3. Fact-by-fact match Q2 v1 vs Q1 v7 (occurrence counts)

| Fact | Q1v7 | Q2v1 | Verdict |
|---|---:|---:|---|
| N=1,352 (EN) | 3 | 3 | MATCH |
| N=1.352 (VN) | 3 | 3 | MATCH |
| 614 boys | 6 | 6 | MATCH |
| 738 girls | 6 | 6 | MATCH |
| 45.4% / 45,4% | 4 | 4 | MATCH |
| 54.6% / 54,6% | 4 | 4 | MATCH |
| β=0.376 (SAD←bullying) | 7 | 7 | MATCH |
| β=0.510 (RLLA←self-esteem) | 7 | 7 | MATCH |
| β=0.669 (integrated) | 4 | 4 | MATCH |
| β=0.215 (GAD←bullying) | 5 | 5 | MATCH |
| β=0.253 (SocAD←bullying) | 8 | 8 | MATCH |
| GAD F=44.48 | yes | yes | MATCH |
| SocAD F=45.98 | yes | yes | MATCH |
| SAD F=0.246 (gender invariance) | yes | yes | MATCH |
| GAD female M=59.47 / male M=51.43 | yes | yes | MATCH |
| SocAD female M=52.74 / male M=43.20 | yes | yes | MATCH |
| SAD female M=24.76 / male M=25.42 | yes | yes | MATCH |
| Cronbach α ≥ .70 | both EN/VN | both EN/VN | MATCH |
| CFI ≥ .90, RMSEA ≤ .08 | both | both | MATCH |
| ASEAN prevalence 11.9% / VN 10.1% | both | both | MATCH |
| Grade ns: 368, 316, 340, 328 | 4 each | 4 each | MATCH |

Diff list from regex sweep — all are **format-only**, no value changes:
- `p<0.001` → `p<.001` (APA-style leading-zero drop). Applied uniformly to all 7 distinct p-strings.
- `R²=0,024 / 0,055 / 0,076 / 0,160 / 0,284` newly appear as `0.024 / 0.055 / 0.076 / 0.160 / 0.284` in Q2v1 EN sections (Q1v7 had only the VN-comma forms for these five — polish added the missing EN-decimal forms, so all 11 R² values now appear in both EN-dot and VN-comma forms as appropriate). No magnitude changed.
- `N_values` extra entry `"8.389"` in Q1v7 only is a regex artefact (the inline `8.389` referred to a footer label, not a sample size); does not exist in Q2v1 because that boilerplate was removed.
- Years `2006` (Braun & Clarke) and `2018` (Creswell & Plano Clark) absent from Q2v1 — these are mixed-methods/qualitative references **deliberately removed** because Q2 v1 is the quant-only Frontiers in Psychiatry recast (mixed-methods=0, qualitative=1, Braun=0, Creswell=0 in Q2v1 vs 9/12/4/3 in Q1v7). This is scope change, NOT polish drift.

## 4. Anti-fabrication specific ref checks (Q2 v1)

| Ref | Required | Found | Status |
|---|---|---|---|
| Niwenahisemo first co-author | "Hong, S" | "Hong S" (numbered ref style) | OK |
| Niwenahisemo wrong form | NOT "Su, H" | not present | OK |
| Anderson first author | "T. L." / "TL" | "Anderson TL, Valiauga R, Tallo C, et al." | OK |
| Anderson DOI | 10.1111/jcap.70009 | present | OK |
| Anderson PMID | 39739929 | present | OK |
| Rose 2002 PMID | 12487497 | present | OK |
| Stankov DOI | 10.1016/j.lindif.2010.05.003 | present | OK |
| Small & Blanc DOI | 10.3389/fpsyt.2020.589618 | present | OK |
| Karasu PMID 3094389 | optional (not cited) | not cited in Q2v1 | OK (Karasu was never in Q1v7 either) |
| Total DOIs in Refs | ≥ 4 | 6 | UP from Q1v7 (2) |
| Total PMIDs in Refs | ≥ 4 | 4 | UP from Q1v7 (1) |

Polish round actually **strengthened** verifiability (more DOIs/PMIDs surfaced into the reference list, none introduced incorrectly).

## 5. Bilingual EN/VN spot-checks (5 random pairs)

| # | EN excerpt | VN excerpt | Shared numbers verified | Status |
|---|---|---|---|---|
| 1 | Abstract: "1,352 lower secondary…614 boys, 738 girls; 11–14…AMOS 31.0" | "1.352 học sinh…614 nam, 738 nữ; 11–14…AMOS 31.0" | 614, 738, 11, 14, 31.0, 1352 | MATCH |
| 2 | Results: "school bullying…strongest with SAD (β=0.376, p<.001)…GAD (β=0.215)…SocAD (β=0.253)" | "bắt nạt…mạnh nhất với SAD (β=0,376; p<0,001)…GAD (β=0,215) hoặc SocAD (β=0,253)" | 0.376/0,376, 0.215/0,215, 0.253/0,253 | MATCH |
| 3 | Intro: "approximately 11.9% of adolescents in ASEAN and 10.1% in Vietnam…1990 and 2021" | "khoảng 11,9% thanh thiếu niên khu vực ASEAN và 10,1% tại Việt Nam…1990 và 2021" | 11.9/11,9, 10.1/10,1, 1990, 2021 | MATCH |
| 4 | Section 4.4 EN/VN heading: "Peer support null effects: co-rumination hypothesis" / "Hỗ trợ bạn bè không có tác động: giả thuyết co-rumination" | identical structure | n/a | MATCH |
| 5 | Discussion §3: "peer support…β = −0.044, p = .183, R² = 0.002…GAD (β = −0.015)…SAD (β = −0.019)…SocAD (β = −0.079, p = .020)" | "β = −0,044; p = 0,183; R² = 0,002…GAD (β = −0,015)…SAD (β = −0,019)…SocAD (β = −0,079; p = 0,020)" | all 6 βs and 2 p-values | MATCH |

No bilingual mismatch detected on numeric facts.

## 6. Round 2 anti-plagiarism paraphrase check (Rose 2002 co-rumination definition)

- Q1 v7 definition (EN): "defined as extensively discussing problems, speculating about them and focusing on negative feelings [Rose 2002]"
- Q2 v1 polished EN: "a perseverative form of dyadic problem-talk in which adolescents repeatedly revisit stressful events and their accompanying negative affect rather than moving toward resolution (34)"
- Q2 v1 polished VN: "một dạng trò chuyện cặp đôi mang tính lặp lại trong đó thanh thiếu niên liên tục quay lại các sự kiện gây stress và các cảm xúc tiêu cực kèm theo thay vì hướng tới giải quyết (34)"

Paraphrase preserved semantic content (repetitive dyadic discussion of problems with focus on negative affect). Citation still pinned to Rose 2002 (numbered ref 34, PMID 12487497). No fact distortion.

## 7. Drift introduced by polish

- **Numeric values**: 0 drift. All sample sizes, β, p, R², F, means, prevalence, fit indices, percentages match exactly across both languages.
- **Format**: APA-style `p < .001` (leading-zero dropped) applied uniformly — improvement.
- **Citation upgrades**: +4 DOIs and +3 PMIDs surfaced into the reference list — improvement, not drift.
- **Scope reduction (deliberate, not drift)**: mixed-methods/qualitative content (Creswell & Plano Clark 2018, Braun & Clarke 2006, joint-display matrix, interview protocol) removed because Q2 v1 is repositioned as the quant-only Frontiers in Psychiatry recast.
- **Author names / years in retained refs**: 0 drift. Niwenahisemo "Hong S" preserved (not corrupted to "Su H"); Anderson TL first-author preserved; Rose 2002, Stankov 2010, Small & Blanc 2021 all intact with correct DOIs/PMIDs.

## 8. Final verdict

**POLISH SAFE — facts NOT disturbed.**

All quantitative facts (N=1,352, 614 M / 738 F, 45.4% / 54.6%, all β / p / R² / F / fit indices / prevalence / Grade ns / Cronbach thresholds) reproduce bit-exactly between Q1 v7 and Q2 v1 in both English and Vietnamese. The only differences are (a) APA leading-zero p-value formatting, (b) **added** DOIs/PMIDs (strengthens verifiability), (c) anti-plagiarism paraphrase of Rose 2002 definition with citation preserved, and (d) deliberate scope reduction (mixed-methods → quant-only) which removed Braun & Clarke 2006 and Creswell & Plano Clark 2018 references — both deliberate, not polish artefacts.

No bilingual mismatch found in 5 random EN/VN spot-checks. Polish session can be archived as PASS.

## 9. Artefacts

- `06_Scripts/_verify_q2polish_extract.py` / `.json` — full fact harvest.
- `06_Scripts/_verify_q2polish_deeper.py` / `.json` — line-level context for drift candidates.
- `06_Scripts/_verify_q2polish_finalize.py` / `.json` — occurrence-count table.

## 10. References (verification sources)

- Q1 v7 baseline draft (project file): `bai-bao-Q1/Draft_Q1_SongNgu_v7_01062026.docx` (commit-equivalent: file on disk 01/06/2026).
- Q2 v1 polished draft (project file): `bai-bao-Q1/Draft_Q2_SongNgu_v1_07062026.docx` (07/06/2026).
- Rose 2002 (co-rumination): PMID 12487497 — Child Development 73(6) 1830–1843.
- Niwenahisemo, Hong & Kuang 2024 (GAD-7 invariance Rwanda): Frontiers in Psychiatry 2024.
- Anderson TL et al. 2025 (adolescent anxiety narrative review): DOI 10.1111/jcap.70009, PMID 39739929.
- Stankov 2010 (Confucian academic culture): DOI 10.1016/j.lindif.2010.05.003.
- Small & Blanc 2021 (peer disclosure norms East Asia): DOI 10.3389/fpsyt.2020.589618.
- APA p-value formatting standard: APA Publication Manual 7th ed §6.44 (leading zero dropped when value cannot exceed 1).
