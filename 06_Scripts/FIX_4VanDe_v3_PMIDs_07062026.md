# FIX REPORT — `4VanDe_BLOCKING_Q1Q3_v3_07062026.docx` PMID/DOI corrections

**Date:** 07/06/2026
**Trigger:** `06_Scripts/AUDIT_4VanDe_v3_FINAL_07062026.md` — FAIL 2/6 refs (Karasu PMID + Pezzella PMID/DOI/issue đều SAI)
**File fixed:** `bai-bao-Q1/4VanDe_BLOCKING_Q1Q3_v3_07062026.docx`
**Script touched:** `06_Scripts/sinh_4VanDe_v3_07062026.py`

---

## 1. Diff before/after

### Ref [4] — Karasu (FULL rewrite: bài cũ → bài đúng context)

**BEFORE (v3 build đầu):**
```
[4] Karasu, T. B. (1986). The specificity versus nonspecificity dilemma:
Toward identifying therapeutic change agents. American Journal of
Psychiatry, 143(6), 687-695. DOI: 10.1176/ajp.143.6.687. PMID: 3717388.
```

**Vấn đề:**
- PMID `3717388` thực ra = bài "Ionic requirements of peritubular taurine transport in Fundulus kidney" (Wolff et al., Am J Physiol 250(6):R984-90) — hoàn toàn không liên quan tâm thần.
- Title "specificity vs nonspecificity dilemma" có PMID đúng là `3717390`, nhưng context user yêu cầu (đã verify trong `Verify_BaoVeLA v8`) là bài **AJ Psychotherapy "Benefits and Limitations"**.

**AFTER (đã fix):**
```
[4] Karasu, T. B. (1986). The psychotherapies: Benefits and limitations.
American Journal of Psychotherapy, 40(3), 324-342. DOI:
10.1176/appi.psychotherapy.1986.40.3.324. PMID: 3094389.
```

### Ref [6] — Pezzella (3 trường sai cùng lúc)

**BEFORE:**
```
[6] Pezzella, P. (2022). The ICD-11 is now officially in effect.
World Psychiatry, 21(3), 331-332. DOI: 10.1002/wps.21038. PMID: 36073678.
```

**Vấn đề:**
- PMID `36073678` = bài Krueger RF "Incremental integration of nosological innovations..." 21(3):416-417 (sai tác giả).
- DOI `10.1002/wps.21038` = sai bài.
- Issue `21(3)` = sai (đúng là `21(2)`).

**AFTER (đã fix):**
```
[6] Pezzella, P. (2022). The ICD-11 is now officially in effect.
World Psychiatry, 21(2), 331-332. DOI: 10.1002/wps.20982. PMID: 35524598.
```

### Footer (diacritic polish)

**BEFORE:** `Soan 07/06/2026`
**AFTER:** `Soạn 07/06/2026`

---

## 2. Verification PubMed/Crossref

### Karasu 1986 "Benefits and Limitations" — PMID 3094389

- PubMed URL: `https://pubmed.ncbi.nlm.nih.gov/3094389/`
- Bibliographic data (PubMed):
  - Title: "The psychotherapies: Benefits and limitations"
  - Author: Karasu TB
  - Journal: Am J Psychother. 1986 Jul;40(3):324-42.
  - DOI: 10.1176/appi.psychotherapy.1986.40.3.324
- Crossref check (DOI `10.1176/appi.psychotherapy.1986.40.3.324`): khớp journal Am J Psychotherapy vol 40 issue 3 pages 324-342, year 1986, author Karasu T.B.
- **Verdict:** ✓ KHỚP HOÀN TOÀN.

### Pezzella 2022 "ICD-11 officially in effect" — PMID 35524598

- PubMed URL: `https://pubmed.ncbi.nlm.nih.gov/35524598/`
- Bibliographic data:
  - Title: "The ICD-11 is now officially in effect"
  - Author: Pezzella P
  - Journal: World Psychiatry. 2022 Jun;21(2):331-332.
  - DOI: 10.1002/wps.20982
  - PMCID: PMC9077596
- Crossref check (DOI `10.1002/wps.20982`): khớp tác giả Pezzella P., journal World Psychiatry, volume 21, issue 2, pages 331-332, year 2022.
- **Verdict:** ✓ KHỚP HOÀN TOÀN cả PMID, DOI, issue, pages.

### Sanity check — PMID sai cũ KHÔNG còn xuất hiện trong file

| Old (wrong) | grep hits sau rebuild |
|---|---|
| PMID 3717388 | 0 |
| PMID 36073678 | 0 |
| DOI 10.1002/wps.21038 | 0 |
| DOI 10.1176/ajp.143.6.687 | 0 |

| New (correct) | grep hits sau rebuild |
|---|---|
| PMID 3094389 | 1 |
| PMID 35524598 | 1 |
| DOI 10.1002/wps.20982 | 1 |
| DOI 10.1176/appi.psychotherapy.1986.40.3.324 | 1 |
| Footer "Soạn 07/06/2026" (có dấu) | 1 |

---

## 3. File stats after fix

| Item | Value |
|---|---|
| Path | `bai-bao-Q1/4VanDe_BLOCKING_Q1Q3_v3_07062026.docx` |
| Size | 40 997 bytes |
| Word count | 1 596 từ (vẫn trong target 1.500–2.000; giảm 5 từ vs v3 đầu do tiêu đề bài Karasu mới ngắn hơn) |
| Metadata stripped | author='', title='', keywords='', subject='', last_modified_by='' (verified) |
| Footer | "Soạn 07/06/2026" có diacritic |

---

## 4. Refs KHÔNG bị đụng tới (xác nhận)

- [1] Frontiers in Psychiatry guidelines — unchanged
- [2] BMC Public Health guidelines — unchanged
- [3] Frontiers in Public Health guidelines — unchanged
- [5] First et al. 2022 DSM-5-TR PMID 35524596 / DOI 10.1002/wps.20989 — unchanged (audit đã PASS)
- [7] Nội bộ v2_01062026 — unchanged
- [8] Nội bộ tin nhắn thầy MĐ 07/06 — unchanged

---

## 5. Tóm tắt

| Check | Result |
|---|---|
| Ref [4] Karasu PMID `3094389` + DOI `10.1176/appi.psychotherapy.1986.40.3.324` | ✓ FIXED |
| Ref [6] Pezzella PMID `35524598` + DOI `10.1002/wps.20982` + issue `21(2)` | ✓ FIXED |
| Footer "Soạn" có diacritic | ✓ FIXED |
| Bad PMIDs cũ (3717388, 36073678) còn trong file | ✗ 0 hits — clean |
| Metadata stripped | ✓ CONFIRMED |
| Word count trong target | ✓ 1 596 từ |
| Refs khác bị đụng | ✗ KHÔNG — chỉ [4], [6], footer |

**Trạng thái:** sẵn sàng gửi thầy MĐ. Audit re-run đề xuất chạy lại toàn bộ 7 check trong `AUDIT_4VanDe_v3_FINAL_07062026.md` để xác nhận `6 PASS + 1 FIXED → 7 PASS`.
