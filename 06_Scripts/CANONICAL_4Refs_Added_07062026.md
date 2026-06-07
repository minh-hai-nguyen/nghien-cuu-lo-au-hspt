# CANONICAL_INDEX — 4 Q1 v6 Refs Added (QT079-QT082)

**Date:** 2026-06-07
**Source verification:** `06_Scripts/VERIFY_Q1_7Refs_07062026.md` (PubMed/Crossref externally verified)
**Backup created:** `02_Papers-goc/canonical_index.json.bak_07062026_4refs`
**Total entries after update:** 121 (was 117)

---

## 1. New entries added to canonical_index.json

| QT | Descriptor | Journal | Year | DOI | PMID | PDF |
|---|---|---|---|---|---|---|
| QT079 | Kieling_2024_JAMA_Worldwide_Adolescent_GBD | JAMA Psychiatry 81(4):347-356 | 2024 | 10.1001/jamapsychiatry.2023.5051 | 38294785 | — (manual) |
| QT080 | Niwenahisemo_2024_Frontiers_GAD7_Rwanda | Frontiers in Psychiatry 15:1346267 | 2024 | 10.3389/fpsyt.2024.1346267 | 38528981 | DOWNLOADED |
| QT081 | Robson_2025_LancetPsychiatry_CATS_Australia | Lancet Psychiatry 12(1):44-53 | 2025 | 10.1016/S2215-0366(24)00361-4 | 39644904 | — (paywall) |
| QT082 | Grummitt_2025_eClinicalMedicine_OurFutures | eClinicalMedicine 90:103672 | 2025 | 10.1016/j.eclinm.2025.103672 | 41399472 | — (manual) |

**All four:** `verification_status = "EXTERNALLY VERIFIED via PubMed 07/06/2026"`, `used_for = "Q1 v6 reference"`, `added = "2026-06-07"`.

**Note on QT080 (Niwenahisemo):** Second author corrected from "Su H" (Q1 v5 wording) to "Hong S" per ORCID 0000-0003-2294-5251 and Frontiers correspondence — full name Su Hong (surname Hong).

---

## 2. PDF download attempts

### QT080 — Niwenahisemo (Frontiers OA) — DOWNLOADED
- URL: `https://www.frontiersin.org/articles/10.3389/fpsyt.2024.1346267/pdf`
- File: `02_Papers-goc/Chua-phan-loai/Niwenahisemo_2024_GAD7_Rwanda.pdf`
- Size: 618,354 bytes — magic bytes `%PDF-1.4` verified

### QT079 — Kieling (JAMA Psychiatry) — FAILED
- PMC10831630 exists but PMC returns HTML (anti-bot), Europe PMC returns connection-reset.
- Manual URLs:
  - https://pubmed.ncbi.nlm.nih.gov/38294785/
  - https://pmc.ncbi.nlm.nih.gov/articles/PMC10831630/
  - https://jamanetwork.com/journals/jamapsychiatry/fullarticle/2814712
- Note: PMC only lists supplementary PDFs (s001, s002) for free; main article PDF requires JAMA login or open via PMC web reader (manual browser).

### QT082 — Grummitt (eClinicalMedicine) — FAILED
- PMC12702296 exists. Lancet returns Cloudflare 403; PMC returns HTML; Europe PMC returns empty reply.
- Manual URLs:
  - https://pmc.ncbi.nlm.nih.gov/articles/PMC12702296/
  - https://www.thelancet.com/journals/eclinm/article/PIIS2589-5370(25)00606-6/fulltext
  - https://linkinghub.elsevier.com/retrieve/pii/S2589-5370(25)00606-6 (Elsevier marks as free OA resource)

### QT081 — Robson (Lancet Psychiatry) — SKIPPED (paywall, per task instructions)
- Manual URLs:
  - https://pubmed.ncbi.nlm.nih.gov/39644904/
  - https://www.thelancet.com/journals/lanpsy/article/PIIS2215-0366(24)00361-4/fulltext

---

## 3. Summary

- **4 entries added** (QT079-QT082) — JSON validated (121 total).
- **1 PDF successfully downloaded** (QT080 Niwenahisemo, 618 KB, valid magic bytes).
- **3 PDFs require manual download** (QT079, QT081, QT082) — all major endpoints (PMC, Europe PMC, Lancet/JAMA direct) blocked by anti-bot/Cloudflare or paywall.
- Backup of pre-edit canonical at `canonical_index.json.bak_07062026_4refs`.
