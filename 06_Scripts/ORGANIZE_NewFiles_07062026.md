# ORGANIZE NewFiles — 07/06/2026 PM session

Working dir: `C:\Users\OS\OneDrive\read_books\Lo-au\`
Canonical index backup: `02_Papers-goc/canonical_index.json.bak_NewFiles_07062026`
Total entries after this batch: **131** (was 124 → +7 new QT084-QT090)

---

## 1. Grummitt 2025 (QT082) — DONE

- Source: `02_Papers-goc/main.pdf`
- Destination: `02_Papers-goc/The-gioi_Au-My-Uc/QT082_Grummitt_2025_OurFutures_Australia_eClinMed.pdf`
- Magic bytes: `%PDF-1.7` confirmed
- Canonical updated: pdf_path set; full author list expanded (14 authors); `pdf_download_status` flipped FAILED → DOWNLOADED 07/06/2026
- `verification_status_pdf`: page 1 verified directly — N=1785 (Year 8/9, 10 schools NSW Australia; recruited Dec 2022-Apr 2024); anxiety β=-1.05 95%CI[-1.93,-0.12] p=0.024; depression β=-0.94 95%CI[-1.88,0.04] p=0.055 NS; ACTRN12622001582741; Paul Ramsay Foundation; CC BY 4.0

## 2. Kieling 2024 (QT079) — DONE (HTML)

- Source: `02_Papers-goc/Worldwide Prevalence...PMC.html` + `_files/`
- Destination:
  - `02_Papers-goc/The-gioi_Au-My-Uc/QT079_Kieling_2024_Worldwide_GBD_JAMAPsych.html`
  - `02_Papers-goc/The-gioi_Au-My-Uc/QT079_Kieling_2024_Worldwide_GBD_JAMAPsych_files/`
- Canonical updated: `pdf` + `pdf_path` set to HTML; `format` field added; content verified — first author Christian Kieling (UFRGS, Brazil), JAMA Psychiatry 81(4):347-356, PMC10831630

## 3. Lancet Psychiatry zip extraction — 9 PDFs → 7 canonical + 2 archived

Source: `02_Papers-goc/_unzip_lancet/` (extracted from `TheLancet.com_20260606.zip`). All 9 magic bytes `%PDF-1.7` confirmed.

| QT  | File (renamed)                                             | Paper                                                                           | Type            | Issue            |
| --- | ---------------------------------------------------------- | ------------------------------------------------------------------------------- | --------------- | ---------------- |
| QT084 | `QT084_Ostinelli_2025_ADHD_Adults_NMA_LancetPsych.pdf`   | Ostinelli et al. — Adult ADHD pharm/psych/neurostim component NMA (113 RCTs)    | SR + NMA        | LP 12(1):32-43   |
| QT085 | `QT085_Kohrt_2025_EQUIP_WHO_UNICEF_LancetPsych.pdf`      | Kohrt et al. — WHO-UNICEF EQUIP competency framework (DOI ...00183-4)           | Health Policy   | LP 12(1):67-80   |
| QT086 | `QT086_Fineberg_2024_PUI_Commission_LancetPsych.pdf`     | Fineberg/Chamberlain — Lancet Psych PUI Commission announcement (DOI ...00323-7) | Comment        | LP 12(1):11      |
| QT087 | `QT087_Zavlis_2024_PersonalityDisorder_RCT_LancetPsych.pdf` | Zavlis/Luyten/Pilling/Fonagy — call for RCTs on dimensional PD (DOI ...00365-1) | Correspondence | LP 12(1):14      |
| QT088 | `QT088_Editorial_2025_Psychedelics_Validity_LancetPsych.pdf` | Lancet Psych Editorial — Seeking validity with psychedelics                  | Editorial       | LP 12(1):1       |
| QT089 | `QT089_Augustinavicius_2024_Climate_COP29_LancetPsych.pdf` | Multi-author Comment — mental health × COP29 (DOI ...00375-4)                  | Comment         | LP 12(1):9       |
| QT090 | `QT090_Zhang_2025_CoerciveHospitalisation_China_LancetPsych.pdf` | Zhang/Suo/Gao — coercive hospitalisation in China                          | Correspondence  | LP 12(1):14-15   |

**Archived (low value — Corrections):**
- `_Archive/Lancet_corrections_07062026/Correction_Santomauro_LancetPsych_2024_11_1012-21.pdf`
- `_Archive/Lancet_corrections_07062026/Correction_Zavlis_LancetPsych_2024_Nov21.pdf`

**Relevance to LA/Q1:**
- **QT084 (Ostinelli ADHD NMA)** — most useful as NCS/comorbidity context
- **QT085 (EQUIP)** — useful for task-sharing/non-specialist training discussion (school counsellor capacity in VN)
- **QT086 (PUI Commission)** — useful for screen-time / social media subsection
- QT087-QT090 — off-scope, kept for completeness

## 4. Zip + temp folder cleanup

- `02_Papers-goc/TheLancet.com_20260606.zip` → `_Archive/zips_07062026/` ✅
- `02_Papers-goc/_unzip_lancet/` — **emptied but folder NOT removed** (OneDrive sync lock — "in use"). All 9 PDFs successfully moved out; only the empty directory shell remains. User can delete manually after OneDrive releases the handle, or it will be removed on next reboot.

## 5. Robson 2025 (QT081 CATS) — NOT FOUND

Status: **NOT in zip**. User needs to download manually.

Suggested URLs (already in QT081 canonical):
- https://pubmed.ncbi.nlm.nih.gov/39644904/
- https://www.thelancet.com/journals/lanpsy/article/PIIS2215-0366(24)00361-4/fulltext
- https://doi.org/10.1016/S2215-0366(24)00361-4
- ResearchGate search: "Robson Husin Dashti CATS trajectories adolescence"

Note: Lancet Psych is paywalled — institutional access likely required.

---

**Files moved:** 12 (Grummitt PDF, Kieling HTML + assets folder, 7 Lancet PDFs to The-gioi_Au-My-Uc, 2 Corrections to _Archive, 1 zip to _Archive)
**Canonical entries added:** 7 (QT084-QT090)
**Canonical entries updated:** 2 (QT079 + QT082, both PDF status flipped)
**JSON validates:** Python json.load OK, 131 total entries
