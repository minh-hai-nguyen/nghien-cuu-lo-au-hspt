# FIX QT002 — canonical_index.json (07/06/2026)

**Trigger:** `06_Scripts/AUDIT_Outlines_Q1Q3_07062026.md` — QT002 entry referenced non-existent PDF.

## Problem
- Entry `QT002` (Saikia 2023 India Assam) declared `pdf: QT002_Saikia_2023_India_Assam.pdf` which does NOT exist on disk.
- Actual Saikia 2023 IJCM PDF lives at: `02_Papers-goc/11-bai-ban-dau-va-mo-rong/11_Saikia_2023_IJCM.pdf`.
- A historical warning file `_MISLABELED_QT002_Saikia_actually_Bhardwaj_2020.pdf` exists in the same folder — that file is actually a Bhardwaj 2020 paper that was mistakenly named "Saikia" in a previous workflow.
- Outlines that cite `N=287` for Saikia 2023 are FABRICATED. The real Saikia 2023 IJCM paper reports `N=360 (180 male / 180 female)`.

## Old QT002 entry
```json
"QT002": {
  "id": "QT002",
  "descriptor": "Saikia_et_al_2023_India_Assam",
  "summary": "QT002_Saikia_et_al_2023_India_Assam.docx",
  "translation": "QT002_Saikia_2023_India_Assam.docx",
  "pdf": "QT002_Saikia_2023_India_Assam.pdf",
  "pdf_folder": "11-bai-ban-dau-va-mo-rong"
}
```

## New QT002 entry
```json
"QT002": {
  "id": "QT002",
  "descriptor": "Saikia_et_al_2023_India_Assam",
  "summary": "QT002_Saikia_et_al_2023_India_Assam.docx",
  "translation": "QT002_Saikia_2023_India_Assam.docx",
  "pdf": "11_Saikia_2023_IJCM.pdf",
  "pdf_folder": "11-bai-ban-dau-va-mo-rong",
  "pdf_path": "02_Papers-goc/11-bai-ban-dau-va-mo-rong/11_Saikia_2023_IJCM.pdf",
  "verification_note": "PDF path corrected 07/06/2026; sample N=360 (180m/180f) — outlines using N=287 are FABRICATED"
}
```

## Backup
- Original canonical preserved at: `02_Papers-goc/canonical_index.json.bak_QT002fix_07062026`
- Existing prior backups also retained: `.backup`, `.bak_07062026`, `.bak_20260607_112420`, `.bak_3refs_07062026`.

## Cross-checks performed on canonical_index.json

### Bhardwaj 2020 (separate entry confirmed)
- `QT011` Bhardwaj_et_al_2020_India_Chandigarh → `pdf: QT011_Bhardwaj_et_al_2020_India_Chandigarh.pdf` in `The-gioi_Khac/`. VERIFIED present on disk. No clash with QT002.
- `BB02` Bhardwaj_2020_India → intentionally points to `_MISLABELED_QT002_Saikia_actually_Bhardwaj_2020.pdf` as historical marker. KEPT per task constraint (do not delete warning file).

### Other entries scanned
- `QT055` Parade_Epigenetics_ChildhoodMaltreatment_TranslPsych_2021 — descriptor says "ChildhoodMaltreatment" but `pdf` correctly uses shorter filename `QT055_Parade_Epigenetics_Maltreatment_TranslPsych_2021.pdf`. File present in `The-gioi_Au-My-Uc/`. NOT a mislabel — just abbreviated filename.
- `QT057` NeuralCircuits_Anxiety_FrontNeuralCirc_2025 — PDF filename `QT057_NeuralCircuits_Mechanisms_Anxiety_FrontNeuralCirc_2025.pdf` exists in both `The-gioi_Khac/` (canonical folder) and `The-gioi_Au-My-Uc/`. Canonical entry consistent.
- `QT072` Lee_2025 and `QT073` He_2025 — `pdf_folder: 02_Papers-goc` (root). PDFs `QT072_Lee_2025_CyberbullyingMeta_TVA.pdf` and `QT073_He_2025_PowerUpCBTD_China_JAD.pdf` exist directly in `02_Papers-goc/` root. Path is unusual (root folder rather than subfolder) but consistent.

### No top-level metadata block
- `canonical_index.json` starts directly with the `QT001` entry — there is no header/metadata block to strip.

## No other mislabeled QT/BB/VN entries detected in this pass
Only QT002 required correction. The `_MISLABELED_` warning file is retained as a historical marker per instructions.
