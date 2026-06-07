#!/usr/bin/env python3
"""
move_chua_phan_loai_pdfs.py
============================

Moves PDFs from ``02_Papers-goc/Chua-phan-loai/`` (and its subfolders
``Viet-nam/``, ``tai-them-27052026/``) into the four canonical category
folders, based on the plan documented in
``06_Scripts/CATEGORIZE_PLAN_Chua-phan-loai_07062026.md``.

Behaviour
---------
- HIGH-confidence items are moved automatically (with optional rename to
  canonical filename if a duplicate match is identified).
- MEDIUM / LOW confidence items are written to a manual-review log and
  NOT moved.
- Stub files (size < 30 KB) are flagged and NEVER moved automatically,
  even if HIGH confidence.
- ``--dry-run`` prints the actions without touching the filesystem.

Safety
------
- If the destination path already exists, the source file is left in
  place and the conflict is logged.
- The ``bai-bao-khgdvn/`` folder is never touched (per project memory).
- ``canonical_index.json`` is NOT modified by this script.

Usage
-----
    python move_chua_phan_loai_pdfs.py --dry-run
    python move_chua_phan_loai_pdfs.py            # actual move
"""
from __future__ import annotations

import argparse
import datetime as _dt
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(r"c:\Users\OS\OneDrive\read_books\Lo-au")
PAPERS_ROOT = PROJECT_ROOT / "02_Papers-goc"
SOURCE_ROOT = PAPERS_ROOT / "Chua-phan-loai"

DEST_VN = PAPERS_ROOT / "Viet-Nam"
DEST_DNA = PAPERS_ROOT / "Dong-Nam-A"
DEST_WEST = PAPERS_ROOT / "The-gioi_Au-My-Uc"
DEST_OTHER = PAPERS_ROOT / "The-gioi_Khac"

MANUAL_REVIEW_LOG = (
    PROJECT_ROOT / "06_Scripts" / "manual_review_chua_phan_loai_07062026.log"
)
STUB_BYTES_THRESHOLD = 30_000   # files < 30 KB treated as stubs / cloud-only

# ---------------------------------------------------------------------------
# Categorization plan
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class PlanEntry:
    source_rel: str                # path relative to SOURCE_ROOT, forward slashes
    destination: Path              # target folder
    confidence: str                # HIGH / MEDIUM / LOW
    rename_to: Optional[str] = None  # if duplicate-of-canonical, rename on move
    note: str = ""

# NOTE: items with rename_to are duplicates of an already-indexed paper.
# The script will SKIP them if the renamed target already exists (most
# likely case), or move-and-rename otherwise.
PLAN: list[PlanEntry] = [
    # ---------- Viet-Nam ----------
    PlanEntry("HoangTrungHoc_2025_COVID_VN.pdf", DEST_VN, "HIGH",
              rename_to="VN014_HoangTrungHoc_2025_VN_COVID.pdf",
              note="Duplicate of VN014 (stub-sized: 22 B -- will be flagged)"),
    PlanEntry("NgoAnhVinh_2024_DTTS_LangSon.pdf", DEST_VN, "HIGH",
              rename_to="VN015_NgoAnhVinh_2024_LangSon_DTTS.pdf"),
    PlanEntry("Pham_2024_QualityCare_VN_Adolescents.pdf", DEST_VN, "MEDIUM",
              note="Stub-sized (7 KB) -- manual review; QT076 candidate"),
    PlanEntry("VN_Hue_LoAu_TramCam_VTN_2025.pdf", DEST_VN, "MEDIUM",
              rename_to=None,
              note="Likely dupe of VN019 (HoThiTrucQuynh Hue 2025) -- verify"),
    PlanEntry("VN_TPHCM_DASS_THPT_2024.pdf", DEST_VN, "MEDIUM",
              rename_to=None,
              note="Likely dupe of VN020 (TranHoVinhLoc TPHCM DASS-Y) -- verify"),
    PlanEntry("VN_PNT_DASS_THPT_2025.pdf", DEST_VN, "MEDIUM",
              rename_to=None,
              note="Likely dupe of VN026 (TranDucSi LongAn PNT 2025) -- verify"),
    PlanEntry("VN_Multicenter_2631HS_TPHCM_2025.pdf", DEST_VN, "HIGH",
              rename_to="VN029_Duong_SocPsychiatry_2631HS_TPHCM_2025.pdf"),
    PlanEntry("SchoolFactors_VN_Anxiety_2021.pdf", DEST_VN, "MEDIUM",
              note="Stub-sized (16 B); likely dupe of VN027 -- manual review"),
    PlanEntry("Study on school-related factors impacting mental health and well-being of adolescents in Viet Nam.pdf",
              DEST_VN, "HIGH",
              rename_to="VN022_UNICEF_VN_2022_SchoolFactors.pdf"),
    PlanEntry("TCNCYH_2025_LoAu_TramCam.pdf", DEST_VN, "MEDIUM",
              note="Likely dupe of VN028 (DaoThiNgoan TCNCYH SVY4 HMU 2025) -- verify"),
    PlanEntry("VNAMHS-Report_Eng_15-Feb-2023.pdf", DEST_VN, "HIGH",
              rename_to="VN002_VNAMHS_2022_National.pdf"),
    PlanEntry("tai-them-27052026/V-NAMHS_2022.pdf", DEST_VN, "MEDIUM",
              note="V-NAMHS variant -- likely same dataset as VN002; verify Vietnamese vs English"),
    PlanEntry("Viet-nam/16.-NguyenThiVan.pdf", DEST_VN, "HIGH",
              rename_to="VN004_NguyenThiVan_2020_STAI_TPHCM.pdf"),
    PlanEntry("Viet-nam/TRANNGUYENNGOC-TamThan.pdf", DEST_VN, "HIGH",
              rename_to="VN005_TranNguyenNgoc_2018_LuanAn_ThuGian_GAD.pdf"),
    PlanEntry("Viet-nam/CVv443S402020122.pdf", DEST_VN, "LOW",
              note="Opaque VN journal code (Vietnam Journals Online CVv443) -- inspect; QT077 candidate"),
    PlanEntry("Small_Blanc_2021_TamGiao_Vietnam.pdf", DEST_VN, "MEDIUM",
              note="Tam Giao = three Vietnamese teachings; QT085 candidate"),

    # ---------- Dong-Nam-A ----------
    PlanEntry("Indonesia_Adolescent_MH_2024.pdf", DEST_DNA, "HIGH",
              note="QT078 candidate"),
    PlanEntry("1-s2.0-S266691532500054X-main.pdf", DEST_DNA, "LOW",
              note="Opaque ScienceDirect ID; possibly DNA -- inspect"),
    PlanEntry("PIIS2468266725000982.pdf", DEST_DNA, "LOW",
              note="Lancet Public Health PII; geography unknown -- inspect"),

    # ---------- The-gioi_Au-My-Uc ----------
    PlanEntry(
        "Child Adoles Psych Nursing - 2025 - Anderson - Contributing Factors to "
        "the Rise in Adolescent Anxiety and Associated Mental.pdf",
        DEST_WEST, "HIGH",
        rename_to="QT014_Anderson_2025_Wiley_Narrative.pdf",
    ),
    PlanEntry("Social_Anxiety_Young_People_7Countries_2020_PLOSONE.pdf", DEST_WEST, "HIGH",
              rename_to="QT035_Jefferies_SocialAnxiety_7Countries_2020.pdf"),
    PlanEntry("Digital_MH_SocialAnxiety_MetaAnalysis_2025.pdf", DEST_WEST, "MEDIUM",
              note="Possibly QT040 (Walder DMHI SAD) -- verify"),
    PlanEntry("Mindfulness_Nature_MentalHealth_2023.pdf", DEST_WEST, "HIGH",
              rename_to="QT052_Mindfulness_NatureMH_IPD_MA_2023.pdf"),
    PlanEntry("Pharmacotherapy_Anxiety_2020_Frontiers.pdf", DEST_WEST, "HIGH",
              rename_to="QT053_Pharmacotherapy_Anxiety_Review_Frontiers_2020.pdf"),
    PlanEntry("Chronic_Stress_Neuroinflammation_FrontPsych_2023.pdf", DEST_WEST, "HIGH",
              rename_to="QT054_ChronicStress_Neuroinflammation_FrontPsych_2023.pdf"),
    PlanEntry("Epigenetics_Childhood_Maltreatment_TranslPsych_2021.pdf", DEST_WEST, "HIGH",
              rename_to="QT055_Parade_Epigenetics_Maltreatment_TranslPsych_2021.pdf"),
    PlanEntry("Neural_Circuit_Pathological_Anxiety_NatRevNeuro_2024.pdf", DEST_WEST, "HIGH",
              note="QT079 candidate (Nat Rev Neuro 2024)"),
    PlanEntry("Neural_Circuits_Mechanisms_Anxiety_2025_FrontNeuralCirc.pdf", DEST_WEST, "HIGH",
              rename_to="QT057_NeuralCircuits_Mechanisms_Anxiety_FrontNeuralCirc_2025.pdf"),
    PlanEntry("GWAS_Anxiety_NatHumBehav_2023.pdf", DEST_WEST, "HIGH",
              note="QT080 candidate"),
    PlanEntry("GWAS_MultiAncestry_NatGenetics_2024.pdf", DEST_WEST, "HIGH",
              note="QT081 candidate"),
    PlanEntry("GWAS_122K_GABAergic_NatGenetics_2025.pdf", DEST_WEST, "HIGH",
              note="QT082 candidate"),
    PlanEntry("CBT_Delivery_GAD_TranslPsych_2025.pdf", DEST_WEST, "MEDIUM",
              note="Possibly QT020 (Liu CBT NMA TranslPsych 2025) -- verify journal vs NMA"),
    PlanEntry("tai-them-27052026/Pascoe_2020.pdf", DEST_WEST, "HIGH",
              rename_to="QT067_Pascoe_AcademicStress_IJAY_2020.pdf"),
    PlanEntry("tai-them-27052026/Steare_2023_AcademicPressure_SR.pdf", DEST_WEST, "HIGH",
              note="QT083 candidate (UCL UK SR)"),
    PlanEntry("tai-them-27052026/Compas_2017_Coping_MetaAnalysis.pdf", DEST_WEST, "HIGH",
              note="QT084 candidate (Compas Vanderbilt USA)"),

    # ---------- The-gioi_Khac ----------
    PlanEntry("Nakie_2022_Ethiopia_HighSchool_BMCPsych.pdf", DEST_OTHER, "HIGH",
              rename_to="QT006_Nakie_et_al_2022_Ethiopia.pdf"),
    PlanEntry("Chen_2023_Chinese_SecondarySchool_BMCPsych.pdf", DEST_OTHER, "HIGH",
              rename_to="QT007_Chen_et_al_2023_China_BMCPsych.pdf"),
    PlanEntry("Wen_2020_LPA_Anxiety_Rural_China.pdf", DEST_OTHER, "HIGH",
              rename_to="QT008_Wen_2020_China_Rural_LPA.pdf"),
    PlanEntry("Qiu_2022_ParentingStyle_Resilience_FrontPsych.pdf", DEST_OTHER, "HIGH",
              rename_to="QT009_Qiu_et_al_2022_China_Parenting.pdf"),
    PlanEntry("Bhardwaj_2020_Chandigarh_DASS21.pdf", DEST_OTHER, "HIGH",
              rename_to="QT011_Bhardwaj_et_al_2020_India_Chandigarh.pdf"),
    PlanEntry("Korea_Adolescent_MH_Trends_2024.pdf", DEST_OTHER, "HIGH",
              rename_to="QT034_Korea_Cho_MH_Trends_NatSciRep_2024.pdf"),
    PlanEntry("Microbiota_GutBrain_Anxiety_2024_FrontNeuro.pdf", DEST_OTHER, "HIGH",
              rename_to="QT056_Jiang_Microbiota_GutBrain_Anxiety_FrontNeuro_2024.pdf"),
    PlanEntry("Neuroinflammation_Neuromodulation_TranslPsych_2022.pdf", DEST_OTHER, "HIGH",
              rename_to="QT058_Guo_Neuroinflammation_Neuromodulation_TranslPsych_2022.pdf"),
    PlanEntry("Stankov_2010_Confucian_Academic.pdf", DEST_OTHER, "LOW",
              note="Stub-sized (1.8 KB) -- manual review"),

    # ---------- Defer (LOW) ----------
    PlanEntry("fpubh-12-1232856.pdf", DEST_OTHER, "LOW",
              note="Frontiers Public Health opaque ID -- inspect"),
    PlanEntry("journal.pone.0316825.pdf", DEST_OTHER, "LOW",
              note="PLOS ONE opaque ID -- inspect"),
]


# ---------------------------------------------------------------------------
# Worker
# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--dry-run", action="store_true",
                    help="Preview without moving any file.")
    args = ap.parse_args()

    if not SOURCE_ROOT.is_dir():
        print(f"[FATAL] Source folder missing: {SOURCE_ROOT}", file=sys.stderr)
        return 2

    manual_lines: list[str] = []
    moved = 0
    skipped_dupe = 0
    skipped_missing = 0
    skipped_stub = 0
    skipped_manual = 0
    skipped_conflict = 0

    now = _dt.datetime.now().isoformat(timespec="seconds")
    manual_lines.append(f"# Manual-review log -- generated {now}\n")
    manual_lines.append(f"# Source: {SOURCE_ROOT}\n\n")

    for entry in PLAN:
        src = SOURCE_ROOT / entry.source_rel
        target_name = entry.rename_to or src.name
        dst = entry.destination / target_name

        prefix = f"[{entry.confidence}]"
        rel_disp = entry.source_rel

        # Guarded existence + stat (OneDrive reparse points can throw)
        try:
            exists = src.exists()
        except OSError as exc:
            print(f"{prefix} REPARSE-ERR -- {rel_disp}: {exc}")
            manual_lines.append(
                f"REPARSE\t{rel_disp}\t(OneDrive reparse/cloud stub error: {exc}; force-download in OneDrive then retry)\n"
            )
            skipped_missing += 1
            continue

        if not exists:
            print(f"{prefix} MISSING -- {rel_disp}")
            manual_lines.append(f"MISSING\t{rel_disp}\t(file not found on disk)\n")
            skipped_missing += 1
            continue

        # Stub check (small file: probably a placeholder)
        try:
            size = src.stat().st_size
        except OSError as exc:
            size = -1
            print(f"{prefix} STAT-FAIL -- {rel_disp}: {exc}")
            manual_lines.append(f"STAT-FAIL\t{rel_disp}\t({exc})\n")
            skipped_missing += 1
            continue
        is_stub = 0 <= size < STUB_BYTES_THRESHOLD

        if is_stub:
            print(f"{prefix} STUB ({size} B) -- {rel_disp} -- not moved")
            manual_lines.append(
                f"STUB\t{rel_disp}\tsize={size}B; redownload before move; note={entry.note}\n"
            )
            skipped_stub += 1
            continue

        if entry.confidence != "HIGH":
            print(f"{prefix} MANUAL -- {rel_disp} -> {dst} ({entry.note})")
            manual_lines.append(
                f"{entry.confidence}\t{rel_disp}\t-> {dst}\t{entry.note}\n"
            )
            skipped_manual += 1
            continue

        # HIGH confidence
        if dst.exists():
            # Duplicate of canonical target already in place; skip the move.
            print(f"{prefix} DUPE -- {rel_disp} -- target {dst.name} already in {dst.parent.name}")
            manual_lines.append(
                f"DUPE\t{rel_disp}\ttarget exists: {dst}\t(safe to delete source after manual SHA/size verify)\n"
            )
            skipped_dupe += 1
            continue

        # Conflict guard: file at destination with a different size = unusual
        # but treat as conflict and skip.
        if any(p.name == dst.name for p in entry.destination.glob("*")):
            print(f"{prefix} CONFLICT -- {rel_disp} -- destination name collision")
            manual_lines.append(
                f"CONFLICT\t{rel_disp}\tdest name collision at {dst}\n"
            )
            skipped_conflict += 1
            continue

        action = "DRY-MOVE" if args.dry_run else "MOVE"
        print(f"{prefix} {action} -- {rel_disp} -> {dst}")
        if not args.dry_run:
            entry.destination.mkdir(parents=True, exist_ok=True)
            shutil.move(str(src), str(dst))
        moved += 1

    # Always write manual review log (even on dry-run)
    MANUAL_REVIEW_LOG.parent.mkdir(parents=True, exist_ok=True)
    with MANUAL_REVIEW_LOG.open("w", encoding="utf-8") as f:
        f.writelines(manual_lines)

    print()
    print("=" * 60)
    print(f"  HIGH-confidence moves : {moved}{' (dry-run)' if args.dry_run else ''}")
    print(f"  Duplicates skipped    : {skipped_dupe}")
    print(f"  Stub files skipped    : {skipped_stub}")
    print(f"  Manual-review skipped : {skipped_manual}")
    print(f"  Missing on disk       : {skipped_missing}")
    print(f"  Name conflicts        : {skipped_conflict}")
    print(f"  Manual log written to : {MANUAL_REVIEW_LOG}")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
