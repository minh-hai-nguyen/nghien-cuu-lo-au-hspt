#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
move_phase2_pdfs.py

Di chuyen cac PDF HIGH-confidence tu 02_Papers-goc/Chua-phan-loai/ vao folder
canonical va cap nhat 02_Papers-goc/canonical_index.json theo plan o
06_Scripts/PHASE2_categorize_chua_phan_loai.md.

Mac dinh chay --dry-run. Truyen --apply de thuc su di chuyen.

Truoc khi --apply:
  - Backup canonical_index.json -> canonical_index.json.bak_phase2_<ts>
  - Kiem tra SHA256 voi duplicate target neu co
  - In log tung file
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Cau hinh
# ---------------------------------------------------------------------------

ROOT = Path(r"c:/Users/OS/OneDrive/read_books/Lo-au")
PAPERS = ROOT / "02_Papers-goc"
CHUA = PAPERS / "Chua-phan-loai"
INDEX = PAPERS / "canonical_index.json"

# Danh sach HIGH-confidence moves theo plan PHASE 2.
# Moi entry:
#   src           : duong dan tuong doi tu Chua-phan-loai/
#   dst_folder    : folder dich (Viet-Nam / Dong-Nam-A / The-gioi_Au-My-Uc / The-gioi_Khac)
#   dst_filename  : ten file moi sau khi rename (None = giu nguyen)
#   qt_id         : QT/VN id de gan/update trong canonical_index
#   action        : "new" (tao entry moi) | "update_folder" (chi set pdf_folder)
#                   | "verify_dupe" (so SHA voi target_dupe, neu match thi xoa src)
#   target_dupe   : duong dan target hien tai trong canonical (khi action=verify_dupe)
#   entry         : dict canonical (khi action=new)

PLAN: list[dict] = [
    # ---- QT076 / QT077 / QT078: entries da co, chi can set pdf_folder + move ----
    {
        "src": "Small_Blanc_2021_TamGiao_Vietnam.pdf",
        "dst_folder": "The-gioi_Au-My-Uc",
        "dst_filename": "QT076_Small_Blanc_2021_USA_TamGiao_VN.pdf",
        "qt_id": "QT076",
        "action": "update_folder",
    },
    {
        "src": "Stankov_2010_Confucian_Academic.pdf",
        "dst_folder": "The-gioi_Au-My-Uc",
        "dst_filename": "QT077_Stankov_2010_Australia_ConfucianAcademic.pdf",
        "qt_id": "QT077",
        "action": "update_folder",
    },
    {
        "src": "Rose_2002_CoRumination.pdf",
        "dst_folder": "The-gioi_Au-My-Uc",
        "dst_filename": "QT078_Rose_2002_USA_CoRumination.pdf",
        "qt_id": "QT078",
        "action": "update_folder",
    },
    # ---- QT079: Li 2025 (Australia, Chinese Aus, rumination + PTSD) ----
    {
        "src": "1-s2.0-S266691532500054X-main.pdf",
        "dst_folder": "The-gioi_Au-My-Uc",
        "dst_filename": "QT079_Li_2025_Australia_Rumination_PTSD_Chinese_EuroAus.pdf",
        "qt_id": "QT079",
        "action": "new",
        "entry": {
            "id": "QT079",
            "descriptor": "Li_2025_Australia_Rumination_PTSD",
            "summary": None,
            "translation": None,
            "pdf": "QT079_Li_2025_Australia_Rumination_PTSD_Chinese_EuroAus.pdf",
            "pdf_folder": "The-gioi_Au-My-Uc",
            "doi": None,
            "journal": "Journal of Mood and Anxiety Disorders (Elsevier)",
            "year": 2025,
            "authors": [
                "James Haoxiang Li", "Larissa Shiying Qiu", "Joshua Wong",
                "Winnie Lau", "Richard A. Bryant", "July Lies",
                "Belinda J. Liddell", "Laura Jobson",
            ],
            "country": "Australia",
            "sample": "111 European Australian + 111 Chinese Australian trauma survivors",
            "topic": "Cultural moderators of rumination-PTSD link",
            "source": "PHASE2_categorize_chua_phan_loai_07062026",
        },
    },
    # ---- QT040 dupe: Walder 2025 DMHI SAD (verified - cung paper voi QT040 existing) ----
    # SHA khac (typesetter version) nhung tac gia + tieu de + nam giong het QT040
    # -> verify_dupe va xoa source (KHONG tao QT080 moi)
    {
        "src": "Digital_MH_SocialAnxiety_MetaAnalysis_2025.pdf",
        "qt_id": "QT040",
        "action": "verify_dupe",
        "target_dupe": str(PAPERS / "The-gioi_Au-My-Uc" /
                            "QT040_Walder_JMIR_DMHI_SAD_2025.pdf"),
        "dst_folder": "The-gioi_Au-My-Uc",
        "dst_filename": None,
        "force_delete_on_content_match": True,
        "note": "SHA khac nhung tieu de+tac gia+nam khop QT040 - manual review recommended",
    },
    # ---- VN028: dupe cua Dao Thi Ngoan (verify SHA -> xoa source) ----
    {
        "src": "TCNCYH_2025_LoAu_TramCam.pdf",
        "qt_id": "VN028",
        "action": "verify_dupe",
        "target_dupe": str(PAPERS / "Viet-Nam" /
                            "VN028_DaoThiNgoan_TCNCYH_SVY4_HMU_2025.pdf"),
        "dst_folder": "Viet-Nam",
        "dst_filename": None,
    },
    # ---- VN031: Tran Thi My Luong 2020 ----
    {
        "src": "Viet-nam/CVv443S402020122.pdf",
        "dst_folder": "Viet-Nam",
        "dst_filename": "VN031_TranThiMyLuong_2020_THPT_LoAu.pdf",
        "qt_id": "VN031",
        "action": "new",
        "entry": {
            "id": "VN031",
            "descriptor": "TranThiMyLuong_2020_THPT_LoAu",
            "summary": None,
            "translation": None,
            "pdf": "VN031_TranThiMyLuong_2020_THPT_LoAu.pdf",
            "pdf_folder": "Viet-Nam",
            "doi": None,
            "journal": "Tap chi Khoa hoc - Truong Dai hoc Thu Do Ha Noi (CVv443)",
            "year": 2020,
            "authors": ["Tran Thi My Luong"],
            "country": "Vietnam",
            "sample": "Hoc sinh THPT - khao sat thuc trang",
            "topic": "Roi loan lo au o hoc sinh Trung hoc pho thong",
            "source": "PHASE2_categorize_chua_phan_loai_07062026",
        },
    },
    # ---- V-NAMHS bo sung: tao bien the VN002b ----
    # User can quyet: tao VN032 hay them suffix "b" cho VN002
    {
        "src": "tai-them-27052026/V-NAMHS_2022.pdf",
        "dst_folder": "Viet-Nam",
        "dst_filename": "VN002b_VNAMHS_2022_Main_Findings.pdf",
        "qt_id": "VN002b",
        "action": "new",
        "entry": {
            "id": "VN002b",
            "descriptor": "VNAMHS_2022_Main_Findings_Variant",
            "summary": None,
            "translation": None,
            "pdf": "VN002b_VNAMHS_2022_Main_Findings.pdf",
            "pdf_folder": "Viet-Nam",
            "doi": None,
            "journal": "Institute of Sociology - VNAMHS Report (Hanoi 2022)",
            "year": 2022,
            "authors": [
                "Vu Manh Loi", "Nguyen Duc Vinh", "Dao Thi Khanh Hoa",
                "Holly E. Erskine", "Cartiah McGrath", "Krystina Wallis",
                "Sarah J. Blondell", "Harvey A. Whiteford", "James G. Scott",
                "Robert Blum", "Shoshanna Fine", "Mengmeng Li", "Astha Ramaiya",
            ],
            "country": "Vietnam",
            "sample": "5996 adolescents (10-17y) - national survey",
            "topic": "V-NAMHS Main Findings Report (51-page short report variant of VN002)",
            "source": "PHASE2_categorize_chua_phan_loai_07062026",
            "note": "Variant of VN002 - verify if duplicate content before keeping",
        },
    },
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def sha256_of(path: Path, chunk: int = 1 << 16) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for blk in iter(lambda: f.read(chunk), b""):
            h.update(blk)
    return h.hexdigest()


def load_index() -> dict:
    return json.loads(INDEX.read_text(encoding="utf-8"))


def save_index(idx: dict) -> None:
    INDEX.write_text(
        json.dumps(idx, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def backup_index() -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = INDEX.with_suffix(f".json.bak_phase2_{ts}")
    shutil.copy2(INDEX, bak)
    return bak


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--apply", action="store_true",
                    help="Actually move files + update index (default = dry-run).")
    args = ap.parse_args()
    dry = not args.apply

    print("=" * 70)
    print(f"PHASE 2 move script - {'DRY RUN' if dry else 'APPLY MODE'}")
    print(f"Root      : {ROOT}")
    print(f"Index     : {INDEX}")
    print(f"Plan items: {len(PLAN)}")
    print("=" * 70)

    if not INDEX.exists():
        print(f"[FATAL] canonical_index.json not found: {INDEX}")
        return 2

    idx = load_index()
    if not dry:
        bak = backup_index()
        print(f"[backup] {bak.name}")

    stats = {"moved": 0, "skipped": 0, "dupe_deleted": 0,
             "entries_added": 0, "entries_updated": 0, "errors": 0}

    for i, item in enumerate(PLAN, 1):
        src = CHUA / item["src"]
        qt = item["qt_id"]
        action = item["action"]
        print(f"\n[{i:02d}] {qt} | {action:14s} | {item['src']}")

        if not src.exists():
            print(f"     [SKIP] source missing: {src}")
            stats["skipped"] += 1
            continue

        if src.stat().st_size < 5000:
            print(f"     [SKIP] source is stub-sized ({src.stat().st_size} B)")
            stats["skipped"] += 1
            continue

        # --- verify_dupe branch ---
        if action == "verify_dupe":
            tgt = Path(item["target_dupe"])
            if not tgt.exists():
                print(f"     [WARN] target dupe missing: {tgt}")
                stats["errors"] += 1
                continue
            try:
                h_src = sha256_of(src)
                h_tgt = sha256_of(tgt)
            except Exception as e:
                print(f"     [ERR ] sha256 failed: {e}")
                stats["errors"] += 1
                continue
            same = h_src == h_tgt
            sz_match = src.stat().st_size == tgt.stat().st_size
            print(f"     src sha   : {h_src[:16]}... ({src.stat().st_size} B)")
            print(f"     tgt sha   : {h_tgt[:16]}... ({tgt.stat().st_size} B)")
            print(f"     sha equal : {same} | size equal: {sz_match}")
            if same:
                if dry:
                    print("     [DRY] would DELETE source (true duplicate)")
                else:
                    src.unlink()
                    print("     [DEL] deleted duplicate source")
                stats["dupe_deleted"] += 1
            else:
                print("     [WARN] content differs - manual review needed; NOT deleting")
                stats["errors"] += 1
            continue

        # --- update_folder / new branches both move file ---
        dst_dir = PAPERS / item["dst_folder"]
        dst_name = item["dst_filename"] or src.name
        dst = dst_dir / dst_name

        if dst.exists():
            print(f"     [WARN] dst already exists: {dst}")
            print(f"            cannot overwrite; skipping move")
            stats["errors"] += 1
            continue

        # verify_not_dupe_of (QT080 vs QT040 etc.)
        if "verify_not_dupe_of" in item:
            other_qt = item["verify_not_dupe_of"]
            other = idx.get(other_qt, {})
            other_pdf = other.get("pdf")
            print(f"     [check] vs {other_qt}: pdf={other_pdf}")
            if other_pdf and other.get("pdf_folder"):
                other_path = PAPERS / other["pdf_folder"] / other_pdf
                if other_path.exists():
                    try:
                        if sha256_of(src) == sha256_of(other_path):
                            print(f"     [WARN] SHA match with {other_qt}; skipping new entry")
                            stats["errors"] += 1
                            continue
                    except Exception:
                        pass

        if dry:
            print(f"     [DRY] would MOVE -> {dst}")
        else:
            dst_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(str(src), str(dst))
            print(f"     [MV ] {dst}")
        stats["moved"] += 1

        # update index
        if action == "update_folder":
            if qt not in idx:
                print(f"     [WARN] {qt} not in index; cannot update_folder")
                stats["errors"] += 1
                continue
            old_folder = idx[qt].get("pdf_folder")
            old_pdf = idx[qt].get("pdf")
            if dry:
                print(f"     [DRY] would set {qt}.pdf_folder = '{item['dst_folder']}'")
                print(f"     [DRY] would set {qt}.pdf        = '{dst_name}'")
            else:
                idx[qt]["pdf_folder"] = item["dst_folder"]
                idx[qt]["pdf"] = dst_name
                print(f"     [UPD] {qt}: folder {old_folder} -> {item['dst_folder']}")
                print(f"     [UPD] {qt}: pdf    {old_pdf} -> {dst_name}")
            stats["entries_updated"] += 1
        elif action == "new":
            if qt in idx:
                print(f"     [WARN] {qt} already exists in index; skipping entry add")
                stats["errors"] += 1
            else:
                if dry:
                    print(f"     [DRY] would ADD entry {qt}")
                else:
                    idx[qt] = item["entry"]
                    print(f"     [ADD] {qt}")
                stats["entries_added"] += 1

    if not dry:
        save_index(idx)
        print(f"\n[save] {INDEX}")

    print("\n" + "=" * 70)
    print("STATS")
    for k, v in stats.items():
        print(f"  {k:18s}: {v}")
    print("=" * 70)
    if dry:
        print("DRY RUN complete. Re-run with --apply to perform changes.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
