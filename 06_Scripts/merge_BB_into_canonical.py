"""Merge BB## entries from canonical_entries_BB_07062026.json into 02_Papers-goc/canonical_index.json.

Workflow:
  1. Backup canonical_index.json -> canonical_index.json.bak_<timestamp>
  2. Load BB entries.
  3. For each BB id: if not in canonical_index -> add; if present -> skip (report).
  4. Write back with indent=2, ensure_ascii=False.
  5. Print summary (added / skipped / total).

Run from project root:
    python 06_Scripts/merge_BB_into_canonical.py

Use --dry-run to preview without writing.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
CANONICAL_PATH = PROJECT_ROOT / "02_Papers-goc" / "canonical_index.json"
BB_ENTRIES_PATH = PROJECT_ROOT / "06_Scripts" / "canonical_entries_BB_07062026.json"


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def write_json(path: Path, data: dict) -> None:
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write("\n")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dry-run", action="store_true", help="Preview without writing")
    args = parser.parse_args()

    if not CANONICAL_PATH.exists():
        print(f"ERROR: canonical index not found: {CANONICAL_PATH}", file=sys.stderr)
        return 1
    if not BB_ENTRIES_PATH.exists():
        print(f"ERROR: BB entries not found: {BB_ENTRIES_PATH}", file=sys.stderr)
        return 1

    canonical = load_json(CANONICAL_PATH)
    bb_entries = load_json(BB_ENTRIES_PATH)

    added: list[str] = []
    skipped: list[str] = []

    for key, entry in bb_entries.items():
        if key in canonical:
            skipped.append(key)
            continue
        canonical[key] = entry
        added.append(key)

    print(f"BB entries total : {len(bb_entries)}")
    print(f"  to add         : {len(added)} -> {added}")
    print(f"  already present: {len(skipped)} -> {skipped}")
    print(f"canonical size before: {len(canonical) - len(added)}")
    print(f"canonical size after : {len(canonical)}")

    if args.dry_run:
        print("[dry-run] no changes written.")
        return 0

    if not added:
        print("Nothing new to add. Skipping backup + write.")
        return 0

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = CANONICAL_PATH.with_suffix(f".json.bak_{ts}")
    shutil.copy2(CANONICAL_PATH, backup_path)
    print(f"Backup written: {backup_path}")

    write_json(CANONICAL_PATH, canonical)
    print(f"Merged: {CANONICAL_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
