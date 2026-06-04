"""
So sánh PROJECT_INVENTORY_13052026.md (máy kia) với máy hiện tại.
Xuất MISSING_FILES_FOR_OTHER_MACHINE.txt — danh sách file/folder cần copy từ máy kia về.

Run: python 06_Scripts/compare_inventory_14052026.py
"""
import re
from pathlib import Path
from datetime import datetime

ROOT = Path(r"c:/Users/HLC/OneDrive/read_books/Lo-au")
INVENTORY = ROOT / "PROJECT_INVENTORY_13052026.md"
OUTPUT = ROOT / "MISSING_FILES_FOR_OTHER_MACHINE.txt"

EXCLUDE_DIRS = {".git", "node_modules", "__pycache__", ".venv", "venv",
                ".claude", "_workspace", ".idea", ".vscode"}
EXCLUDE_EXT = {".pyc"}


def extract_files_from_inventory(md_path: Path) -> dict[str, dict]:
    """
    Parse markdown inventory. Each file line looks like:
      - `02_Papers-goc/.../file.pdf` — 979.2KB, mod 2026-03-24
    Returns dict: relpath -> {"size": str, "mod": str}
    """
    files = {}
    pattern = re.compile(r"^\s*-\s+`([^`]+)`\s+—\s+(.+?)$")
    with md_path.open("r", encoding="utf-8") as f:
        for line in f:
            m = pattern.match(line.rstrip("\n"))
            if not m:
                continue
            relpath = m.group(1).strip()
            meta = m.group(2).strip()
            # Skip if path looks like a directory entry (no extension and no slash → unlikely)
            files[relpath] = {"meta": meta}
    return files


def scan_local(root: Path) -> set[str]:
    """Scan local filesystem, return set of relative paths (forward slash)."""
    local = set()
    for p in root.rglob("*"):
        if not p.is_file():
            continue
        # Skip excluded dirs
        if any(part in EXCLUDE_DIRS for part in p.parts):
            continue
        if p.suffix in EXCLUDE_EXT:
            continue
        rel = p.relative_to(root).as_posix()
        local.add(rel)
    return local


def classify_priority(path: str) -> str:
    """Classify priority based on path."""
    if path.startswith("02_Papers-goc/"):
        return "1_CRITICAL"
    if path.startswith("03_Ban-dich/"):
        return "2_HIGH"
    if path.startswith("Tom-tat-tung-bai/"):
        return "2_HIGH"
    if path.startswith("01_Bao-cao/"):
        return "3_MEDIUM_HIGH"
    if path.startswith("06_Scripts/"):
        return "4_MEDIUM"
    if path.startswith("04_Co-so-du-lieu/"):
        return "4_MEDIUM"
    if path.startswith("05_Cong-cu/"):
        return "4_MEDIUM"
    if path.startswith("Cropped_Figures/"):
        return "5_MEDIUM_LOW"
    if path.startswith("DocFiles/"):
        return "5_MEDIUM_LOW"
    if path.startswith("madam Thao/"):
        return "5_MEDIUM_LOW"
    if path.startswith("Bai_goc_BaoCao_CanThiep_10042026/"):
        return "5_MEDIUM_LOW"
    if path.startswith("paper-may/"):
        return "5_MEDIUM_LOW"
    if path.startswith("chưa-sắp-xếp/"):
        return "5_MEDIUM_LOW"
    if path.startswith("tro-ly-nghien-cuu-tam-ly/") or path.startswith("tro-ly-nghien-cuu-tam-ly-light/"):
        return "4_MEDIUM"
    if path.startswith("rag_") or path.startswith("lib/") or path.startswith("Materials/"):
        return "6_LOW"
    if path.startswith("_Archive") or path.startswith("00_Meta/") or path.startswith("Tập-viết-theo-phong-cách/"):
        return "7_OPTIONAL"
    return "8_OTHER"


PRIORITY_LABELS = {
    "1_CRITICAL": "## ƯU TIÊN 1 — CRITICAL — PDF gốc tài liệu (02_Papers-goc/)",
    "2_HIGH": "## ƯU TIÊN 2 — HIGH — Bản dịch + Tóm tắt (03_Ban-dich/, Tom-tat-tung-bai/)",
    "3_MEDIUM_HIGH": "## ƯU TIÊN 3 — MEDIUM HIGH — Báo cáo + trả lời thầy (01_Bao-cao/)",
    "4_MEDIUM": "## ƯU TIÊN 4 — MEDIUM — Scripts + cơ sở dữ liệu + công cụ + chatbot",
    "5_MEDIUM_LOW": "## ƯU TIÊN 5 — MEDIUM LOW — Hình ảnh + tài liệu phụ trợ",
    "6_LOW": "## ƯU TIÊN 6 — LOW — RAG ChromaDB + thư viện code",
    "7_OPTIONAL": "## ƯU TIÊN 7 — OPTIONAL — Archive + meta + tập viết",
    "8_OTHER": "## ƯU TIÊN 8 — KHÁC",
}


def main():
    print(f"[INFO] Reading inventory: {INVENTORY}")
    inventory = extract_files_from_inventory(INVENTORY)
    print(f"[INFO] Inventory has {len(inventory)} files")

    print(f"[INFO] Scanning local filesystem: {ROOT}")
    local = scan_local(ROOT)
    print(f"[INFO] Local has {len(local)} files")

    # Normalize: inventory uses forward slashes already
    inv_set = set(inventory.keys())
    missing = sorted(inv_set - local)  # On inventory but not local — need to copy from other machine
    extra = sorted(local - inv_set)    # On local but not on inventory — newer or local-only

    print(f"[INFO] Missing on local (need from other machine): {len(missing)}")
    print(f"[INFO] Extra on local (other machine doesn't have): {len(extra)}")

    # Group missing by priority
    by_priority = {}
    for path in missing:
        prio = classify_priority(path)
        by_priority.setdefault(prio, []).append(path)

    # Group missing by top-level folder for quick scan
    by_folder = {}
    for path in missing:
        top = path.split("/", 1)[0] if "/" in path else "(root)"
        by_folder.setdefault(top, []).append(path)

    # Build output
    lines = []
    lines.append("=" * 80)
    lines.append("DANH SÁCH FILE/THƯ MỤC CẦN COPY TỪ MÁY KIA → MÁY NÀY")
    lines.append("=" * 80)
    lines.append(f"Sinh ngày: {datetime.now():%Y-%m-%d %H:%M}")
    lines.append(f"So sánh với: PROJECT_INVENTORY_13052026.md (máy kia, 13/05/2026)")
    lines.append(f"Máy hiện tại: {ROOT}")
    lines.append("")
    lines.append("HƯỚNG DẪN SỬ DỤNG:")
    lines.append("  Gửi file này cho máy kia. Máy kia gom các file/thư mục theo danh sách,")
    lines.append("  nén lại (zip/7z) và copy về máy này theo đường dẫn tương đối tương ứng.")
    lines.append("")
    lines.append("-" * 80)
    lines.append("TỔNG KẾT")
    lines.append("-" * 80)
    lines.append(f"  Tổng file thiếu trên máy này:  {len(missing):,}")
    lines.append(f"  Tổng file thừa trên máy này:   {len(extra):,} (máy kia không có)")
    lines.append("")
    lines.append("PHÂN BỔ THEO THƯ MỤC GỐC:")
    for folder in sorted(by_folder.keys()):
        lines.append(f"  {folder:50s} {len(by_folder[folder]):4d} file")
    lines.append("")
    lines.append("=" * 80)
    lines.append("DANH SÁCH FILE THIẾU CHI TIẾT")
    lines.append("=" * 80)

    for prio in sorted(by_priority.keys()):
        lines.append("")
        lines.append(PRIORITY_LABELS.get(prio, prio))
        lines.append(f"  ({len(by_priority[prio])} file)")
        lines.append("")
        # Group by folder within priority
        sub_by_folder = {}
        for path in by_priority[prio]:
            parts = path.split("/")
            folder = "/".join(parts[:-1]) if len(parts) > 1 else "(root)"
            sub_by_folder.setdefault(folder, []).append(path)
        for folder in sorted(sub_by_folder.keys()):
            lines.append(f"  [{folder}/]")
            for path in sorted(sub_by_folder[folder]):
                meta = inventory[path].get("meta", "")
                lines.append(f"    {path}    ({meta})")
            lines.append("")

    lines.append("=" * 80)
    lines.append("PHỤ LỤC — FILE THỪA TRÊN MÁY NÀY (máy kia không có — không cần gom)")
    lines.append("=" * 80)
    lines.append(f"  (Tổng: {len(extra)} file)")
    lines.append("")
    extra_by_folder = {}
    for path in extra:
        top = path.split("/", 1)[0] if "/" in path else "(root)"
        extra_by_folder.setdefault(top, []).append(path)
    for folder in sorted(extra_by_folder.keys()):
        lines.append(f"  [{folder}/]  ({len(extra_by_folder[folder])} file)")
        for path in sorted(extra_by_folder[folder])[:50]:  # Limit 50 per folder
            lines.append(f"    {path}")
        if len(extra_by_folder[folder]) > 50:
            lines.append(f"    ... và {len(extra_by_folder[folder]) - 50} file khác")
        lines.append("")

    lines.append("=" * 80)
    lines.append("HẾT")
    lines.append("=" * 80)

    OUTPUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"[DONE] Output written to: {OUTPUT}")
    print(f"[DONE] Size: {OUTPUT.stat().st_size / 1024:.1f} KB")


if __name__ == "__main__":
    main()
