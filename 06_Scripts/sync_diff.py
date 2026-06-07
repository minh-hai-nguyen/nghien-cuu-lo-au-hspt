# -*- coding: utf-8 -*-
"""
Đồng bộ Lo-au (local) ↔ GitHub repo clone (sidecar).
- Bước 1 (DRY-RUN): inventory + diff + in báo cáo, KHÔNG sửa gì.
- Bước 2 (APPLY): thực thi sync theo phân loại (chạy khi truyền --apply).

Phân loại đường dẫn tương đối:
- SAME         : sha1 trùng → không làm gì.
- LOCAL-ONLY   : chỉ có ở local → copy local → sidecar.
- REMOTE-ONLY  : chỉ có ở sidecar → copy sidecar → local.
- CONFLICT-B   : sha khác, đuôi binary → giữ cả hai (rename remote `_remote`).
- CONFLICT-T   : sha khác, đuôi text → cố merge; fallback giữ cả hai.

Bỏ qua khi quét: .git/, desktop.ini, Thumbs.db, ~$*, .DS_Store.
"""
import hashlib
import json
import os
import shutil
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

# stdout utf-8 cho tiếng Việt
sys.stdout.reconfigure(encoding="utf-8") if hasattr(sys.stdout, "reconfigure") else None

LOCAL = Path(r"c:/Users/HLC/OneDrive/read_books/Lo-au")
SIDECAR = Path(r"C:/Users/HLC/Documents/sync-lo-au/repo")

# Đuôi text — sẽ thử merge
TEXT_EXTS = {".md", ".txt", ".py", ".json", ".csv", ".html", ".css", ".js",
             ".xml", ".yaml", ".yml", ".ini", ".cfg", ".sh", ".graphml", ".log"}

# Pattern bỏ qua khi quét
SKIP_NAMES = {"desktop.ini", "Thumbs.db", ".DS_Store"}
SKIP_PREFIXES = ("~$",)


def is_skipped(rel_path: Path) -> bool:
    parts = rel_path.parts
    # bỏ .git/ nested
    if ".git" in parts:
        return True
    name = rel_path.name
    if name in SKIP_NAMES:
        return True
    if any(name.startswith(p) for p in SKIP_PREFIXES):
        return True
    return False


def sha1_of(path: Path) -> str:
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def inventory(root: Path) -> dict:
    """Trả về dict: relpath str -> (sha1, size)."""
    inv = {}
    n = 0
    for p in root.rglob("*"):
        if not p.is_file():
            continue
        rel = p.relative_to(root)
        if is_skipped(rel):
            continue
        try:
            inv[str(rel).replace("\\", "/")] = (sha1_of(p), p.stat().st_size)
        except Exception as e:
            print(f"  !! lỗi đọc {rel}: {e}", file=sys.stderr)
        n += 1
        if n % 500 == 0:
            print(f"  ...quét {n} file", file=sys.stderr, flush=True)
    return inv


def categorize(local_inv: dict, remote_inv: dict) -> dict:
    """Phân loại từng đường dẫn."""
    out = {"SAME": [], "LOCAL-ONLY": [], "REMOTE-ONLY": [],
           "CONFLICT-B": [], "CONFLICT-T": []}
    all_paths = set(local_inv) | set(remote_inv)
    for rel in sorted(all_paths):
        lh = local_inv.get(rel)
        rh = remote_inv.get(rel)
        if lh and rh:
            if lh[0] == rh[0]:
                out["SAME"].append(rel)
            else:
                ext = Path(rel).suffix.lower()
                if ext in TEXT_EXTS:
                    out["CONFLICT-T"].append(rel)
                else:
                    out["CONFLICT-B"].append(rel)
        elif lh:
            out["LOCAL-ONLY"].append(rel)
        else:
            out["REMOTE-ONLY"].append(rel)
    return out


def fmt_bytes(n: int) -> str:
    for unit in ["B", "KB", "MB", "GB"]:
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def report(cats: dict, local_inv: dict, remote_inv: dict, max_show: int = 50):
    print("\n" + "=" * 70)
    print("BÁO CÁO DRY-RUN")
    print("=" * 70)
    total = sum(len(v) for v in cats.values())
    for k, v in cats.items():
        print(f"  {k:14s} {len(v):5d} file")
    print(f"  {'TỔNG':14s} {total:5d} file (unique paths)")

    # bytes copy mỗi chiều
    bytes_to_remote = sum(local_inv[r][1] for r in cats["LOCAL-ONLY"])
    bytes_to_local = sum(remote_inv[r][1] for r in cats["REMOTE-ONLY"])
    bytes_conf_b = sum(remote_inv[r][1] for r in cats["CONFLICT-B"])
    bytes_conf_t = sum(remote_inv[r][1] for r in cats["CONFLICT-T"])
    print(f"\n  Tổng byte LOCAL → SIDECAR (only-local): {fmt_bytes(bytes_to_remote)}")
    print(f"  Tổng byte SIDECAR → LOCAL (only-remote): {fmt_bytes(bytes_to_local)}")
    print(f"  Tổng byte conflict BINARY (giữ cả 2 phía): {fmt_bytes(bytes_conf_b)}")
    print(f"  Tổng byte conflict TEXT (thử merge): {fmt_bytes(bytes_conf_t)}")

    for cat in ["LOCAL-ONLY", "REMOTE-ONLY", "CONFLICT-B", "CONFLICT-T"]:
        items = cats[cat]
        if not items:
            continue
        print(f"\n--- {cat} (hiển thị {min(max_show, len(items))}/{len(items)}) ---")
        for p in items[:max_show]:
            sz = (local_inv if cat in ("LOCAL-ONLY",) else remote_inv).get(
                p, ("", 0))[1]
            print(f"  {fmt_bytes(sz):>8s}  {p}")


def try_merge_json(local_text: str, remote_text: str, rel: str):
    """Thử merge JSON. Hỗ trợ:
    - dict có key 'entries' là list of dict có 'meta.id' hoặc 'id'.
    - dict thường: union key, conflict key → giữ local.
    Trả về (merged_text, ok)."""
    try:
        L = json.loads(local_text)
        R = json.loads(remote_text)
    except Exception:
        return None, False
    if isinstance(L, dict) and isinstance(R, dict) and "entries" in L and "entries" in R:
        # RAG-style index
        def key_of(e):
            m = e.get("meta") or {}
            return m.get("id") or e.get("id") or json.dumps(e, sort_keys=True)[:50]
        Lm = {key_of(e): e for e in L["entries"]}
        Rm = {key_of(e): e for e in R["entries"]}
        all_keys = list(Lm) + [k for k in Rm if k not in Lm]
        merged_entries = []
        for k in all_keys:
            le = Lm.get(k); re = Rm.get(k)
            if le and re:
                # chọn entry dài hơn (nhiều thông tin hơn)
                if len(json.dumps(le, ensure_ascii=False)) >= len(
                        json.dumps(re, ensure_ascii=False)):
                    merged_entries.append(le)
                else:
                    merged_entries.append(re)
            else:
                merged_entries.append(le or re)
        merged = dict(L)
        merged["entries"] = merged_entries
        # cập nhật meta count nếu có
        if isinstance(merged.get("meta"), dict) and "n_entries" in merged["meta"]:
            merged["meta"]["n_entries"] = len(merged_entries)
        return json.dumps(merged, ensure_ascii=False, indent=1), True
    if isinstance(L, list) and isinstance(R, list):
        # list (vd kg_triples): dedupe theo json string
        seen = set(); out = []
        for x in L + R:
            k = json.dumps(x, sort_keys=True, ensure_ascii=False)
            if k not in seen:
                seen.add(k); out.append(x)
        return json.dumps(out, ensure_ascii=False, indent=1), True
    if isinstance(L, dict) and isinstance(R, dict):
        # dict thường: union key; conflict → local thắng
        out = dict(R); out.update(L)
        return json.dumps(out, ensure_ascii=False, indent=1), True
    return None, False


def try_merge_md(local_text: str, remote_text: str, rel: str):
    """Merge MD đơn giản: nếu một bản chứa toàn bộ bản kia → lấy bản dài hơn.
    Nếu khác hẳn → fallback (None, False)."""
    if remote_text.strip() in local_text:
        return local_text, True
    if local_text.strip() in remote_text:
        return remote_text, True
    # khác hẳn → fallback (giữ cả hai)
    return None, False


def apply_sync(cats, local_inv, remote_inv, log_lines):
    """Bước 2: áp dụng."""
    # LOCAL-ONLY → copy local → sidecar
    for rel in cats["LOCAL-ONLY"]:
        src = LOCAL / rel; dst = SIDECAR / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        log_lines.append(f"COPY local→sidecar: {rel}")
    # REMOTE-ONLY → copy sidecar → local
    for rel in cats["REMOTE-ONLY"]:
        src = SIDECAR / rel; dst = LOCAL / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        log_lines.append(f"COPY sidecar→local: {rel}")
    # CONFLICT-B → giữ cả hai, rename remote bản
    for rel in cats["CONFLICT-B"]:
        p = Path(rel)
        new_name = f"{p.stem}_remote{p.suffix}"
        new_rel = str(p.with_name(new_name)).replace("\\", "/")
        # ghi remote bản sang local với tên _remote
        src = SIDECAR / rel; dst = LOCAL / new_rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        # đảm bảo sidecar có cả bản local gốc + remote _remote
        # copy local sang sidecar đè bản remote ở đường dẫn gốc
        shutil.copy2(LOCAL / rel, SIDECAR / rel)
        # và đặt _remote trên sidecar = remote bản gốc cũ → nhưng sidecar đã thay gốc
        # → ghi _remote lại bằng copy từ local mới tạo
        shutil.copy2(LOCAL / new_rel, SIDECAR / new_rel)
        log_lines.append(f"CONFLICT-B keep-both: {rel} (giữ local) + {new_rel} (= remote)")
    # CONFLICT-T → thử merge
    for rel in cats["CONFLICT-T"]:
        ext = Path(rel).suffix.lower()
        ltxt = (LOCAL / rel).read_text(encoding="utf-8", errors="replace")
        rtxt = (SIDECAR / rel).read_text(encoding="utf-8", errors="replace")
        merged = None; ok = False
        if ext == ".json":
            merged, ok = try_merge_json(ltxt, rtxt, rel)
        elif ext == ".md":
            merged, ok = try_merge_md(ltxt, rtxt, rel)
        # nếu chưa OK → fallback giữ cả hai như CONFLICT-B
        if ok and merged is not None:
            (LOCAL / rel).write_text(merged, encoding="utf-8")
            shutil.copy2(LOCAL / rel, SIDECAR / rel)
            log_lines.append(f"MERGE {ext}: {rel}")
        else:
            p = Path(rel)
            new_name = f"{p.stem}_remote{p.suffix}"
            new_rel = str(p.with_name(new_name)).replace("\\", "/")
            shutil.copy2(SIDECAR / rel, LOCAL / new_rel)
            shutil.copy2(LOCAL / rel, SIDECAR / rel)
            shutil.copy2(LOCAL / new_rel, SIDECAR / new_rel)
            log_lines.append(f"CONFLICT-T keep-both (no auto-merge): {rel} + {new_rel}")


def main():
    apply = "--apply" in sys.argv

    print(f"Local : {LOCAL}")
    print(f"Sidecar: {SIDECAR}")
    if not SIDECAR.exists():
        print("!! Sidecar chưa có (chưa clone xong?)"); sys.exit(1)

    print("\n>>> Quét LOCAL...")
    local_inv = inventory(LOCAL)
    print(f"  {len(local_inv)} file")

    print("\n>>> Quét SIDECAR...")
    remote_inv = inventory(SIDECAR)
    print(f"  {len(remote_inv)} file")

    cats = categorize(local_inv, remote_inv)
    report(cats, local_inv, remote_inv, max_show=50)

    if not apply:
        print("\n[DRY-RUN] Không áp dụng. Chạy lại với --apply để thực thi.")
        return

    print("\n>>> ÁP DỤNG SYNC...")
    log_lines = [f"# CHANGELOG_sync — {datetime.now().isoformat()}", ""]
    apply_sync(cats, local_inv, remote_inv, log_lines)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path_local = LOCAL / f"CHANGELOG_sync_{ts}.md"
    log_path_local.write_text("\n".join(log_lines), encoding="utf-8")
    shutil.copy2(log_path_local, SIDECAR / log_path_local.name)
    print(f"\nĐã lưu changelog: {log_path_local}")
    print(f"Tổng hành động: {len(log_lines)-2}")


if __name__ == "__main__":
    main()
