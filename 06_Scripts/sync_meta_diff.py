# -*- coding: utf-8 -*-
"""
Đồng bộ Lo-au qua GitHub API (KHÔNG clone toàn bộ).
Bước 1 (mặc định): so metadata + dry-run.
Bước 2 (--download): tải các file remote-only và conflict.
Bước 3 (--apply): áp dụng sync sau khi đã có remote_files/.

Tính git-blob-SHA: sha1(b"blob <size>\\0" + content) — giống hệt git internal.
"""
import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

LOCAL = Path(r"c:/Users/HLC/OneDrive/read_books/Lo-au")
SIDECAR = Path(r"C:/Users/HLC/Documents/sync-lo-au")
REMOTE_TREE = SIDECAR / "remote_tree.json"
REMOTE_FILES = SIDECAR / "remote_files"  # tải về đây

REPO = "minh-hai-nguyen/nghien-cuu-lo-au-hspt"

TEXT_EXTS = {".md", ".txt", ".py", ".json", ".csv", ".html", ".css", ".js",
             ".xml", ".yaml", ".yml", ".ini", ".cfg", ".sh", ".graphml", ".log"}

SKIP_NAMES = {"desktop.ini", "Thumbs.db", ".DS_Store"}
SKIP_PREFIXES = ("~$",)


def is_skipped(rel_path: Path) -> bool:
    if ".git" in rel_path.parts:
        return True
    n = rel_path.name
    return n in SKIP_NAMES or any(n.startswith(p) for p in SKIP_PREFIXES)


def git_blob_sha(path: Path) -> str:
    """Tính git-blob SHA giống `git hash-object <file>`."""
    size = path.stat().st_size
    h = hashlib.sha1()
    h.update(f"blob {size}\0".encode())
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def load_remote_tree() -> dict:
    """Trả về dict relpath -> (sha, size)."""
    d = json.loads(REMOTE_TREE.read_text(encoding="utf-8"))
    out = {}
    for t in d["tree"]:
        if t["type"] != "blob":
            continue
        out[t["path"]] = (t["sha"], t.get("size", 0))
    return out


def inventory_local() -> dict:
    """Trả về dict relpath -> (git_blob_sha, size)."""
    out = {}
    n = 0
    for p in LOCAL.rglob("*"):
        if not p.is_file():
            continue
        rel = p.relative_to(LOCAL)
        if is_skipped(rel):
            continue
        try:
            out[str(rel).replace("\\", "/")] = (git_blob_sha(p), p.stat().st_size)
        except Exception as e:
            print(f"  !! {rel}: {e}", file=sys.stderr)
        n += 1
        if n % 500 == 0:
            print(f"  ...quét {n} file local", file=sys.stderr, flush=True)
    return out


def fmt(n: int) -> str:
    for u in ["B", "KB", "MB", "GB"]:
        if n < 1024:
            return f"{n:.1f}{u}"
        n /= 1024
    return f"{n:.1f}TB"


def categorize(loc: dict, rem: dict):
    out = {"SAME": [], "LOCAL-ONLY": [], "REMOTE-ONLY": [],
           "CONFLICT-B": [], "CONFLICT-T": []}
    all_paths = set(loc) | set(rem)
    for p in sorted(all_paths):
        lv = loc.get(p)
        rv = rem.get(p)
        if lv and rv:
            if lv[0] == rv[0]:
                out["SAME"].append(p)
            else:
                ext = Path(p).suffix.lower()
                if ext in TEXT_EXTS:
                    out["CONFLICT-T"].append(p)
                else:
                    out["CONFLICT-B"].append(p)
        elif lv:
            out["LOCAL-ONLY"].append(p)
        else:
            out["REMOTE-ONLY"].append(p)
    return out


def report(cats, loc, rem, max_show=40):
    print("\n" + "=" * 70)
    print("DRY-RUN — so metadata (KHÔNG tải file)")
    print("=" * 70)
    for k in ["SAME", "LOCAL-ONLY", "REMOTE-ONLY", "CONFLICT-B", "CONFLICT-T"]:
        print(f"  {k:13s} {len(cats[k]):5d} file")
    btr = sum(loc[p][1] for p in cats["LOCAL-ONLY"])
    btl = sum(rem[p][1] for p in cats["REMOTE-ONLY"])
    bcb = sum(rem[p][1] for p in cats["CONFLICT-B"])
    bct = sum(rem[p][1] for p in cats["CONFLICT-T"])
    print(f"\n  Byte LOCAL→remote (push): {fmt(btr)}")
    print(f"  Byte REMOTE→local (cần TẢI VỀ): {fmt(btl)}")
    print(f"  Byte CONFLICT BINARY (tải về để giữ _remote): {fmt(bcb)}")
    print(f"  Byte CONFLICT TEXT (tải về để merge): {fmt(bct)}")
    print(f"  TỔNG cần tải về từ remote: {fmt(btl + bcb + bct)}")
    for cat in ["REMOTE-ONLY", "CONFLICT-B", "CONFLICT-T", "LOCAL-ONLY"]:
        items = cats[cat]
        if not items:
            continue
        print(f"\n--- {cat} (hiển thị {min(max_show, len(items))}/{len(items)}) ---")
        for p in items[:max_show]:
            sz = (loc if cat == "LOCAL-ONLY" else rem).get(p, ("", 0))[1]
            print(f"  {fmt(sz):>8s}  {p}")


def download_blobs(paths_with_sha):
    """Tải các blob qua gh api blobs/<sha>. Lưu vào REMOTE_FILES/<relpath>."""
    REMOTE_FILES.mkdir(parents=True, exist_ok=True)
    fail = 0
    for i, (rel, sha) in enumerate(paths_with_sha):
        dst = REMOTE_FILES / rel
        if dst.exists():
            continue  # đã tải
        dst.parent.mkdir(parents=True, exist_ok=True)
        # gh api: trả JSON {content: base64, encoding: base64}
        try:
            r = subprocess.run(
                ["gh", "api", f"repos/{REPO}/git/blobs/{sha}",
                 "--jq", ".content"],
                capture_output=True, text=False, check=True)
            import base64
            b64 = r.stdout.decode().strip()
            content = base64.b64decode(b64)
            dst.write_bytes(content)
        except Exception as e:
            print(f"  !! lỗi tải {rel}: {e}", file=sys.stderr)
            fail += 1
        if (i + 1) % 50 == 0:
            print(f"  ...đã tải {i+1}/{len(paths_with_sha)}", flush=True)
    print(f"Hoàn tất tải. Thất bại: {fail}")
    return fail


def main():
    mode = "diff"
    if "--download" in sys.argv:
        mode = "download"
    if "--apply" in sys.argv:
        mode = "apply"

    print(f"Local : {LOCAL}")
    print(f"Sidecar: {SIDECAR}")
    print(f"Mode : {mode}")
    print()

    if not REMOTE_TREE.exists():
        print("!! Chưa có remote_tree.json — chạy gh api trước.")
        sys.exit(1)

    print(">>> Đọc remote tree...")
    rem = load_remote_tree()
    print(f"  {len(rem)} blob")

    print("\n>>> Quét local + tính git-blob-SHA...")
    loc = inventory_local()
    print(f"  {len(loc)} file")

    cats = categorize(loc, rem)
    report(cats, loc, rem)

    # lưu phân loại để bước sau dùng
    cats_path = SIDECAR / "categories.json"
    cats_save = {k: list(v) for k, v in cats.items()}
    cats_save["_loc"] = {p: list(v) for p, v in loc.items()}
    cats_save["_rem"] = {p: list(v) for p, v in rem.items()}
    cats_path.write_text(json.dumps(cats_save, ensure_ascii=False, indent=1),
                         encoding="utf-8")
    print(f"\nĐã lưu phân loại: {cats_path}")

    if mode == "download":
        # Tải file remote-only + conflict (B + T)
        to_dl = []
        for p in cats["REMOTE-ONLY"]:
            to_dl.append((p, rem[p][0]))
        for p in cats["CONFLICT-B"] + cats["CONFLICT-T"]:
            to_dl.append((p, rem[p][0]))
        print(f"\n>>> Tải về {len(to_dl)} file...")
        download_blobs(to_dl)


if __name__ == "__main__":
    main()
