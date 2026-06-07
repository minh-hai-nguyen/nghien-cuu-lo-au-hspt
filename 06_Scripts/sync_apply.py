# -*- coding: utf-8 -*-
"""
Áp dụng sync sau khi đã tải xong remote files vào sidecar/remote_files/.
- REMOTE-ONLY: copy remote → local.
- CONFLICT-B: rename bản remote ở local thành <stem>_remote<ext>; bản local giữ nguyên.
- CONFLICT-T: thử merge .json/.md; fallback giữ cả hai.
- LOCAL-ONLY: không động chạm (giữ nguyên ở local — sẽ push).

Cũng sinh CHANGELOG_sync_<ts>.md ở local + sidecar.
"""
import hashlib, json, os, shutil, sys
from datetime import datetime
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

LOCAL = Path(r"c:/Users/HLC/OneDrive/read_books/Lo-au")
SIDECAR = Path(r"C:/Users/HLC/Documents/sync-lo-au")
REMOTE_FILES = SIDECAR / "remote_files"
CATS_FILE = SIDECAR / "categories.json"


def load_cats():
    return json.loads(CATS_FILE.read_text(encoding="utf-8"))


def try_merge_json(local_text, remote_text):
    try:
        L = json.loads(local_text); R = json.loads(remote_text)
    except Exception:
        return None, False
    if isinstance(L, dict) and isinstance(R, dict) and "entries" in L and "entries" in R:
        def key_of(e):
            m = e.get("meta") or {}
            return m.get("id") or e.get("id") or json.dumps(e, sort_keys=True)[:50]
        Lm = {key_of(e): e for e in L["entries"]}
        Rm = {key_of(e): e for e in R["entries"]}
        all_keys = list(Lm) + [k for k in Rm if k not in Lm]
        merged = []
        for k in all_keys:
            le, re = Lm.get(k), Rm.get(k)
            if le and re:
                merged.append(le if len(json.dumps(le, ensure_ascii=False))
                              >= len(json.dumps(re, ensure_ascii=False)) else re)
            else:
                merged.append(le or re)
        out = dict(L); out["entries"] = merged
        if isinstance(out.get("meta"), dict) and "n_entries" in out["meta"]:
            out["meta"]["n_entries"] = len(merged)
        return json.dumps(out, ensure_ascii=False, indent=1), True
    if isinstance(L, list) and isinstance(R, list):
        seen = set(); o = []
        for x in L + R:
            k = json.dumps(x, sort_keys=True, ensure_ascii=False)
            if k not in seen:
                seen.add(k); o.append(x)
        return json.dumps(o, ensure_ascii=False, indent=1), True
    if isinstance(L, dict) and isinstance(R, dict):
        out = dict(R); out.update(L)
        return json.dumps(out, ensure_ascii=False, indent=1), True
    return None, False


def try_merge_md(local_text, remote_text):
    if remote_text.strip() in local_text:
        return local_text, True
    if local_text.strip() in remote_text:
        return remote_text, True
    # khác hẳn → fallback
    return None, False


def try_merge_py(local_text, remote_text):
    # py: thường thì local thắng (script tôi vừa sửa); fallback giữ cả hai
    if remote_text.strip() in local_text or local_text.strip() in remote_text:
        return local_text if len(local_text) >= len(remote_text) else remote_text, True
    return None, False


def main():
    cats = load_cats()
    log = [f"# CHANGELOG_sync — {datetime.now().isoformat()}", ""]
    n_remote_only = 0; n_conflict_b = 0; n_conflict_t_merge = 0; n_conflict_t_keep = 0
    miss = 0

    print(">>> REMOTE-ONLY: copy remote → local")
    for rel in cats["REMOTE-ONLY"]:
        src = REMOTE_FILES / rel
        dst = LOCAL / rel
        if not src.exists():
            miss += 1; print(f"  !! không có blob đã tải: {rel}"); continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        n_remote_only += 1
        log.append(f"COPY remote→local: {rel}")
    print(f"  copy {n_remote_only}/{len(cats['REMOTE-ONLY'])} (miss {miss})")

    print("\n>>> CONFLICT-B: rename bản remote thành <stem>_remote<ext>")
    for rel in cats["CONFLICT-B"]:
        src = REMOTE_FILES / rel
        if not src.exists():
            miss += 1; print(f"  !! không có blob: {rel}"); continue
        p = Path(rel)
        new_name = f"{p.stem}_remote{p.suffix}"
        new_rel = str(p.with_name(new_name)).replace("\\", "/")
        dst = LOCAL / new_rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        n_conflict_b += 1
        log.append(f"CONFLICT-B keep-both: giữ {rel} (= local) + tạo {new_rel} (= remote)")
    print(f"  rename {n_conflict_b}/{len(cats['CONFLICT-B'])}")

    print("\n>>> CONFLICT-T: thử merge")
    for rel in cats["CONFLICT-T"]:
        ext = Path(rel).suffix.lower()
        src = REMOTE_FILES / rel
        if not src.exists():
            miss += 1; print(f"  !! không có blob: {rel}"); continue
        ltxt = (LOCAL / rel).read_text(encoding="utf-8", errors="replace")
        rtxt = src.read_text(encoding="utf-8", errors="replace")
        merged, ok = (None, False)
        if ext == ".json":
            merged, ok = try_merge_json(ltxt, rtxt)
        elif ext == ".md":
            merged, ok = try_merge_md(ltxt, rtxt)
        elif ext == ".py":
            merged, ok = try_merge_py(ltxt, rtxt)
        if ok and merged is not None:
            (LOCAL / rel).write_text(merged, encoding="utf-8")
            n_conflict_t_merge += 1
            log.append(f"MERGE {ext}: {rel}")
            print(f"  MERGE OK: {rel}")
        else:
            p = Path(rel)
            new_name = f"{p.stem}_remote{p.suffix}"
            new_rel = str(p.with_name(new_name)).replace("\\", "/")
            shutil.copy2(src, LOCAL / new_rel)
            n_conflict_t_keep += 1
            log.append(f"CONFLICT-T keep-both: giữ {rel} (= local) + tạo {new_rel} (= remote)")
            print(f"  KEEP BOTH: {rel}")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOCAL / f"CHANGELOG_sync_{ts}.md"
    log_path.write_text("\n".join(log), encoding="utf-8")
    print(f"\nĐã lưu changelog: {log_path}")

    print(f"\nTÓM TẮT:")
    print(f"  REMOTE-ONLY copy về: {n_remote_only}")
    print(f"  CONFLICT-B rename:    {n_conflict_b}")
    print(f"  CONFLICT-T merge OK:  {n_conflict_t_merge}")
    print(f"  CONFLICT-T keep-both: {n_conflict_t_keep}")
    print(f"  Thiếu blob (chưa tải): {miss}")
    print(f"  Tổng hành động ghi log: {len(log) - 2}")


if __name__ == "__main__":
    main()
