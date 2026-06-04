# -*- coding: utf-8 -*-
"""Trích báo cáo agent audit QT batch 1,3,4,6,7 còn thiếu."""
import json, io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

TR = r"C:/Users/HLC/.claude/projects/c--Users-HLC-OneDrive-read-books-Lo-au/e2b0fc9b-f088-41b3-b3a6-7bc5d969c587.jsonl"

# nhận diện báo cáo audit batch QT
def is_qt_report(t):
    if len(t) < 600:
        return False
    has_batch = ("QT Batch" in t or "QT batch" in t or "DATA-INTEGRITY AUDIT" in t
                 or "Data-Integrity Audit" in t)
    # đếm số mã QT0xx
    import re
    n = len(set(re.findall(r"QT0\d\d", t)))
    return has_batch and n >= 2

hits = []
with open(TR, encoding="utf-8") as f:
    for ln, line in enumerate(f):
        line = line.strip()
        if not line:
            continue
        try:
            obj = json.loads(line)
        except Exception:
            continue
        msg = obj.get("message", {})
        content = msg.get("content")
        chunks = []
        if isinstance(content, str):
            chunks.append(content)
        elif isinstance(content, list):
            for c in content:
                if isinstance(c, dict):
                    if c.get("type") == "text":
                        chunks.append(c.get("text", ""))
                    elif c.get("type") == "tool_result":
                        tc = c.get("content")
                        if isinstance(tc, str):
                            chunks.append(tc)
                        elif isinstance(tc, list):
                            for x in tc:
                                if isinstance(x, dict) and x.get("type") == "text":
                                    chunks.append(x.get("text", ""))
        for t in chunks:
            if is_qt_report(t):
                hits.append((ln, t))

# loại trùng
seen = set()
uniq = []
for ln, t in hits:
    key = t[:200]
    if key in seen:
        continue
    seen.add(key)
    uniq.append((ln, t))

print(f"== {len(uniq)} báo cáo QT batch ==\n")
for ln, t in uniq:
    print("=" * 70)
    print(f"[dòng {ln}]")
    print(t)
    print()
