# -*- coding: utf-8 -*-
"""Trích các đoạn nói về 8 lỗi cấu trúc Q2.5 từ transcript."""
import json, io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

TR = r"C:/Users/HLC/.claude/projects/c--Users-HLC-OneDrive-read-books-Lo-au/e2b0fc9b-f088-41b3-b3a6-7bc5d969c587.jsonl"
KEYS = ["8 lỗi cấu trúc", "lỗi cấu trúc Q2", "structural issue", "8 structural",
        "school-attachment", "school attachment", "Q2.5", "Q25_SEM",
        "duplicated Methods", "β-vs-OR", "untranslated journal"]

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
            score = sum(1 for k in KEYS if k in t)
            if score >= 2 and len(t) > 500:
                hits.append((ln, score, len(t), t))

hits.sort(key=lambda x: -x[1])
seen = set()
print(f"== {len(hits)} đoạn khớp ==\n")
for ln, sc, n, t in hits[:8]:
    if t[:150] in seen:
        continue
    seen.add(t[:150])
    print("=" * 70)
    print(f"[dòng {ln}, score {sc}, {n} ký tự]")
    print(t[:6000])
    print()
