# -*- coding: utf-8 -*-
"""Trích các báo cáo agent audit QT từ transcript jsonl (phục hồi sau nén context)."""
import json, io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

TR = r"C:/Users/HLC/.claude/projects/c--Users-HLC-OneDrive-read-books-Lo-au/e2b0fc9b-f088-41b3-b3a6-7bc5d969c587.jsonl"

# dấu hiệu báo cáo audit QT của agent
KEYS = ["EXISTS-CORRECT", "CLAIM-MISMATCH", "NOT-FOUND", "BIBLIO-WRONG",
        "rag_QT_batch", "audit QT", "QT RAG"]

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
        # gom mọi text trong message
        def texts(o):
            out = []
            msg = o.get("message", {})
            content = msg.get("content")
            if isinstance(content, str):
                out.append(content)
            elif isinstance(content, list):
                for c in content:
                    if isinstance(c, dict):
                        if c.get("type") == "text":
                            out.append(c.get("text", ""))
                        elif c.get("type") == "tool_result":
                            tc = c.get("content")
                            if isinstance(tc, str):
                                out.append(tc)
                            elif isinstance(tc, list):
                                for x in tc:
                                    if isinstance(x, dict) and x.get("type") == "text":
                                        out.append(x.get("text", ""))
            return out
        for t in texts(obj):
            if any(k in t for k in KEYS) and len(t) > 400:
                hits.append((ln, len(t), t))

print(f"== {len(hits)} đoạn khớp ==\n")
for ln, n, t in hits:
    print("=" * 70)
    print(f"[dòng {ln}, {n} ký tự]")
    print(t)
    print()
