# -*- coding: utf-8 -*-
"""
Sửa các TẦNG DỮ LIỆU xây trên nền RAG/KG còn sót lỗi lan truyền (17/05/2026):
- rag_main_index.json (3 bản): 3 chỗ "Zhameden" cross-reference còn sót -> Yin và cộng sự.
- rag_authors_index.json (2 bản web): descriptor QT013 "Zhameden_..." -> "Yin_...".
- papers_metadata.json (2 bản web): QT013 tên tác giả; năm QT026/058/061; lead_author = tạp chí/tổ chức.
Đã kiểm tra: rag_questions_index.json, glossary.json, priority_authors.json, authors_normalized.json,
questions_kg.json, vjes_kg.json — KHÔNG có lỗi lan truyền (các marker khớp đều là dương tính giả).
"""
import json, os, io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

RAG_MAIN = [
    "06_Scripts/questions_kg_data/rag_main_index.json",
    "tro-ly-nghien-cuu-tam-ly/web/data/rag_main_index.json",
    "tro-ly-nghien-cuu-tam-ly-light/web/data/rag_main_index.json",
]
RAG_AUTH = [
    "tro-ly-nghien-cuu-tam-ly/web/data/rag_authors_index.json",
    "tro-ly-nghien-cuu-tam-ly-light/web/data/rag_authors_index.json",
]
PMETA = [
    "tro-ly-nghien-cuu-tam-ly/web/data/papers_metadata.json",
    "tro-ly-nghien-cuu-tam-ly-light/web/data/papers_metadata.json",
]

# --- 1. rag_main_index: 3 chỗ Zhameden cross-ref còn sót ---
RAGFIX = {
 "QT024": [("và Zhameden 2025 (can thiệp trường LMIC).",
            "và Yin và cộng sự 2025 (can thiệp trường LMIC).")],
 "QT028": [("cho lo âu tại LMIC (Zhameden) — gợi ý",
            "cho lo âu tại LMIC (Yin và cộng sự) — gợi ý")],
 "QT029": [("xác nhận khoảng trống Zhameden 2025 đã chỉ ra.",
            "xác nhận khoảng trống Yin và cộng sự 2025 đã chỉ ra.")],
}
for rp in RAG_MAIN:
    if not os.path.exists(rp):
        print(f"(bỏ qua) {rp}"); continue
    rag = json.load(open(rp, encoding="utf-8"))
    ent = {(it.get("meta", {}).get("id") or it.get("id", "")): it for it in rag["entries"]}
    ok = 0
    for pid, repls in RAGFIX.items():
        it = ent.get(pid)
        if not it:
            continue
        t = it.get("text", "")
        for old, new in repls:
            if old in t:
                t = t.replace(old, new); ok += 1
            else:
                print(f"  !! rag_main {pid}: không khớp -> {old[:45]}")
        it["text"] = t
    json.dump(rag, open(rp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"rag_main_index {rp.split('/')[0]}: {ok}/3 chỗ Zhameden -> Yin")

# --- 2. rag_authors_index: descriptor QT013 ---
for ap in RAG_AUTH:
    if not os.path.exists(ap):
        print(f"(bỏ qua) {ap}"); continue
    ai = json.load(open(ap, encoding="utf-8"))
    n = 0
    for it in ai["entries"]:
        for k in ("text_en", "text_vn"):
            if k in it and "Zhameden_2025_PLOSONE_LMIC" in it[k]:
                it[k] = it[k].replace("Zhameden_2025_PLOSONE_LMIC", "Yin_2025_PLOSONE_LMIC")
                n += 1
    json.dump(ai, open(ap, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"rag_authors_index {ap.split('/')[0]}: sửa {n} trường descriptor QT013")

# --- 3. papers_metadata: tên tác giả + năm ---
# {id: {field: (old, new)}}
PM = {
 "QT013": {"descriptor": ("Zhameden_2025_PLOSONE_LMIC", "Yin_2025_PLOSONE_LMIC"),
           "lead_author": ("Zhameden", "Yin")},
 "QT024": {"lead_author": ("WHO", "Tarasenko")},
 "QT026": {"lead_author": ("NHS", "Baker"), "year": (2025, 2024)},
 "QT028": {"lead_author": ("AJP", "Zugman")},
 "QT029": {"lead_author": ("BMC", "Li")},
 "QT058": {"year": (2022, 2023)},
 "QT061": {"year": (2024, 2022)},
}
for pp in PMETA:
    if not os.path.exists(pp):
        print(f"(bỏ qua) {pp}"); continue
    pm = json.load(open(pp, encoding="utf-8"))
    idx = {e.get("id"): e for e in pm["entries"]}
    nf = 0
    for pid, fields in PM.items():
        e = idx.get(pid)
        if not e:
            print(f"  !! papers_metadata: không có {pid}"); continue
        for field, (old, new) in fields.items():
            if e.get(field) == old:
                e[field] = new; nf += 1
            else:
                print(f"  !! {pid}.{field}: giá trị hiện tại {e.get(field)!r} != {old!r}")
    json.dump(pm, open(pp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"papers_metadata {pp.split('/')[0]}: sửa {nf} trường")

print("[DONE]")
