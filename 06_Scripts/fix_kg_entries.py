# -*- coding: utf-8 -*-
"""
Sửa KG sau audit QT (16/05/2026).
- kg_triples.json: QT015 sai công cụ đo; loại các REPORTED_ES gán nhầm từ phần
  đối chiếu liên bài; loại HAS_N sai của QT044; sửa HAS_N QT045.
- author_kg_v1.json (3 bản): nhãn bài QT013 "Zhameden_..." -> "Yin_..."
questions_kg.json và vjes_kg.json: đã kiểm tra — KHÔNG có lỗi, không sửa.
"""
import json, io, sys, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# ---------- 1. kg_triples.json ----------
TRP = "06_Scripts/kg_data/kg_triples.json"
tr = json.load(open(TRP, encoding="utf-8"))

# các triple ES gán nhầm (số liệu thực ra của bài khác, lấy nhầm từ bảng đối chiếu)
DROP_ES = {
    ("QT022", "REPORTED_ES", "ES::QT022::OR::5.0"),     # 5.0 là OR game của QT007 Chen
    ("QT024", "REPORTED_ES", "ES::QT024::OR::0.562"),   # 0.562 là OR của Wen 2020
    ("QT026", "REPORTED_ES", "ES::QT026::AOR::2.17"),   # 2.17 là AOR lo âu của QT023
    ("QT028", "REPORTED_ES", "ES::QT028::OR::0.562"),   # 0.562 là OR của Wen 2020
    ("QT029", "REPORTED_ES", "ES::QT029::OR::0.562"),   # 0.562 là OR của Wen 2020
}
new = []
drop_es = drop_qt015 = drop_n = 0
for t in tr:
    s, p, o = t["subject"], t["predicate"], t["object"]
    if (s, p, o) in DROP_ES:
        drop_es += 1
        continue
    # QT015 dùng CES-D, KHÔNG phải GAD-7/PHQ-9
    if s == "QT015" and p == "USED_SCALE" and o in ("Scale::GAD-7", "Scale::PHQ-9"):
        drop_qt015 += 1
        continue
    # QT015 chỉ đo trầm cảm — bỏ Anxiety/GAD
    if s == "QT015" and p == "MEASURED" and o in (
            "Outcome::Generalized Anxiety (GAD)", "Outcome::Anxiety"):
        drop_qt015 += 1
        continue
    # QT044 (MA resilience) HAS_N::10 là số NC từ Mỹ, không phải cỡ mẫu
    if s == "QT044" and p == "HAS_N":
        drop_n += 1
        continue
    # QT045: HAS_N 38 là nhánh can thiệp; tổng tuyển là 77
    if s == "QT045" and p == "HAS_N" and o == "SampleSize::38":
        t = dict(t); t["object"] = "SampleSize::77"
    new.append(t)

# thêm công cụ đo đúng cho QT015
new.append({"subject": "QT015", "predicate": "USED_SCALE", "object": "Scale::CES-D"})

json.dump(new, open(TRP, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(f"kg_triples.json: bỏ {drop_es} ES gán nhầm, {drop_qt015} triple QT015 sai, "
      f"{drop_n} HAS_N sai; thêm CES-D; QT045 HAS_N->77. Tổng triple: {len(tr)} -> {len(new)}")

# ---------- 2. author_kg_v1.json (3 bản) ----------
AKS = [
    "06_Scripts/author_kg_data/author_kg_v1.json",
    "tro-ly-nghien-cuu-tam-ly/web/data/author_kg_v1.json",
    "tro-ly-nghien-cuu-tam-ly-light/web/data/author_kg_v1.json",
]
for ap in AKS:
    if not os.path.exists(ap):
        print(f"(bỏ qua) {ap}")
        continue
    ak = json.load(open(ap, encoding="utf-8"))
    fixed = 0
    for n in ak["nodes"]:
        if n.get("id") == "QT013" and n.get("label") == "Zhameden_2025_PLOSONE_LMIC":
            n["label"] = "Yin_2025_PLOSONE_LMIC"
            fixed += 1
    json.dump(ak, open(ap, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"{ap}: sửa nhãn QT013 = {fixed}")

print("[DONE]")
