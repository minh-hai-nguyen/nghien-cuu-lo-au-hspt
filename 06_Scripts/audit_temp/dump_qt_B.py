# -*- coding: utf-8 -*-
"""Dump các mục RAG QT đợt 2 (QT034-QT067) để dựng script sửa."""
import json, io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

RAG = "06_Scripts/questions_kg_data/rag_main_index.json"
IDS = ["QT034", "QT040", "QT044", "QT045", "QT047", "QT057", "QT058",
       "QT060", "QT061", "QT062", "QT063", "QT064", "QT065", "QT066", "QT067"]

rag = json.load(open(RAG, encoding="utf-8"))
ent = {(it.get("meta", {}).get("id") or it.get("id", "")): it for it in rag["entries"]}

for pid in IDS:
    it = ent.get(pid)
    print("#" * 60, pid)
    if it is None:
        print("(KHONG CO TRONG INDEX)")
        continue
    print(it.get("text", ""))
    print()
