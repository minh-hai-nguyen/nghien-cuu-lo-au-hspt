# -*- coding: utf-8 -*-
"""
VJES word count compliance check.
VJES rule: 5.000 ≤ total ≤ 7.000 từ (cả bài, 8-12 trang).
Đếm theo NHIỀU CÁCH để đảm bảo dưới mọi cách diễn giải đều OK.
"""
from docx import Document
from pathlib import Path
import re

ROOT = Path(r"c:/Users/HLC/OneDrive/read_books/Lo-au")


def wc(text):
    """Whitespace-split word count, strip empty."""
    return len([w for w in text.split() if w.strip()])


def get_paragraphs(path):
    doc = Document(path)
    return [p.text for p in doc.paragraphs if p.text.strip()]


def find_section_index(paras, keyword):
    for i, p in enumerate(paras):
        if keyword in p:
            return i
    return -1


def analyze(path):
    paras = get_paragraphs(path)
    print(f"\n{'='*70}\n{path.name}\n{'='*70}")
    print(f"Total paragraphs: {len(paras)}")

    # Identify sections
    idx_title_vn = 0  # first paragraph is title VN (after rebuild)
    idx_tom_tat = next((i for i, p in enumerate(paras) if p.startswith("Tóm tắt")), -1)
    idx_tu_khoa = next((i for i, p in enumerate(paras) if p.startswith("Từ khóa")), -1)
    idx_title_en = -1
    for i, p in enumerate(paras):
        if i > idx_tu_khoa and (p.lower().startswith("risk factors") or p.lower().startswith("research gaps") or p.startswith("Risk") or "junior secondary" in p.lower()):
            idx_title_en = i
            break
    idx_abstract = next((i for i, p in enumerate(paras) if p.startswith("Abstract")), -1)
    idx_keywords = next((i for i, p in enumerate(paras) if p.startswith("Keywords")), -1)
    idx_dat_van_de = next((i for i, p in enumerate(paras) if p.startswith("1. Đặt vấn đề")), -1)
    idx_ket_luan = next((i for i, p in enumerate(paras) if p.startswith("4. Kết luận")), -1)
    idx_tltk = next((i for i, p in enumerate(paras) if "Tài liệu tham khảo" in p and len(p) < 30), -1)

    print(f"Title VN at index: {idx_title_vn}")
    print(f"Tóm tắt at: {idx_tom_tat}")
    print(f"Từ khóa at: {idx_tu_khoa}")
    print(f"Title EN at: {idx_title_en}")
    print(f"Abstract EN at: {idx_abstract}")
    print(f"Keywords at: {idx_keywords}")
    print(f"1. Đặt vấn đề at: {idx_dat_van_de}")
    print(f"4. Kết luận at: {idx_ket_luan}")
    print(f"Tài liệu tham khảo at: {idx_tltk}")

    # Component texts
    title_vn = paras[idx_title_vn]
    tom_tat = paras[idx_tom_tat]  # includes "Tóm tắt: ..."
    tu_khoa = paras[idx_tu_khoa]
    title_en = paras[idx_title_en] if idx_title_en != -1 else ""
    abstract_en = paras[idx_abstract] if idx_abstract != -1 else ""
    keywords = paras[idx_keywords] if idx_keywords != -1 else ""

    # Body: from "1. Đặt vấn đề" to before TLTK
    body_paras = paras[idx_dat_van_de:idx_tltk]
    body_text = "\n".join(body_paras)

    # TLTK
    tltk_paras = paras[idx_tltk+1:]
    tltk_text = "\n".join(tltk_paras)

    # Counts
    n_title_vn = wc(title_vn)
    n_title_en = wc(title_en)
    n_tom_tat = wc(tom_tat)  # includes "Tóm tắt:" prefix; not significant
    n_tu_khoa = wc(tu_khoa)
    n_abstract_en = wc(abstract_en)
    n_keywords = wc(keywords)
    n_body = wc(body_text)
    n_tltk = wc(tltk_text)
    n_tltk_entries = len(tltk_paras)

    print(f"\n=== Component word counts ===")
    print(f"  Title VN:        {n_title_vn:>6} words   ({len(title_vn)} chars)")
    print(f"  Tóm tắt VN:      {n_tom_tat:>6} words")
    print(f"  Từ khóa VN:      {n_tu_khoa:>6} words")
    print(f"  Title EN:        {n_title_en:>6} words")
    print(f"  Abstract EN:     {n_abstract_en:>6} words")
    print(f"  Keywords EN:     {n_keywords:>6} words")
    print(f"  BODY (§1-§4):    {n_body:>6} words")
    print(f"  TLTK ({n_tltk_entries} entries): {n_tltk:>6} words")

    # Various ways to compute total
    head_block = n_title_vn + n_tom_tat + n_tu_khoa + n_title_en + n_abstract_en + n_keywords
    print(f"\n=== Đếm theo các cách (5.000 ≤ X ≤ 7.000 áp dụng cho ?) ===")
    way1 = n_body
    way2 = n_body + n_tltk
    way3 = n_body + head_block
    way4 = n_body + head_block + n_tltk
    print(f"  1) BODY only:                          {way1:>6}  {'✓' if 5000 <= way1 <= 7000 else '✗'}")
    print(f"  2) BODY + TLTK:                        {way2:>6}  {'✓' if 5000 <= way2 <= 7000 else '✗'}")
    print(f"  3) BODY + Title + Abstract + Keywords: {way3:>6}  {'✓' if 5000 <= way3 <= 7000 else '✗'}")
    print(f"  4) ALL (Title+Abstract+Body+TLTK):     {way4:>6}  {'✓' if 5000 <= way4 <= 7000 else '✗'}")

    # VJES interpretation = total submission (the literal text of guideline)
    print(f"\n  ⭐ VJES interpretation = #4 ALL = {way4}")
    if 5000 <= way4 <= 7000:
        print(f"     STATUS: ✓ PASS — within 5.000-7.000")
    elif way4 > 7000:
        print(f"     STATUS: ✗ OVER LIMIT — cần cắt {way4 - 7000} từ")
    else:
        print(f"     STATUS: ✗ UNDER LIMIT — cần thêm {5000 - way4} từ")

    return way4


for name in ["Bai1_YTNC_HSTHCS_v1.docx", "Bai2_CanThiep_HSTHCS_v1.docx"]:
    analyze(ROOT / "bai-bao-khgdvn" / name)
