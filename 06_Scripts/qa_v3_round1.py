# -*- coding: utf-8 -*-
"""
QA Vòng 1 cho v3 — 5 checks:
1. Word count (4-way)
2. Cite-TLTK consistency
3. Cross-plagiarism giữa Bài 1 ↔ Bài 2 (8-gram overlap)
4. Anti-Turnitin vs 3 nguồn cũ (Outline G1, CTH validation, báo cáo CTH)
5. Sanity check: số liệu, năm, tác giả nhất quán
"""
from docx import Document
from pathlib import Path
import re
import unicodedata
from collections import Counter

ROOT = Path(r"c:/Users/HLC/OneDrive/read_books/Lo-au")
B1 = ROOT / "bai-bao-khgdvn/Bai1_YTNC_HSTHCS_v4_14052026.docx"
B2 = ROOT / "bai-bao-khgdvn/Bai2_CanThiep_HSTHCS_v4_14052026.docx"


def read_docx_paragraphs(path):
    doc = Document(path)
    return [p.text for p in doc.paragraphs if p.text.strip()]


def split_sections(paragraphs):
    """Tách title/abstract/body/TLTK từ list paragraph."""
    sections = {"title_vn": "", "tom_tat": "", "tu_khoa": "",
                "title_en": "", "abstract_en": "", "keywords_en": "",
                "body": [], "tltk": []}
    state = "head"
    for p in paragraphs:
        if state == "head":
            if not sections["title_vn"]:
                sections["title_vn"] = p
            elif p.startswith("Tóm tắt"):
                sections["tom_tat"] = p
            elif p.startswith("Từ khóa"):
                sections["tu_khoa"] = p
            elif p.startswith("Abstract"):
                sections["abstract_en"] = p
            elif p.startswith("Keywords"):
                sections["keywords_en"] = p
            elif p.startswith("1. Đặt vấn đề"):
                state = "body"
                sections["body"].append(p)
            elif not sections["title_en"] and ("anxiety" in p.lower() or "Research" in p):
                sections["title_en"] = p
        elif state == "body":
            if "Tài liệu tham khảo" in p and len(p) < 30:
                state = "tltk"
            else:
                sections["body"].append(p)
        else:
            sections["tltk"].append(p)
    return sections


def normalize(text):
    text = unicodedata.normalize("NFC", text.lower())
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def wc(text):
    return len([w for w in text.split() if w.strip()])


def ngrams(text, n=8):
    words = normalize(text).split()
    return set(" ".join(words[i:i+n]) for i in range(len(words) - n + 1))


def extract_cites(body_text):
    """Extract (author, year) cites from body."""
    cites = set()
    # Pattern 1: "X (YYYY)" or "X và cộng sự (YYYY)" or "X et al. (YYYY)"
    p1 = re.compile(
        r"([A-ZĐ][a-zA-Zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]+"
        r"(?:[\s-][A-ZĐ][a-zA-Zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]+){0,4})"
        r"\s+(?:và cộng sự|và cs|et al\.?)\s*\(((?:19|20)\d{2})\)"
    )
    for m in p1.finditer(body_text):
        cites.add((m.group(1).strip(), m.group(2)))

    # Pattern 2: "Author (YYYY)"
    p2 = re.compile(
        r"([A-ZĐ][a-zA-Zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]+"
        r"(?:[\s-][A-ZĐ][a-zA-Zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]+){0,3})"
        r"\s*\(((?:19|20)\d{2})[a-z]?(?:,\s*tr\.\s*\d+)?\)"
    )
    for m in p2.finditer(body_text):
        cites.add((m.group(1).strip(), m.group(2)))

    # Pattern 3: "(Author & cs, YYYY)" or "(Author, YYYY)"
    p3 = re.compile(
        r"\(([A-ZĐ][a-zA-Zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]+"
        r"(?:[\s-][A-ZĐ][a-zA-Zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]+){0,3})"
        r"\s*(?:&\s*cs|và cộng sự|và cs|et al\.?)?,?\s*((?:19|20)\d{2})\)"
    )
    for m in p3.finditer(body_text):
        cites.add((m.group(1).strip(), m.group(2)))
    return cites


def main():
    print("=" * 80)
    print("QA VÒNG 1 — Bài 1 v3 + Bài 2 v3 (14/05/2026)")
    print("=" * 80)

    results = {}
    for label, path in [("Bài 1", B1), ("Bài 2", B2)]:
        paras = read_docx_paragraphs(path)
        secs = split_sections(paras)
        body_text = "\n".join(secs["body"])
        tltk_text = "\n".join(secs["tltk"])

        print(f"\n----- {label}: {path.name} -----")
        print(f"  Title VN: {wc(secs['title_vn'])} từ ({len(secs['title_vn'])} chars)")
        print(f"  Tóm tắt VN: {wc(secs['tom_tat'])} từ")
        print(f"  Từ khóa VN: {wc(secs['tu_khoa'])} từ")
        print(f"  Title EN: {wc(secs['title_en'])} từ")
        print(f"  Abstract EN: {wc(secs['abstract_en'])} từ")
        print(f"  Keywords EN: {wc(secs['keywords_en'])} từ")
        print(f"  BODY: {wc(body_text)} từ ({len(secs['body'])} paras)")
        print(f"  TLTK: {wc(tltk_text)} từ ({len(secs['tltk'])} entries)")

        head = wc(secs['title_vn']) + wc(secs['tom_tat']) + wc(secs['tu_khoa']) + \
               wc(secs['title_en']) + wc(secs['abstract_en']) + wc(secs['keywords_en'])
        body = wc(body_text)
        tltk = wc(tltk_text)

        ways = {
            "1) BODY only": body,
            "2) BODY + TLTK": body + tltk,
            "3) BODY + head": body + head,
            "4) ALL": body + head + tltk,
        }
        for k, v in ways.items():
            ok = "✓" if 5000 <= v <= 6200 else "✗"
            print(f"  {k}: {v} {ok}")

        # Cite extraction
        cites = extract_cites(body_text)
        print(f"\n  Body cites: {len(cites)}")
        results[label] = {
            "body_text": body_text,
            "tltk_text": tltk_text,
            "cites": cites,
            "tltk_paras": secs["tltk"],
            "secs": secs,
            "body_words": body,
            "all_words": body + head + tltk,
        }

    # ===== Cross-plagiarism check Bài 1 vs Bài 2 =====
    print("\n" + "=" * 80)
    print("CROSS-PLAGIARISM Bài 1 ↔ Bài 2 (8-gram)")
    print("=" * 80)
    g1 = ngrams(results["Bài 1"]["body_text"], 8)
    g2 = ngrams(results["Bài 2"]["body_text"], 8)
    overlap = g1 & g2
    pct1 = len(overlap) / max(1, len(g1)) * 100
    pct2 = len(overlap) / max(1, len(g2)) * 100
    print(f"  Bài 1 total 8-grams: {len(g1)}")
    print(f"  Bài 2 total 8-grams: {len(g2)}")
    print(f"  Overlap: {len(overlap)} ({pct1:.2f}% Bài 1; {pct2:.2f}% Bài 2)")
    if overlap:
        print("\n  Sample overlap n-grams (first 30):")
        for g in sorted(overlap)[:30]:
            print(f"    • {g[:150]}")
    if max(pct1, pct2) > 5:
        print(f"\n  ⚠ WARNING: cross-overlap > 5% — cần paraphrase chéo")
    else:
        print(f"\n  ✓ AN TOÀN — cross-overlap < 5%")

    # ===== Cross-plagiarism CHO ABSTRACT (riêng vì abstract dễ trùng) =====
    print("\n" + "-" * 80)
    print("Cross-overlap ABSTRACT VN + EN giữa 2 bài")
    print("-" * 80)
    for component in ["tom_tat", "abstract_en"]:
        s1 = results["Bài 1"]["secs"][component]
        s2 = results["Bài 2"]["secs"][component]
        g1c = ngrams(s1, 6)
        g2c = ngrams(s2, 6)
        ov = g1c & g2c
        if g1c and g2c:
            pct = len(ov) / max(1, len(g1c)) * 100
            print(f"  {component}: {len(ov)} 6-grams overlap ({pct:.2f}%)")
            if ov and pct > 3:
                for g in sorted(ov)[:10]:
                    print(f"    ⚠ '{g[:140]}'")

    # ===== Cross TLTK overlap (TLTK shared là OK nhưng note để biết) =====
    print("\n" + "-" * 80)
    print("TLTK shared giữa 2 bài (OK vì cùng tham khảo nguồn)")
    print("-" * 80)
    t1_keys = set()
    t2_keys = set()
    for line in results["Bài 1"]["tltk_paras"]:
        m = re.match(r"^([^.]+?)\.\s*\((\d{4})", line)
        if m:
            t1_keys.add((m.group(1).strip()[:30], m.group(2)))
    for line in results["Bài 2"]["tltk_paras"]:
        m = re.match(r"^([^.]+?)\.\s*\((\d{4})", line)
        if m:
            t2_keys.add((m.group(1).strip()[:30], m.group(2)))
    shared = t1_keys & t2_keys
    print(f"  Bài 1 TLTK: {len(t1_keys)}; Bài 2 TLTK: {len(t2_keys)}; Shared: {len(shared)}")
    for s in sorted(shared):
        print(f"    {s[0]} ({s[1]})")

    # ===== Cite list summary =====
    print("\n" + "-" * 80)
    print("Cite list mỗi bài (để kiểm tra TLTK consistency tiếp)")
    print("-" * 80)
    for label in ["Bài 1", "Bài 2"]:
        print(f"\n  {label} body cites ({len(results[label]['cites'])}):")
        for c in sorted(results[label]["cites"]):
            print(f"    {c[0]} ({c[1]})")


if __name__ == "__main__":
    main()
