# -*- coding: utf-8 -*-
"""
QA Round 1 cho Q2.5 paper EN + VN:
- Word count 4-way (Frontiers 5000-8000)
- Cite-TLTK consistency
- Cross-plagiarism vs 3 own published papers
- Cross EN ↔ VN structural similarity
"""
from docx import Document
from pathlib import Path
import re
import unicodedata

ROOT = Path(r"c:/Users/HLC/OneDrive/read_books/Lo-au")
EN = ROOT / "bai-bao-khgdvn/Q25_SEM_Pathways_EN_v3.docx"
VN = ROOT / "bai-bao-khgdvn/Q25_SEM_Pathways_VN_v3.docx"

# Source papers for self-plagiarism check
SOURCES = [
    ("Bài 1 VJES YTNC", ROOT / "bai-bao-khgdvn/Bai1_YTNC_HSTHCS_v4_14052026.docx"),
    ("Bài 2 VJES Can thiệp", ROOT / "bai-bao-khgdvn/Bai2_CanThiep_HSTHCS_v4_14052026.docx"),
    ("CTH validation", ROOT / "bai-bao-khgdvn/cong_thi_hang_extracted.txt"),
]


def read_docx_text(path):
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def read_text(path):
    return path.read_text(encoding="utf-8") if path.exists() else ""


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


def split_paper(text):
    """Split into title/abstract/body/refs based on heading markers."""
    sections = {"title": "", "abstract": "", "keywords": "", "body": "", "refs": ""}
    lines = [l for l in text.split("\n") if l.strip()]
    if not lines:
        return sections
    sections["title"] = lines[0]
    state = "head"
    body_lines = []
    refs_lines = []
    for line in lines[1:]:
        if line.startswith("Abstract:") or line.startswith("Tóm tắt:"):
            sections["abstract"] = line
        elif line.startswith("Keywords:") or line.startswith("Từ khóa:"):
            sections["keywords"] = line
        elif line.startswith("1. ") and state == "head":
            state = "body"
            body_lines.append(line)
        elif state == "body":
            if line == "References" or line == "Tài liệu tham khảo":
                state = "refs"
            else:
                body_lines.append(line)
        elif state == "refs":
            refs_lines.append(line)
    sections["body"] = "\n".join(body_lines)
    sections["refs"] = "\n".join(refs_lines)
    return sections


def extract_cites(body):
    """Extract (author, year) cites — supports EN + VN patterns."""
    cites = set()
    patterns = [
        # Author et al. (YYYY) or "and cộng sự (YYYY)"
        r"([A-Z][a-zA-ZÀ-ỹ]+(?:[\s-][A-ZĐ][a-zA-ZÀ-ỹ]+){0,3})\s+(?:et al\.?|and colleagues|và cộng sự)\s*\((\d{4})\)",
        # Author (YYYY)
        r"([A-Z][a-zA-ZÀ-ỹ]+(?:[\s-][A-ZĐ][a-zA-ZÀ-ỹ]+){0,3})\s*\((\d{4})[a-z]?\)",
        # (Author, YYYY) or (Author et al., YYYY)
        r"\(([A-Z][a-zA-ZÀ-ỹ]+(?:[\s-][A-ZĐ][a-zA-ZÀ-ỹ]+){0,3})(?:\s+et al\.?|\s+và cộng sự|\s*&[^,)]+)?,?\s*(\d{4})[a-z]?\)",
    ]
    for pat in patterns:
        for m in re.finditer(pat, body):
            cites.add((m.group(1).strip(), m.group(2)))
    return cites


def main():
    print("=" * 80)
    print("QA ROUND 1 — Q2.5 paper EN + VN")
    print("=" * 80)

    docs = {"EN": read_docx_text(EN), "VN": read_docx_text(VN)}
    parsed = {lang: split_paper(text) for lang, text in docs.items()}

    for lang in ["EN", "VN"]:
        s = parsed[lang]
        print(f"\n----- {lang} -----")
        print(f"  Title: {wc(s['title'])} words ({len(s['title'])} chars)")
        print(f"  Abstract: {wc(s['abstract'])} words")
        print(f"  Keywords: {wc(s['keywords'])} words")
        print(f"  Body: {wc(s['body'])} words")
        print(f"  Refs: {wc(s['refs'])} words")
        head = wc(s['title']) + wc(s['abstract']) + wc(s['keywords'])
        body = wc(s['body'])
        refs = wc(s['refs'])
        all_words = head + body + refs
        # Frontiers in Psychiatry word limit: 5,000–8,000 (original research)
        ways = {
            "1) BODY only": body,
            "2) BODY+REFS": body + refs,
            "3) BODY+HEAD": body + head,
            "4) ALL": all_words,
        }
        for k, v in ways.items():
            ok = "✓" if 5000 <= v <= 8000 else "✗"
            print(f"  {k}: {v} {ok}")

        cites = extract_cites(s['body'])
        print(f"  Body cites: {len(cites)}")

    # Cross-plagiarism EN ↔ VN (expected high due to translation)
    print("\n" + "-" * 80)
    print("Cross EN ↔ VN structural overlap (expected — same paper)")
    print("-" * 80)
    en_grams = ngrams(parsed["EN"]["body"], 8)
    vn_grams = ngrams(parsed["VN"]["body"], 8)
    overlap = en_grams & vn_grams
    print(f"  EN 8-grams: {len(en_grams)}")
    print(f"  VN 8-grams: {len(vn_grams)}")
    print(f"  Overlap: {len(overlap)} (expected near 0 due to language difference)")

    # Anti-self-plagiarism vs 3 published papers
    print("\n" + "=" * 80)
    print("ANTI-SELF-PLAGIARISM vs 3 published papers (EN body)")
    print("=" * 80)
    en_body_grams = en_grams
    for label, path in SOURCES:
        if path.suffix == ".txt":
            src_text = read_text(path)
        else:
            src_text = read_docx_text(path)
        src_grams = ngrams(src_text, 8)
        ov = en_body_grams & src_grams
        pct = len(ov) / max(1, len(en_body_grams)) * 100
        print(f"  vs {label}: {len(ov)} 8-grams overlap ({pct:.2f}%)")
        if ov:
            for g in sorted(ov)[:5]:
                print(f"    • {g[:120]}")

    print("\n" + "=" * 80)
    print("ANTI-SELF-PLAGIARISM vs 3 published papers (VN body)")
    print("=" * 80)
    vn_body_grams = vn_grams
    for label, path in SOURCES:
        if path.suffix == ".txt":
            src_text = read_text(path)
        else:
            src_text = read_docx_text(path)
        src_grams = ngrams(src_text, 8)
        ov = vn_body_grams & src_grams
        pct = len(ov) / max(1, len(vn_body_grams)) * 100
        print(f"  vs {label}: {len(ov)} 8-grams overlap ({pct:.2f}%)")
        if ov and pct > 0.5:
            for g in sorted(ov)[:5]:
                print(f"    • {g[:120]}")


if __name__ == "__main__":
    main()
