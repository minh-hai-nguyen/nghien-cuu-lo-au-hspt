# -*- coding: utf-8 -*-
"""
Verify cite-TLTK consistency cho 2 bài KHGDVN.
Mỗi cite trong body phải có 1 TLTK tương ứng, và ngược lại.
"""
from docx import Document
from pathlib import Path
import re
import unicodedata

ROOT = Path(r"c:/Users/HLC/OneDrive/read_books/Lo-au")


def read_docx(path):
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def normalize_name(s):
    s = unicodedata.normalize("NFC", s).strip()
    return s


def extract_body_cites(text):
    """
    Extract cite-like patterns: 'X (YYYY)' or 'X và cộng sự (YYYY)' or '(X & cs, YYYY)' or '(X, YYYY)'.
    Returns list of (author_surface, year).
    """
    cites = set()

    # Pattern 1: "Author (YYYY)" or "Author và cộng sự (YYYY)"
    for m in re.finditer(r"([A-ZĐ][a-zA-Zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]+(?:\s[A-ZĐ][a-zA-Zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]+){0,4})\s+(?:và cộng sự|và cs|et al\.?)\s*\(((?:19|20)\d{2})\)", text):
        cites.add((normalize_name(m.group(1)), m.group(2)))

    # Pattern 2: "Author1, Author2 và cộng sự (YYYY)" - take first author
    # Pattern 3: "Author (YYYY)"
    for m in re.finditer(r"([A-ZĐ][a-zA-Zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]+(?:\s[A-ZĐ][a-zA-Zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]+){0,3})\s*\(((?:19|20)\d{2})[a-z]?(?:,\s*tr\.\s*\d+)?\)", text):
        cites.add((normalize_name(m.group(1)), m.group(2)))

    # Pattern 4: "(Author, YYYY)" or "(Author & cs, YYYY)"
    for m in re.finditer(r"\(([A-ZĐ][a-zA-Zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]+(?:\s[A-ZĐ][a-zA-Zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]+){0,3})\s*(?:&\s*cs|và cộng sự|và cs|et al\.?|&[^,)]+)?,\s*((?:19|20)\d{2})\)", text):
        cites.add((normalize_name(m.group(1)), m.group(2)))

    return cites


def extract_tltk_entries(body_text, tltk_section_keyword="Tài liệu tham khảo"):
    """Find TLTK section (after the keyword) and parse each entry to get (lead_author_surname, year)."""
    idx = body_text.find(tltk_section_keyword)
    if idx == -1:
        return []
    tltk_text = body_text[idx + len(tltk_section_keyword):]

    entries = []
    # Each entry roughly starts at beginning of line with author surname or all-caps org
    for line in tltk_text.split("\n"):
        line = line.strip()
        if len(line) < 20:
            continue
        # Try to extract first surname + year
        m = re.match(r"^([^.]+?)\.\s*\((\d{4})", line)
        if m:
            first_chunk = m.group(1).strip()
            year = m.group(2)
            # First surname = first word(s) before comma
            # English: "Lastname, F. M." → surname = "Lastname"
            # Vietnamese: "Hoàng Trung Học, ..." → surname = "Hoàng" (full name format)
            surname = first_chunk.split(",")[0].strip().split("&")[0].strip()
            entries.append((surname, year, line[:150]))
    return entries


def main():
    for name in ["Bai1_YTNC_HSTHCS_v1.docx", "Bai2_CanThiep_HSTHCS_v1.docx"]:
        path = ROOT / "bai-bao-khgdvn" / name
        print("=" * 70)
        print(f"Verifying: {name}")
        print("=" * 70)
        text = read_docx(path)

        # Split body and TLTK
        idx = text.find("Tài liệu tham khảo")
        if idx == -1:
            print("[WARN] Cannot find 'Tài liệu tham khảo' section")
            continue
        body = text[:idx]
        tltk_text = text[idx:]

        cites = extract_body_cites(body)
        tltk = extract_tltk_entries(text)

        print(f"\n[INFO] Body cites found: {len(cites)}")
        for c in sorted(cites):
            print(f"  {c[0]} ({c[1]})")

        print(f"\n[INFO] TLTK entries: {len(tltk)}")

        # Build TLTK surnames lookup
        tltk_keys = set()
        tltk_full = {}
        for surname, year, line in tltk:
            key = (surname.split()[0], year)  # first word only for matching
            tltk_keys.add(key)
            tltk_full[(surname, year)] = line

        # Check each body cite vs TLTK
        print(f"\n=== CITE → TLTK CHECK ===")
        missing_tltk = []
        for author, year in sorted(cites):
            # First word of author surface
            first = author.split()[0]
            last = author.split()[-1] if len(author.split()) > 1 else author
            found = False
            for sn, yr in tltk_keys:
                if yr == year and (sn == first or sn == last or sn.startswith(first[:5])):
                    found = True
                    break
            # Also check by checking if 'last' appears in any TLTK line
            if not found:
                for (sn, yr), line in tltk_full.items():
                    if yr == year and (author in line or last in line):
                        found = True
                        break
            if found:
                print(f"  ✓ {author} ({year})")
            else:
                print(f"  ✗ MISSING TLTK: {author} ({year})")
                missing_tltk.append((author, year))

        if missing_tltk:
            print(f"\n[!! ALERT] {len(missing_tltk)} cite không có TLTK tương ứng:")
            for a, y in missing_tltk:
                print(f"      → {a} ({y})")

        # Reverse: TLTK entries that body doesn't cite?
        print(f"\n=== TLTK → BODY CHECK (TLTK có nhưng body không cite?) ===")
        cited_years = {y for _, y in cites}
        cited_keywords = " ".join(a for a, _ in cites)
        for (sn, yr, line) in tltk:
            first = sn.split()[0]
            # check if first or any other notable keyword from line appears in body cites
            matched = False
            for body_author, body_year in cites:
                if yr == body_year:
                    if first in body_author or any(part in body_author for part in sn.split()):
                        matched = True
                        break
                    # Also surname format check
                    last = body_author.split()[-1]
                    if last == first or last == sn.split()[-1]:
                        matched = True
                        break
            if not matched:
                # Just print first 100 chars
                print(f"  ⚠ TLTK uncited?: {line[:120]}")


if __name__ == "__main__":
    main()
