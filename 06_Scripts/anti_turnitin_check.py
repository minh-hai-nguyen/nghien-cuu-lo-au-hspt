# -*- coding: utf-8 -*-
"""
Anti-Turnitin check cho Bài 1 + Bài 2 KHGDVN.

Cách kiểm tra:
1. Trích xuất text từ 2 file Bai1/Bai2 .docx
2. So với 3 nguồn có thể có trên Turnitin DB:
   - Bài CONG THI HANG VA CS (đã extract sang txt)
   - Narrative_Review_OUTLINE_G1.md (có khả năng dựa luận án CTH)
   - Tất cả file trong 01_Bao-cao/ (các bản đã gửi thầy CTH)
3. Tìm cụm 8-gram (8 từ liên tiếp) trùng nhau — đó là dấu hiệu Turnitin sẽ bắt.
"""
from docx import Document
from pathlib import Path
import re
import unicodedata
from collections import Counter

ROOT = Path(r"c:/Users/HLC/OneDrive/read_books/Lo-au")


def normalize(text):
    """Lowercase, remove punctuation, collapse whitespace."""
    text = unicodedata.normalize("NFC", text.lower())
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def read_docx(path):
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def read_text(path):
    return path.read_text(encoding="utf-8")


def ngrams(text, n=8):
    words = text.split()
    return set(" ".join(words[i:i+n]) for i in range(len(words) - n + 1))


def check_overlap(target_text, source_texts, n=8):
    """Tìm n-gram trong target xuất hiện trong bất kỳ source nào."""
    target_grams = ngrams(target_text, n)
    source_grams = set()
    for src in source_texts:
        source_grams.update(ngrams(src, n))
    overlap = target_grams & source_grams
    return overlap, len(target_grams), len(source_grams)


def main():
    print("=" * 70)
    print("ANTI-TURNITIN CHECK — Bài 1 + Bài 2 KHGDVN")
    print("=" * 70)

    # Targets
    bai1 = read_docx(ROOT / "bai-bao-khgdvn/Bai1_YTNC_HSTHCS_v1.docx")
    bai2 = read_docx(ROOT / "bai-bao-khgdvn/Bai2_CanThiep_HSTHCS_v1.docx")
    bai1_n = normalize(bai1)
    bai2_n = normalize(bai2)

    print(f"\n[INFO] Bài 1: {len(bai1.split())} words raw")
    print(f"[INFO] Bài 2: {len(bai2.split())} words raw")

    # Sources to check against
    sources = []

    # 1. Bài CONG THI HANG VA CS
    cth_path = ROOT / "bai-bao-khgdvn/cong_thi_hang_extracted.txt"
    if cth_path.exists():
        cth = read_text(cth_path)
        sources.append(("Bài CONG THI HANG (Công Thị Hằng + Đào Minh Đức 2026)", normalize(cth)))
        print(f"[INFO] Source 1 (CTH bài): {len(cth.split())} words")

    # 2. Narrative_Review_OUTLINE_G1.md
    nr_path = ROOT / "01_Bao-cao/Narrative_Review_OUTLINE_G1.md"
    if nr_path.exists():
        nr = read_text(nr_path)
        sources.append(("Narrative_Review_OUTLINE_G1.md", normalize(nr)))
        print(f"[INFO] Source 2 (Outline G1): {len(nr.split())} words")

    # 3. Báo cáo CTH các bản (sample 3 file lớn nhất 01_Bao-cao)
    bao_cao_dir = ROOT / "01_Bao-cao"
    docx_files = list(bao_cao_dir.glob("*.docx"))
    # Skip backup + lock files
    docx_files = [f for f in docx_files if not f.name.startswith("~$") and "BACKUP" not in f.name]
    # Take a few representative files
    sample_files = sorted(docx_files, key=lambda f: f.stat().st_size, reverse=True)[:5]
    for f in sample_files:
        try:
            txt = read_docx(f)
            sources.append((f.name, normalize(txt)))
            print(f"[INFO] Source: {f.name} ({len(txt.split())} words)")
        except Exception as e:
            print(f"[WARN] Failed to read {f.name}: {e}")

    # Check overlap
    for target_name, target_text in [("Bài 1", bai1_n), ("Bài 2", bai2_n)]:
        print(f"\n{'=' * 70}")
        print(f"Target: {target_name}")
        print(f"{'=' * 70}")

        target_grams = ngrams(target_text, 8)
        print(f"  Total 8-grams in target: {len(target_grams)}")

        # Per-source overlap
        all_overlap_grams = set()
        for src_name, src_text in sources:
            src_grams = ngrams(src_text, 8)
            overlap = target_grams & src_grams
            all_overlap_grams |= overlap
            if overlap:
                pct = len(overlap) / max(1, len(target_grams)) * 100
                print(f"  [{src_name}]")
                print(f"    Overlap: {len(overlap)} 8-grams ({pct:.2f}%)")
                # Show 3 sample overlap
                for g in list(overlap)[:5]:
                    print(f"      • '{g[:120]}'")

        # Total estimated overlap
        total_pct = len(all_overlap_grams) / max(1, len(target_grams)) * 100
        print(f"\n  TỔNG estimated overlap (union across sources): {len(all_overlap_grams)} 8-grams ({total_pct:.2f}%)")
        if total_pct < 5:
            print("  ✓ AN TOÀN — overlap thấp, ít rủi ro Turnitin > 25%")
        elif total_pct < 10:
            print("  ⚠ THẬN TRỌNG — overlap trung bình, có thể cần sửa một số đoạn")
        else:
            print("  ✗ NGUY HIỂM — overlap cao, cần paraphrase lại các đoạn trùng")


if __name__ == "__main__":
    main()
