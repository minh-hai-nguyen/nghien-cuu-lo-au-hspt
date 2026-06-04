# -*- coding: utf-8 -*-
"""QC 5 vòng cho Bài 1 & Bài 2 v2:
1. Citation ↔ Reference: in-text citation phải có entry trong TLTK
2. Năm + tác giả + tạp chí logic
3. Số liệu: liệt kê các effect size / số liệu để kiểm tra
4. Cross-overlap 2 bài
5. Red-flag claim có thể bịa
"""
import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from collections import Counter

FILES = {
    'B1': r'c:\Users\OS\OneDrive\read_books\Lo-au\bai-bao-khgdvn\Bai1_YTNC_HSTHCS_v3_14052026.docx',
    'B2': r'c:\Users\OS\OneDrive\read_books\Lo-au\bai-bao-khgdvn\Bai2_CanThiep_HSTHCS_v3_14052026.docx',
}


def get_text(path):
    d = Document(path)
    return [p.text for p in d.paragraphs]


def split_body_refs(paras):
    """Tách body và phần TLTK."""
    body = []
    refs = []
    in_refs = False
    for p in paras:
        if p.strip().startswith('Tài liệu tham khảo'):
            in_refs = True
            continue
        if in_refs:
            refs.append(p)
        else:
            body.append(p)
    return body, refs


# =================================================================
# VÒNG 1: Citation ↔ Reference
# =================================================================
def extract_citations(body_text):
    """Tìm các citation (Author, năm) và (Author và cs, năm) trong body."""
    cites = set()
    # Pattern: Tác giả (năm) hoặc Tác giả và cs. (năm) hoặc Author et al. (năm)
    patterns = [
        r'([A-ZĐÀÁẢẠÂẦẤẨẬĂẰẮẲẶÈÉẺẸÊỀẾỂỆÌÍỈỊÒÓỎỌÔỒỐỔỘƠỜỚỞỢÙÚỦỤƯỪỨỬỰỲÝỶỴ][a-zA-ZÀ-ỹ-]+(?:\s+[A-ZĐÀÁẢẠÂẦẤẨẬĂẰẮẲẶÈÉẺẸÊỀẾỂỆÌÍỈỊÒÓỎỌÔỒỐỔỘƠỜỚỞỢÙÚỦỤƯỪỨỬỰỲÝỶỴ][a-zA-ZÀ-ỹ-]+){0,4})(?:\s+(?:và\s+cộng\s+sự|và\s+cs|et\s+al\.?|và\s+cs\.))?\s*\((\d{4})\)',
        r'\(([A-ZĐÀÁẢẠÂẦẤẨẬĂẰẮẲẶ][a-zA-ZÀ-ỹ-]+(?:\s+[A-Z][a-zA-ZÀ-ỹ-]+){0,4})(?:\s*&?\s*[a-zA-ZÀ-ỹ]+)?\s*&?\s*(?:cộng\s+sự|cs\.?|et\s+al\.?)?,?\s*(\d{4})\)',
    ]
    for pat in patterns:
        for m in re.finditer(pat, body_text):
            author = m.group(1).strip()
            year = m.group(2)
            cites.add((author, year))
    return cites


def extract_ref_keys(refs):
    """Trích key (first author surname + year) từ mỗi reference entry."""
    keys = set()
    for r in refs:
        if not r.strip(): continue
        # Pattern: Surname, Initial[, ...]. (Year). or Surname A B C, (Year).
        m = re.match(r'^([A-ZĐÀÁẢẠÂẦẤẨẬĂẰẮẲẶÈÉẺẸÊỀẾỂỆÌÍỈỊÒÓỎỌÔỒỐỔỘƠỜỚỞỢÙÚỦỤƯỪỨỬỰỲÝỶỴ][a-zA-ZÀ-ỹ-]+)', r)
        m_year = re.search(r'\((\d{4})\)', r)
        if m and m_year:
            keys.add((m.group(1), m_year.group(1)))
    return keys


print('='*70)
print('VÒNG 1: Citation ↔ Reference list')
print('='*70)

for label, path in FILES.items():
    paras = get_text(path)
    body, refs = split_body_refs(paras)
    body_text = '\n'.join(body)
    cites = extract_citations(body_text)
    ref_keys = extract_ref_keys(refs)

    # Find orphans — smart matching for Vietnamese names + romanization variants
    def normalize_name(n):
        n = n.lower()
        # Common romanization swaps
        for a, b in [('dương', 'duong'), ('hoàng', 'hoang'), ('ngô', 'ngo'),
                     ('nguyễn', 'nguyen'), ('phạm', 'pham'), ('trần', 'tran'),
                     ('lê', 'le'), ('võ', 'vo'), ('đỗ', 'do'), ('hồ', 'ho'),
                     ('đặng', 'dang'), ('bùi', 'bui')]:
            n = n.replace(a, b)
        # Remove diacritics roughly
        import unicodedata
        n = ''.join(c for c in unicodedata.normalize('NFD', n) if unicodedata.category(c) != 'Mn')
        return n.strip()

    def match_cite_to_ref(cite_name, cite_year, ref_keys):
        """Return True if there's a matching reference."""
        cite_norm = normalize_name(cite_name)
        cite_words = cite_norm.split()
        for ref_name, ref_year in ref_keys:
            if ref_year != cite_year: continue
            ref_norm = normalize_name(ref_name)
            # Match if cite first/last word in ref OR ref in cite
            if ref_norm in cite_norm or cite_norm in ref_norm:
                return True
            # Match by last word (surname for Western, given name for VN)
            if cite_words and cite_words[-1] == ref_norm.split()[0]:
                return True
            if cite_words and cite_words[0] == ref_norm.split()[0]:
                return True
        return False

    cite_surnames = {}
    for c in cites:
        cite_surnames[(c[0], c[1])] = c
    ref_dict = {(r[0], r[1]): r for r in ref_keys}

    print(f'\n--- {label} ---')
    print(f'  Citations in body: {len(cites)}')
    print(f'  References in list: {len(ref_keys)}')

    # Print suspected mismatches (sau khi smart matching)
    orphans = []
    for c in cites:
        if not match_cite_to_ref(c[0], c[1], ref_keys):
            orphans.append(c)
    unused = []
    for r in ref_keys:
        # check if any citation matches this ref
        matched = False
        for c in cites:
            if match_cite_to_ref(c[0], c[1], [(r[0], r[1])]):
                matched = True; break
        if not matched:
            unused.append(r)

    if orphans:
        print(f'  ⚠ Citations TRONG BODY nhưng KHÔNG có trong TLTK ({len(orphans)}):')
        for o in sorted(orphans):
            print(f'    - {o[0]} ({o[1]})')
    if unused:
        print(f'  ⚠ TLTK có nhưng KHÔNG được cite trong body ({len(unused)}):')
        for u in sorted(unused):
            print(f'    - {u[0]} ({u[1]})')


# =================================================================
# VÒNG 2: Năm + tác giả logic — flag năm bất thường
# =================================================================
print('\n' + '='*70)
print('VÒNG 2: Sanity check năm + flag năm tương lai / quá xa')
print('='*70)

for label, path in FILES.items():
    paras = get_text(path)
    body, refs = split_body_refs(paras)
    body_text = '\n'.join(body)
    cites = extract_citations(body_text)
    print(f'\n--- {label} ---')
    years = [int(c[1]) for c in cites]
    if years:
        print(f'  Năm citations range: {min(years)} - {max(years)}')
        # Flag năm tương lai (sau 2026)
        future = [c for c in cites if int(c[1]) > 2026]
        old = [c for c in cites if int(c[1]) < 1990]
        if future:
            print(f'  ⚠ Năm tương lai (>2026): {future}')
        if old:
            print(f'  ⚠ Năm quá xa (<1990): {old}')


# =================================================================
# VÒNG 3: Liệt kê số liệu để kiểm tra fact
# =================================================================
print('\n' + '='*70)
print('VÒNG 3: Tổng hợp số liệu để fact-check thủ công')
print('='*70)

NUMERIC_PATTERNS = [
    (r'g\s*=\s*([\d,\.]+)', 'effect size g'),
    (r'SMD\s*=\s*([\d,\.]+)', 'SMD'),
    (r'β\s*=\s*[-−]?([\d,\.]+)', 'beta coefficient'),
    (r'OR\s*=\s*([\d,\.]+)', 'OR'),
    (r'r\s*=\s*[-−]?([\d,\.]+)', 'r'),
    (r'(\d+(?:[,\.]\d+)?)\s*%', 'percent'),
    (r'(\d+(?:[,\.]\d+)?)\s*triệu', 'triệu'),
    (r'(\d{3,})\s+(?:học sinh|trẻ em|vị thành niên|cặp|mẫu)', 'sample size'),
]

for label, path in FILES.items():
    paras = get_text(path)
    body, _ = split_body_refs(paras)
    body_text = '\n'.join(body)
    print(f'\n--- {label} ---')
    for pat, name in NUMERIC_PATTERNS:
        matches = re.findall(pat, body_text)
        if matches:
            # unique values
            def safe_float(x):
                try: return float(x.replace(',', '.').rstrip('.'))
                except: return 0
            uniq = sorted(set(matches), key=safe_float)
            print(f'  {name}: {uniq[:15]}{"..." if len(uniq)>15 else ""}')


# =================================================================
# VÒNG 4: Cross-overlap giữa 2 bài v2
# =================================================================
print('\n' + '='*70)
print('VÒNG 4: Cross-plagiarism v2 vs v2 (chi tiết)')
print('='*70)

def get_sents(path):
    return [s.strip() for s in re.split(r'(?<=[.!?])\s+', '\n'.join(get_text(path)))
            if len(s.split()) >= 6]


def ngrams(s, n=7):
    w = s.split()
    return set(' '.join(w[i:i+n]) for i in range(len(w)-n+1)) if len(w) >= n else set()


s1 = get_sents(FILES['B1'])
s2 = get_sents(FILES['B2'])

b1_idx = {}
for i, s in enumerate(s1):
    for ng in ngrams(s):
        b1_idx.setdefault(ng, []).append(i)

overlaps = []
for j, s2_sent in enumerate(s2):
    matched = sum(1 for ng in ngrams(s2_sent) if ng in b1_idx)
    if matched >= 3:
        overlaps.append((j, s2_sent, matched))
overlaps.sort(key=lambda x: -x[2])
print(f'\nCâu Bài 2 trùng ≥3 7-gram với Bài 1: {len(overlaps)}')
print('Top 10 overlap MẠNH cần fix tiếp:')
for j, s, n in overlaps[:10]:
    is_ref_or_title = ('Article' in s or 'Journal' in s or 'doi.org' in s or s.startswith('Khoảng trống'))
    flag = ' [TLTK/title — OK]' if is_ref_or_title else ' [⚠ BODY — cần fix]'
    print(f'\n[B2 sent {j}, {n} 7-grams]{flag}')
    print(f'  {s[:200]}')


# =================================================================
# VÒNG 5: Red-flag claims có thể AI bịa
# =================================================================
print('\n' + '='*70)
print('VÒNG 5: Red-flag claims có thể AI bịa — cần thầy verify thủ công')
print('='*70)

RED_FLAGS = [
    (r'(\d{4,})\s*(?:bài|nghiên cứu|thử nghiệm)', 'số mẫu/N có thể bịa'),
    (r'(?:khoảng|gần)\s+(\d+)%', 'percent không citation'),
    (r'(?:dao động|từ)\s+(\d[\d,\.]+)\s*(?:đến|–|-)\s*(\d[\d,\.]+)', 'dải số'),
    (r'(?:200|300|500|1000)\s+(?:chuyên gia|nhà|người)', 'ước lượng số chuyên gia/người'),
]

# Liệt kê các câu KHÔNG có citation kế bên
def sentences_without_citation(paras):
    text = '\n'.join(paras)
    sents = re.split(r'(?<=[.!?])\s+', text)
    flagged = []
    for s in sents:
        if len(s.split()) < 8: continue
        # Check if sentence has citation pattern
        has_cite = bool(re.search(r'\(\w+[^()]{0,80}\d{4}\)|\w+\s+(?:và\s+cộng\s+sự|et\s+al\.?)\s*\(\d{4}\)', s))
        # Check if sentence has specific numeric claim
        has_number = bool(re.search(r'\d+%|\d{3,}|g\s*=|SMD\s*=|β\s*=|r\s*=|OR\s*=', s))
        if has_number and not has_cite:
            flagged.append(s.strip())
    return flagged

for label, path in FILES.items():
    paras = get_text(path)
    body, _ = split_body_refs(paras)
    flagged = sentences_without_citation(body)
    print(f'\n--- {label}: Câu có số liệu nhưng KHÔNG có citation kế bên ({len(flagged)}) ---')
    for s in flagged[:8]:
        print(f'  ⚠ {s[:200]}')

print('\n' + '='*70)
print('QC complete. Em rà soát + báo cáo các điểm bất thường cho thầy.')
print('='*70)
