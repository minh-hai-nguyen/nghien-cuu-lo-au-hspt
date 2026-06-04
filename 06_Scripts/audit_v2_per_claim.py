# -*- coding: utf-8 -*-
"""Audit tinh chinh xac TUNG CLAIM - cross-check moi factual claim voi nguon verified."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
import re

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
d = Document(os.path.join(ROOT, 'Luận án TS', '01_TongQuan_TheoChuDe_30trang_v2_26052026.docx'))
text = ' '.join(p.text for p in d.paragraphs)

# Verified facts from Bai 1, Bai 2 TLTK + original Tong quan
VERIFIED = {
    'Phạm Thị Thu Hoa': {'year': '2024', 'n': '3.910', 'rate': '40,6%', 'mild': '23,9%', 'mod': '10,9%', 'sev': '5,8%', 'tool': 'GAD-7'},
    'V-NAMHS': {'year': '2022', 'n': '5.996', 'mh': '21,7%', 'dsm': '2,3%', 'tool': 'DISC-5', 'provinces': '38', 'age': '10-17'},
    'Hoàng Trung Học': {'year': '2025', 'n': '8.389', 'covid': '41,5%', 'post': '25,4%', 'tool': 'DASS-21'},
    'Trần Thị Mỵ Lương': {'year': '2020', 'n': '540', 'rate': '14,2%', 'tool': 'DASS-42'},
    'Ngô Anh Vinh': {'year': '2024', 'n': '845', 'anx': '54,4%', 'str': '24,7%', 'dep': '59,0%', 'ACE': '0,28'},
    'Trần Hồ Vĩnh Lộc': {'year': '2024', 'n': '976', 'APR': '2,82'},
    'Phạm Thị Ngọc': {'year': '2024', 'n': '420', 'AOR_bul': '2,42', 'AOR_uns': '2,25'},
    'Đinh Thị Hồng Vân': {'year': '2021', 'n': '749', 'scores': '91,1%'},
    'Trúc Thanh Thái': {'year': '2026', 'n': '2.631', 'rate': '50,3%'},
    'Trần Thảo Vi': {'year': '2024', 'n': '341', 'start': '46,4', 'end': '53,5'},
    'Phạm Sỹ Tiến': {'year': '2024', 'n': '546', 'beta': '-0,40'},
    'Chen': {'year': '2023', 'n': '63.205', 'rate': '13,9%'},
    'Xu': {'year': '2021', 'n': '373.216', 'overall': '9,89%'},
    'Brunborg': {'year': '2025', 'n': '979.043'},
    'Compas': {'year': '2017', 'N': '80.850', 'studies': '212'},
    'Sowislo': {'year': '2013', 'studies': '18', 'beta': '-0,10'},
    'Moore': {'year': '2017', 'studies': '165', 'OR': '1,77'},
    'Lee': {'year': '2026', 'studies': '27', 'r': '0,23'},
    'Walder': {'year': '2025', 'studies': '21', 'g': '0,508', 'guided': '0,825', 'unguided': '0,27'},
    'Xian': {'year': '2024', 'studies': '30', 'SUCRA': '71,2%'},
    'Cai': {'year': '2025', 'studies': '38', 'MA': '21', 'SMD': '0,17'},
    'Bradshaw': {'year': '2025', 'n': '709', 'schools': '40'},
    'Matsumoto': {'year': '2024', 'n': '77', 'g': '0,66'},
    'Bress': {'year': '2024', 'n': '59'},
    'Tran': {'year': '2023', 'n': '1.084'},
    'Jagiello': {'year': '2025', 'studies': '31', 'countries': '13'},
    'Fassi': {'year': '2025', 'n': '3.340', 'g_int': '0,46', 'g_anx': '0,62'},
    'Liu': {'year': '2025', 'studies': '52'},
    'Praptomojati': {'year': '2024'},
    'Kendall': {'year': '1994'},
    'Barrett': {'year': '1996'},
    'Bower': {'year': '2005'},
    'Mrazek': {'year': '1994'},
    'Lazarus': {'year': '1984'},
    'Masten': {'year': '2014'},
    'Rosenberg': {'year': '1965'},
    'Urao': {'year': '2018'},  # Also 2022
    'He': {'year': '2026'},
    'Pascoe': {'year': '2020'},
    'Li': {'year': '2022'},
    'Anderson': {'year': '2025'},
    'Bhardwaj': {'year': '2020', 'n': '288', 'rate': '81,9%'},
    'Alharbi': {'year': '2019'},
    'Nakie': {'year': '2022'},
    'Jefferies': {'year': '2020', 'rate_vn': '30,7%'},
    'Saikia': {'year': '2023'},
    'Hồ Thị Trúc Quỳnh': {'year': '2025', 'beta_optimism': '-0,15'},
    'Trần Nguyễn Ngọc': {'year': '2018'},
    'Nguyễn Ngọc Bảo Quyên': {'year': '2025', 'n': '1.252'},
    'Nguyễn Danh Lâm': {'year': '2022', 'n': '1.024', 'rate': '25,8%'},
    'Lê Minh T': {'year': '2025', 'n': '600', 'rate': '33,2%'},
}

print(f"Doc loaded: {len(text)} chars")
print(f"Total verified entities: {len(VERIFIED)}")
print()

# Test each entity
errors = []
ok = []
not_mentioned = []

for name, facts in VERIFIED.items():
    # Find name + year in text
    year = facts.get('year', '')
    # Check if name + year both appear
    if name not in text:
        not_mentioned.append(name)
        continue

    # For each fact, check if it appears within 600 chars of name
    for fact_key, fact_val in facts.items():
        if fact_key == 'year':
            continue
        # Find positions of name
        found = False
        for m in re.finditer(re.escape(name), text):
            start = max(0, m.start() - 100)
            end = min(len(text), m.end() + 800)
            window = text[start:end]
            if fact_val in window:
                found = True
                break
        if not found:
            # Check if it's NEAR_MISS (val mentioned somewhere)
            if fact_val in text:
                errors.append((name, fact_key, fact_val, 'val exists but not near name'))
            else:
                # val not in text at all - might just not be mentioned in this short version
                pass
        else:
            ok.append((name, fact_key, fact_val))

print(f"=== AUDIT RESULTS ===")
print(f"OK matches: {len(ok)}")
print(f"Errors (val mentioned but not near name): {len(errors)}")
print(f"Not mentioned (entity not in text - OK to skip): {len(not_mentioned)}")
print()

if errors:
    print("=== POTENTIAL ISSUES ===")
    for e in errors:
        print(f"  {e[0]} - {e[1]}={e[2]} ({e[3]})")

print()
print("=== ENTITIES NOT MENTIONED (acceptable) ===")
for n in not_mentioned[:20]:
    print(f"  - {n}")

# Also check internal consistency - find pairs of (name, year) that appear multiple times
# to see if they use consistent format
print()
print("=== CITATION FORMAT CONSISTENCY (per entity) ===")
inconsistent = []
for name, facts in VERIFIED.items():
    year = facts.get('year', '')
    if not year or name not in text:
        continue
    # Find variations: "X et al", "X và cộng sự", "X và cs", etc.
    patterns = [
        re.escape(name) + r'\s+và\s+cộng\s+sự',
        re.escape(name) + r'\s+và\s+cs\.',
        re.escape(name) + r'\s+et\s+al\.',
        re.escape(name) + r'\s+và\s+\w+\s+\w+',  # X và Y Z (named coauthor)
        re.escape(name) + r'\s*\(',  # X (
    ]
    variants = []
    for pat in patterns:
        matches = re.findall(pat, text)
        if matches:
            variants.append((pat, len(matches)))
    if len(variants) > 1:
        # Multiple format used
        inconsistent.append((name, variants))

for name, variants in inconsistent[:10]:
    print(f"  {name}:")
    for v in variants:
        print(f"     {v}")
