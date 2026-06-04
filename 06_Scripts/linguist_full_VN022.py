# -*- coding: utf-8 -*-
"""
LINGUIST EXPERT — full check 542 remaining paragraphs (73%)
After fix expert review (3 errors), now check the rest comprehensively.
"""
import sys, io, re, random
import os
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

PATH = '03_Ban-dich/VN022_UNICEF_VN_2022_SchoolFactors_FULL.docx'
d = Document(PATH)
all_paras = [(i, p.text) for i, p in enumerate(d.paragraphs) if p.text.strip() and len(p.text) > 80]

# Re-derive same 200 already reviewed (50 + 150)
random.seed(2026)
first_50 = sorted(random.sample(range(len(all_paras)), 50))
unsampled_after_50 = [i for i in range(len(all_paras)) if i not in first_50]
random.seed(99)
next_150 = sorted(random.sample(unsampled_after_50, 150))
already = set(first_50) | set(next_150)
remaining_indices = [i for i in range(len(all_paras)) if i not in already]
remaining = [all_paras[i] for i in remaining_indices]

print(f'Total substantial: {len(all_paras)}')
print(f'Already reviewed: {len(already)}')
print(f'Remaining (this check): {len(remaining)}')
print('='*70)

remaining_text = '\n\n'.join(p[1] for p in remaining)
all_issues = []

# ============================================================
# CHECK 1: Anglicism — English words leaking outside parens
# ============================================================
print('\nCHECK 1: ENGLISH WORDS LEAKING (outside parens)')
print('-'*70)

# Words that should be translated (not allowed standalone)
suspect = [
    (r'\bteacher(s)?\b(?![\s)])', 'teacher'),
    (r'\bstudent(s)?\b(?![\s)])', 'student'),
    (r'\bschool(s)?\b(?![\s)])', 'school'),
    (r'\bteen(ager)?(s)?\b', 'teen'),
    (r'\b(boy|girl)(s)?\b(?![\s)])', 'boy/girl'),
    (r'\bfamily\b', 'family'),
    (r'\bstrong(ly)?\b', 'strong'),
    (r'\bhealth\b(?![\s)-])', 'health'),
    (r'\bmental\b(?![\s)-])', 'mental'),
    (r'\banxiety\b(?![\s)])', 'anxiety'),
    (r'\bdepression\b(?![\s)])', 'depression'),
    (r'\bbullying\b(?![\s)])', 'bullying'),
    (r'\bgender\b(?![\s)])', 'gender'),
    (r'\bage(s)?\b(?![\s)])', 'age'),
]

for pat, label in suspect:
    real_leaks = []
    for m in re.finditer(pat, remaining_text, re.IGNORECASE):
        # Check context
        before = remaining_text[max(0, m.start()-100):m.start()]
        # In parens?
        if '(' in before and before.rfind('(') > before.rfind(')'):
            continue
        # In reference?
        ctx = remaining_text[max(0, m.start()-200):m.end()+100]
        if any(skip in ctx for skip in ['http', 'doi.org', 'pp.', 'PLOS', 'Journal', 'Lancet',
                                        'Comprehensive Study', 'BMJ', 'BMC', 'al.', '&', 'Vol',
                                        'New York', 'Edition', 'Press', 'Routledge']):
            continue
        # Skip 'in' patterns from URLs
        real_leaks.append(m)
    if len(real_leaks) > 2:
        print(f'  {label}: {len(real_leaks)} potential leaks')
        for m in real_leaks[:3]:
            ctx = remaining_text[max(0, m.start()-60):m.end()+60].replace('\n', ' ')
            print(f'    ...{ctx[:180]}...')
        all_issues.append(f'{label}: {len(real_leaks)} leaks')

# ============================================================
# CHECK 2: Common mistranslation patterns (literal translations)
# ============================================================
print('\nCHECK 2: LITERAL/AWKWARD TRANSLATIONS')
print('-'*70)
literal = [
    (r'làm cho\s+\w+\s+(trở thành|trở nên)', 'make X become — awkward, prefer "khiến X trở nên"'),
    (r'tới mức (mà|là)', 'to the extent that — awkward'),
    (r'\bnhư là\b', 'as is — overuse'),
    (r'theo cách (mà|đó|nào)', 'in a way that — overuse'),
    (r'\bbên cạnh đó,\s+\w+', 'besides that — usually OK but can be overused'),
    (r'\bthực sự rằng\b', 'really that — Anglicism'),
    (r'\btrong cái mà\b', 'in which — Anglicism'),
    (r'nó là\s+\w+\s+(rằng|mà)', 'it is X that — Anglicism cleft'),
    (r'có nghĩa là rằng', '"means that" double — redundant'),
    (r'\btùy thuộc vào\b', 'depends on — usually OK'),
    (r'\bdo bởi\b', 'caused by — redundant'),
    (r'\bdựa vào trên\b', 'based on — redundant double'),
    (r'\bmỗi cá nhân\b', 'each individual — natural in formal'),
    (r'\bcó trải qua\b', 'have experienced — usually awkward, prefer "đã trải qua"'),
    (r'không có khả năng để', 'unable to — should be "không thể"'),
]

for pat, note in literal:
    matches = list(re.finditer(pat, remaining_text))
    if len(matches) > 0:
        print(f'  Pattern: {pat[:50]} ({len(matches)} hits) — {note}')
        for m in matches[:2]:
            ctx = remaining_text[max(0, m.start()-50):m.end()+50].replace('\n', ' ')
            print(f'    ...{ctx[:160]}...')
        if len(matches) > 3:
            all_issues.append(f'Literal {note}: {len(matches)}')

# ============================================================
# CHECK 3: Capitalization (Bộ/Sở/Cục consistency)
# ============================================================
print('\nCHECK 3: CAPITALIZATION OF GOV ENTITIES')
print('-'*70)
# After verbs like "rằng", "Một X từ Y", entities should be capitalized
issues_cap = []
for m in re.finditer(r'\brằng\s+(bộ|sở|cục|phòng)\s+\w+', remaining_text):
    ctx = remaining_text[max(0, m.start()-40):m.end()+40].replace('\n', ' ')
    issues_cap.append(ctx)
    print(f'  Lowercase entity: ...{ctx[:200]}...')

# Look for "the {Department/Ministry}" → "Bộ/Sở"
for m in re.finditer(r'\bthe\s+(Ministry|Department|Office)', remaining_text):
    ctx = remaining_text[max(0, m.start()-40):m.end()+60].replace('\n', ' ')
    print(f'  Untranslated: ...{ctx[:180]}...')
    issues_cap.append(ctx)

# ============================================================
# CHECK 4: Quote integrity (quotes from interviews)
# ============================================================
print('\nCHECK 4: QUOTE INTEGRITY (translated student/teacher quotes)')
print('-'*70)

# Find suspicious metaphors that may be mistranslated
suspect_metaphors = [
    'gà công nghiệp',  # industrial chicken — let me check
    'quy luật',
    'pháp luật',
    'kẻ lật tẩy',
    'mác học sinh',
    'lá ngón',
    'cỏ ngắt',
    'heartbreak grass',
]
for term in suspect_metaphors:
    for m in re.finditer(term, remaining_text):
        ctx = remaining_text[max(0, m.start()-100):m.end()+100].replace('\n', ' ')
        print(f'  [{term}]: ...{ctx[:250]}...')

# ============================================================
# CHECK 5: Number consistency in remaining
# ============================================================
print('\nCHECK 5: NUMBER FORMAT IN REMAINING')
print('-'*70)
period_dec = re.findall(r'\b\d+\.\d+\s*%', remaining_text)
print(f'  Period decimal: {len(period_dec)}')
for d_ in period_dec[:5]:
    idx = remaining_text.find(d_)
    if idx > 0:
        print(f'    Found: {d_} at ...{remaining_text[max(0,idx-30):idx+30]}...')

# ============================================================
# CHECK 6: Sentence flow problems
# ============================================================
print('\nCHECK 6: BROKEN SENTENCES / BAD SPLITS')
print('-'*70)
# Sentences starting with lowercase (after period)
bad_starts = re.findall(r'\.\s+([a-zà-ỹ])', remaining_text)
print(f'  Sentences starting lowercase after period: {len(bad_starts)}')

# Untranslated punctuation issues
weird_punct = []
for m in re.finditer(r'\s+\,', remaining_text):  # space before comma
    ctx = remaining_text[max(0, m.start()-20):m.end()+20]
    weird_punct.append((m.start(), ctx))
if weird_punct:
    print(f'  Space before comma: {len(weird_punct)} instances')

# ============================================================
# FINAL REPORT
# ============================================================
print('\n' + '='*70)
print(f'TOTAL ISSUES IDENTIFIED IN REMAINING {len(remaining)} PARAGRAPHS:')
print('='*70)
for issue in all_issues:
    print(f'  - {issue}')
print(f'\nTotal: {len(all_issues)} issue categories')
