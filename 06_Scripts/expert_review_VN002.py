# -*- coding: utf-8 -*-
"""
Expert review VN002 bằng 4 góc nhìn: tâm lý học, nhà nghiên cứu, giáo viên, chuyên gia ngôn ngữ.
Identify specific issues → list fixes.
"""
import os, sys, io, re, random
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
TRANS_PATH = os.path.join(ROOT, '03_Ban-dich', 'VN002_VNAMHS_2022_National_FULL.docx')
SUMM_PATH = os.path.join(ROOT, 'Tom-tat-tung-bai', 'VN002_VNAMHS_2022_National.docx')

d = Document(TRANS_PATH)
all_paras = [(i, p.text) for i, p in enumerate(d.paragraphs) if p.text.strip() and len(p.text) > 80]
total = len(all_paras)
print(f'Translation paragraphs with >80 chars: {total}')

# Sample 40 random paragraphs
random.seed(2026)
sampled = random.sample(range(total), min(40, total))
sampled.sort()

# Also include first 5 and last 10 (critique tail)
sampled = list(set(sampled + list(range(5)) + list(range(total-10, total))))
sampled.sort()

# ===== LENS 1: PSYCHOLOGY SCHOLAR =====
# Check for: category errors (lo âu vs trầm cảm), DISC vs DSM, instrument misalignment
psych_issues = []
full_text = '\n'.join(t for _, t in all_paras)

# Category errors
cat_checks = [
    ('lo âu và trầm cảm gộp chung', 'lo âu / trầm cảm', False),  # legitimate slash
    ('GAD-7 + PHQ-9 gộp', 'kết hợp GAD-7 với PHQ-9', True),
    ('DISC-5 là screening', 'DISC-5 sàng lọc', True),
]
for label, pattern, is_error in cat_checks:
    if pattern in full_text:
        if is_error:
            psych_issues.append(f'Category error: "{pattern}" at {label}')

# Check instrument alignment
# DISC-5 should be paired with DSM-5, not "sàng lọc"
if 'DISC-5 sàng lọc' in full_text or 'DISC-5 là sàng lọc' in full_text:
    psych_issues.append('DISC-5 incorrectly labeled as "sàng lọc" (should be "chẩn đoán")')

# Check "mental health" translation consistency
trans_mh = full_text.count('sức khoẻ tâm thần') + full_text.count('SKTT')
trans_mh_alt = full_text.count('sức khỏe tâm thần')  # alternative spelling
if trans_mh_alt > 2 and trans_mh > 5:
    psych_issues.append(f'Inconsistent spelling: "sức khoẻ" ({trans_mh}) vs "sức khỏe" ({trans_mh_alt})')

# Check impairment framing
if 'lĩnh vực suy giảm' in full_text and 'impairment domain' in full_text:
    print('  ✓ Impairment domains properly paired')

# Prevalence vs proportion usage
if 'prevalence' in full_text.lower() and 'tỷ lệ' in full_text:
    print('  ✓ prevalence / tỷ lệ consistent')

# ===== LENS 2: RESEARCHER =====
# Methodology: weighted, statistical significance, CI, replicability
res_issues = []

# Check "đại diện quốc gia" claim linked to weighting
if 'đại diện quốc gia' in full_text:
    # Must appear near "weighted" or "gia trọng"
    idx = full_text.find('đại diện quốc gia')
    nearby = full_text[max(0, idx-300):idx+300]
    if 'weighted' not in nearby.lower() and 'gia trọng' not in nearby:
        res_issues.append('"đại diện quốc gia" claim needs weighting mention nearby')

# Check that significance tests statement is preserved
if 'tests of statistical significance are not included' in full_text or 'không công bố' in full_text:
    print('  ✓ Statistical significance caveat preserved')

# Response rate + non-response analysis
if '81,1' in full_text:
    idx = full_text.find('81,1')
    nearby = full_text[max(0, idx-300):idx+300]
    if 'non-response' in nearby.lower() or 'từ chối' in nearby:
        print('  ✓ Response rate analyzed')

# Check replicability info
if 'wave 2' in full_text.lower() or 'lặp lại' in full_text:
    print('  ✓ Wave 2 / repeat study mentioned')

# ===== LENS 3: TEACHER / PEDAGOGY =====
# Audience accessibility, example clarity, jargon management
teach_issues = []

# Check jargon is defined
jargon_terms = ['subthreshold', 'full threshold', 'weighting', 'DISC-5', 'DSM-5', 'mhGAP']
for term in jargon_terms:
    if term in full_text:
        # Check if definition is nearby
        idx = full_text.find(term)
        nearby = full_text[max(0, idx-200):idx+500]
        has_def = any(k in nearby for k in ['dưới ngưỡng', 'đầy đủ ngưỡng', 'gia trọng', 'chẩn đoán', 'American Psychiatric', 'mental health gap'])
        if term == 'DISC-5' and ('Diagnostic Interview' in nearby or 'đo lường' in nearby):
            continue  # defined
        if term == 'DSM-5' and ('Diagnostic and Statistical' in nearby or 'Hiệp hội Tâm thần' in nearby):
            continue
        if not has_def:
            teach_issues.append(f'Jargon "{term}" needs definition on first use')

# Check reader-friendly summary
if 'TÓM TẮT' in full_text:
    print('  ✓ Executive summary present')

# ===== LENS 4: LINGUIST =====
# Anglicism, literal translation, sentence flow
ling_issues = []

# Anglicism check
english_leaks = [
    (r'\bfull threshold\b(?![\s)])', 'full threshold (needs VN translation)'),
    (r'\bmental health\b(?![\s)])', 'mental health (should be SKTT)'),
    (r'\bchildren\b(?![\s)])', 'children'),
    (r'\bsurvey\b(?![\s)])', 'survey'),
    (r'\brelated\b(?![\s)])', 'related'),
]
for pat, label in english_leaks:
    matches = list(re.finditer(pat, full_text, re.IGNORECASE))
    # Filter out legitimate (in parens or in quotes)
    real = []
    for m in matches:
        before = full_text[max(0, m.start()-100):m.start()]
        in_paren = '(' in before and before.rfind('(') > before.rfind(')')
        if not in_paren:
            real.append(m)
    if len(real) > 3:
        ling_issues.append(f'Anglicism "{label}": {len(real)} real leaks')

# Literal translations
literal_patterns = [
    (r'làm cho\s+\w+\s+trở nên', 'make X become — prefer "khiến X trở nên"'),
    (r'\btheo cách mà\b', 'in such a way — overuse'),
    (r'\bnhư là\b', 'as is — overuse'),
    (r'\btrong khi đó\b', 'meanwhile — OK but overused'),
]
for pat, note in literal_patterns:
    hits = len(re.findall(pat, full_text))
    if hits > 3:
        ling_issues.append(f'Literal pattern {note}: {hits} hits')

# Quote marks consistency
curly_quotes = full_text.count('"') + full_text.count('"')
straight_quotes = full_text.count('"')
if straight_quotes > 30 and curly_quotes > 30:
    ling_issues.append(f'Mixed quote marks: curly {curly_quotes}, straight {straight_quotes}')

# Number format VN (comma decimal)
period_dec = re.findall(r'\b\d+\.\d+\s*%', full_text)
if len(period_dec) > 3:
    ling_issues.append(f'Period decimals (should be comma): {len(period_dec)}')

# Vietnamese diacritics consistency
viet_signs = {
    'sức khoẻ': full_text.count('sức khoẻ'),
    'sức khỏe': full_text.count('sức khỏe'),
    'hoà': full_text.count('hoà'),
    'hòa': full_text.count('hòa'),
}
if viet_signs['sức khoẻ'] > 0 and viet_signs['sức khỏe'] > 2:
    ling_issues.append(f'Tone-mark variants: "sức khoẻ" {viet_signs["sức khoẻ"]} vs "sức khỏe" {viet_signs["sức khỏe"]}')

# ===== REPORT =====
print('\n' + '='*70)
print('EXPERT REVIEW — 4 GÓC NHÌN')
print('='*70)

for lens_name, issues in [
    ('Psychology scholar', psych_issues),
    ('Researcher', res_issues),
    ('Teacher / Pedagogy', teach_issues),
    ('Linguist', ling_issues),
]:
    print(f'\n{lens_name}: {len(issues)} issue(s)')
    for i in issues:
        print(f'  ⚠ {i}')

# Save
import json
OUT = os.path.join(os.path.dirname(__file__), 'expert_VN002_report.json')
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump({
        'psychology': psych_issues,
        'researcher': res_issues,
        'teacher': teach_issues,
        'linguist': ling_issues,
    }, f, ensure_ascii=False, indent=2)
print(f'\nReport saved: {OUT}')
