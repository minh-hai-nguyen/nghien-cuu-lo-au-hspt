# -*- coding: utf-8 -*-
"""Verify that the 152 'lowercase after period' detections are false positives."""
import sys, io, re, random, os
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

PATH = '03_Ban-dich/VN022_UNICEF_VN_2022_SchoolFactors_FULL.docx'
d = Document(PATH)
all_paras = [(i, p.text) for i, p in enumerate(d.paragraphs) if p.text.strip() and len(p.text) > 80]

random.seed(2026)
first_50 = sorted(random.sample(range(len(all_paras)), 50))
unsampled_after_50 = [i for i in range(len(all_paras)) if i not in first_50]
random.seed(99)
next_150 = sorted(random.sample(unsampled_after_50, 150))
already = set(first_50) | set(next_150)
remaining_indices = [i for i in range(len(all_paras)) if i not in already]
remaining = [all_paras[i] for i in remaining_indices]
remaining_text = '\n\n'.join(p[1] for p in remaining)

# Categorize the 152 lowercase-starts
categories = {'numbered_list': 0, 'abbreviation': 0, 'reference_pp': 0,
              'reference_vol': 0, 'reference_etal': 0, 'doi_url': 0,
              'real_issue': 0, 'other': 0}
real_issues = []

for m in re.finditer(r'(\S{0,30})\.\s+([a-zà-ỹ])(\w{0,30})', remaining_text):
    before = m.group(1)
    letter = m.group(2)
    after = m.group(3)
    full_ctx = remaining_text[max(0, m.start()-30):m.end()+30]

    if before.strip().isdigit():  # "1. text"
        categories['numbered_list'] += 1
    elif re.search(r'\b(pp|p|vol|no|ed|al|et|eds|ch|fig|tab|sec|para|ibid)$', before, re.IGNORECASE):
        if 'al' in before.lower(): categories['reference_etal'] += 1
        elif 'vol' in before.lower() or 'no' in before.lower(): categories['reference_vol'] += 1
        elif 'pp' in before.lower() or before.lower().endswith('p'): categories['reference_pp'] += 1
        else: categories['abbreviation'] += 1
    elif 'http' in full_ctx or 'doi' in full_ctx.lower() or 'www.' in full_ctx:
        categories['doi_url'] += 1
    elif re.search(r'\b(ed|edn|inc|ltd|co|st|ave|mr|mrs|dr|prof|sr|jr|vs)$', before, re.IGNORECASE):
        categories['abbreviation'] += 1
    else:
        # Check if it's mid-sentence after abbreviation we missed
        categories['other'] += 1
        if len(real_issues) < 10:
            real_issues.append(full_ctx)

print('Category breakdown of 152 lowercase-after-period detections:')
for cat, count in categories.items():
    print(f'  {cat}: {count}')

print(f'\nTotal categorized: {sum(categories.values())}')
print(f'\nSamples of "other" (potentially real issues):')
for i, ctx in enumerate(real_issues, 1):
    print(f'  {i}. ...{ctx}...')
