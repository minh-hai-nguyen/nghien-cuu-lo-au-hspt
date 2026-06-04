# -*- coding: utf-8 -*-
"""Meta-review: verify each citation/claim in new critique vs refs list + project knowledge."""
import os, sys, io, re
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
PATH = os.path.join(ROOT, '03_Ban-dich', 'VN022_UNICEF_VN_2022_SchoolFactors_FULL.docx')

d = Document(PATH)

# Gather ALL references text (paragraphs 867..end of refs section)
refs_text = []
in_refs = False
for i, p in enumerate(d.paragraphs):
    t = p.text.strip()
    if 'TÀI LIỆU THAM KHẢO' in t:
        in_refs = True
        continue
    if in_refs:
        # Stop when hitting critique
        if t.startswith('PHẢN BIỆN VÀ ĐÁNH GIÁ'):
            break
        refs_text.append(t)
refs_all = ' '.join(refs_text)
print(f'Total refs text length: {len(refs_all):,}')

# Claims in critique that cite specific authors
citations_to_check = [
    ('Patton et al., 2016', ['Patton', '2016']),
    ('Erskine et al., 2021', ['Erskine', '2021']),
    ('Hirota, Guerrero & Skokauskas (2020)', ['Hirota', 'Guerrero', 'Skokauskas', '2020']),
    ('Creswell & Plano Clark (2017)', ['Creswell', 'Plano Clark']),
    ('Goodman 1997', ['Goodman', '1997']),
    ('Dang et al. 2016', ['Dang', '2016']),
    ('Bradshaw et al. 2014', ['Bradshaw', '2014']),
    ('Sun et al. 2011', ['Sun', '2011']),
    ('Cong, Ngoc, Weiss, Luot & Dat 2018', ['Cong', 'Weiss', '2018']),
    ('Fazel, Hoagwood, Stephan & Ford (2014)', ['Fazel', 'Hoagwood', '2014']),
    ('Rutter 2007', ['Rutter', '2007']),
    ('Dupéré et al. 2015', ['Dupéré', '2015']),
    ('Campbell, Bann & Patalay (2021)', ['Campbell', 'Bann', 'Patalay', '2021']),
    ('Dang et al. 2017', ['Dang', '2017']),
    ('Meltzer et al. 2000', ['Meltzer', '2000']),
    ('Dalton, Rapa & Stein, 2020', ['Dalton', 'Rapa', 'Stein', '2020']),
    ('Colvin, Egan & Coulter (2019)', ['Colvin', 'Egan', 'Coulter', '2019']),
    ('Tong, Sainsbury & Craig 2007 (COREQ)', ['Tong', 'Sainsbury', '2007']),
    ('Das et al. 2016', ['Das', '2016']),
    ('Jayanthi, Thirunavukarasu & Rajkumar (2015)', ['Jayanthi', '2015']),
    ('Blum, Sudhinaraset & Emerson 2012', ['Blum', 'Sudhinaraset', 'Emerson', '2012']),
]

print('\nCitation verification against references list:')
print('='*70)
for label, tokens in citations_to_check:
    matches = [t for t in tokens if t in refs_all]
    status = 'IN REFS' if len(matches) >= max(2, len(tokens)-1) else 'MAYBE NOT IN REFS'
    print(f'  [{status}] {label}: matched {matches}')
