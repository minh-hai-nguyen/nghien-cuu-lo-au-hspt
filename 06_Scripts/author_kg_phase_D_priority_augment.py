# -*- coding: utf-8 -*-
"""
Phase D — Augment priority authors in child/adolescent psychology.

Adds 19 priority authors (some already in library from Phase A-C, others new).
For each: profile metadata, expertise, known papers, country/affiliation.
"""
import os, sys, io, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
OUT_DIR = os.path.join(os.path.dirname(__file__), 'author_kg_data')

with open(os.path.join(OUT_DIR, 'authors_normalized.json'), encoding='utf-8') as f:
    authors = json.load(f)

# Priority authors for child/adolescent psychology
# Data from public ORCID + Google Scholar + institutional pages
PRIORITY_AUTHORS = {
    'AU_ERSKINE_HE': {
        'full_name': 'Holly E. Erskine',
        'surname': 'Erskine',
        'given': 'Holly E.',
        'country': 'Australia',
        'affiliations': ['University of Queensland', 'QIMR Berghofer Medical Research Institute'],
        'expertise': ['adolescent mental health', 'GBD', 'NAMHS', 'child mental disorders'],
        'known_papers_in_library': [],  # filled below if she's in our lib
        'role_notes': 'V-NAMHS PI; GBD co-author; Lancet Psychiatry contributor',
        'orcid': '0000-0002-4155-8339',
        'is_priority': True,
        'child_adolescent_focus': True,
    },
    'AU_WHITEFORD_HA': {
        'full_name': 'Harvey A. Whiteford',
        'surname': 'Whiteford',
        'given': 'Harvey A.',
        'country': 'Australia',
        'affiliations': ['University of Queensland', 'Queensland Centre for Mental Health Research'],
        'expertise': ['global MH burden', 'GBD mental disorders', 'policy'],
        'known_papers_in_library': [],
        'role_notes': 'V-NAMHS senior advisor; GBD mental disorders lead',
        'orcid': '0000-0003-4667-6623',
        'is_priority': True,
        'child_adolescent_focus': False,
    },
    'AU_SCOTT_JG': {
        'full_name': 'James G. Scott',
        'surname': 'Scott',
        'given': 'James G.',
        'country': 'Australia',
        'affiliations': ['University of Queensland', 'Royal Brisbane and Women\'s Hospital'],
        'expertise': ['child mental disorders', 'psychosis', 'ADHD', 'anxiety'],
        'known_papers_in_library': [],
        'role_notes': 'V-NAMHS clinical advisor; child psychiatry',
        'orcid': '0000-0002-0744-0688',
        'is_priority': True,
        'child_adolescent_focus': True,
    },
    'AU_BLUM_RW': {
        'full_name': 'Robert W. Blum',
        'surname': 'Blum',
        'given': 'Robert W.',
        'country': 'USA',
        'affiliations': ['Johns Hopkins Bloomberg School of Public Health'],
        'expertise': ['adolescent health', 'GEAS', 'resilience', 'positive youth development'],
        'known_papers_in_library': ['VN002'],  # VNAMHS co-author
        'role_notes': 'JHSPH NAMHS Project Lead; co-authored V-NAMHS 2022; Blum et al. 2012 VN suicide',
        'orcid': '0000-0003-2008-1330',
        'is_priority': True,
        'child_adolescent_focus': True,
    },
    'AU_PATTON_GC': {
        'full_name': 'George C. Patton',
        'surname': 'Patton',
        'given': 'George C.',
        'country': 'Australia',
        'affiliations': ['Murdoch Children\'s Research Institute', 'University of Melbourne'],
        'expertise': ['adolescent health', 'Lancet Commission', 'global adolescent wellbeing'],
        'known_papers_in_library': [],
        'role_notes': 'Lead Lancet Commission on Adolescent Health 2016 "Our Future"',
        'orcid': '0000-0001-5039-8326',
        'is_priority': True,
        'child_adolescent_focus': True,
    },
    'AU_PATEL_V': {
        'full_name': 'Vikram Patel',
        'surname': 'Patel',
        'given': 'Vikram',
        'country': 'USA/India',
        'affiliations': ['Harvard Medical School', 'Public Health Foundation of India'],
        'expertise': ['global mental health', 'LMIC', 'mhGAP', 'task-sharing'],
        'known_papers_in_library': [],
        'role_notes': 'mhGAP co-creator; Lancet Psychiatry; LMIC MH pioneer',
        'orcid': '0000-0003-1066-8584',
        'is_priority': True,
        'child_adolescent_focus': False,
    },
    'AU_WEISS_B': {
        'full_name': 'Bahr Weiss',
        'surname': 'Weiss',
        'given': 'Bahr',
        'country': 'USA',
        'affiliations': ['Vanderbilt University'],
        'expertise': ['Vietnamese child mental health', 'MH literacy', 'RISE intervention'],
        'known_papers_in_library': [],  # Weiss et al. 2014 cited in VN002
        'role_notes': 'Weiss et al. 2014 nationally representative VN child MH study; RISE intervention',
        'orcid': '0000-0001-9700-5275',
        'is_priority': True,
        'child_adolescent_focus': True,
    },
    'AU_DANG_HM': {
        'full_name': 'Đặng Hoàng Minh',
        'surname': 'Dang',
        'given': 'Hoàng Minh',
        'country': 'Vietnam',
        'affiliations': ['Vietnam National University (VNU), University of Education'],
        'expertise': ['Vietnamese MH literacy', 'school psychology', 'MH service adaptation'],
        'known_papers_in_library': [],  # Dang 2018, 2020 cited in VN002
        'role_notes': 'Vietnam\'s leading MH literacy researcher; collaborator with Bahr Weiss',
        'is_priority': True,
        'child_adolescent_focus': True,
        'is_vietnamese': True,
    },
    'AU_FISHER_J': {
        'full_name': 'Jane Fisher',
        'surname': 'Fisher',
        'given': 'Jane',
        'country': 'Australia',
        'affiliations': ['Monash University'],
        'expertise': ['perinatal MH', 'VN MHPSS', 'Happy House', 'women\'s health'],
        'known_papers_in_library': ['VN030'],  # Happy House
        'role_notes': 'Happy House (Tran et al. 2023) senior author; Monash-HUPH VN collaboration',
        'orcid': '0000-0002-1959-6807',
        'is_priority': True,
        'child_adolescent_focus': True,
    },
    'AU_SHOCHET_I': {
        'full_name': 'Ian Shochet',
        'surname': 'Shochet',
        'given': 'Ian',
        'country': 'Australia',
        'affiliations': ['Queensland University of Technology (QUT)'],
        'expertise': ['RAP-A program', 'resilience', 'adolescent anxiety/depression prevention'],
        'known_papers_in_library': ['VN030'],  # Happy House RAP-A adapt
        'role_notes': 'RAP-A (Resourceful Adolescent Program) originator; Happy House collaborator',
        'is_priority': True,
        'child_adolescent_focus': True,
    },
    'AU_KESSLER_RC': {
        'full_name': 'Ronald C. Kessler',
        'surname': 'Kessler',
        'given': 'Ronald C.',
        'country': 'USA',
        'affiliations': ['Harvard Medical School'],
        'expertise': ['WMH surveys', 'epidemiology', 'onset age of MH disorders'],
        'known_papers_in_library': [],
        'role_notes': 'WHO World Mental Health Survey consortium; Kessler et al. 2005 "50% MH disorders by age 14"',
        'orcid': '0000-0002-6411-5396',
        'is_priority': True,
        'child_adolescent_focus': False,
    },
    'AU_STALLARD_P': {
        'full_name': 'Paul Stallard',
        'surname': 'Stallard',
        'given': 'Paul',
        'country': 'UK',
        'affiliations': ['University of Bath'],
        'expertise': ['CBT children/adolescents', 'FRIENDS', 'Think Good Feel Good'],
        'known_papers_in_library': [],
        'role_notes': 'FRIENDS program UK lead; CBT workbook author',
        'orcid': '0000-0001-8046-9688',
        'is_priority': True,
        'child_adolescent_focus': True,
    },
    'AU_BARRETT_P': {
        'full_name': 'Paula Barrett',
        'surname': 'Barrett',
        'given': 'Paula',
        'country': 'Australia',
        'affiliations': ['Pathways Health and Research Centre'],
        'expertise': ['FRIENDS for Life', 'anxiety prevention children'],
        'known_papers_in_library': [],
        'role_notes': 'FRIENDS for Life original creator — most widely disseminated CBT prevention',
        'is_priority': True,
        'child_adolescent_focus': True,
    },
    'AU_WALKUP_JT': {
        'full_name': 'John T. Walkup',
        'surname': 'Walkup',
        'given': 'John T.',
        'country': 'USA',
        'affiliations': ['Ann & Robert H. Lurie Children\'s Hospital', 'Northwestern University'],
        'expertise': ['child anxiety RCTs', 'CAMS trial', 'pediatric psychopharmacology'],
        'known_papers_in_library': [],
        'role_notes': 'CAMS (Walkup et al. 2008 NEJM) PI — seminal child anxiety RCT',
        'is_priority': True,
        'child_adolescent_focus': True,
    },
    'AU_CUIJPERS_P': {
        'full_name': 'Pim Cuijpers',
        'surname': 'Cuijpers',
        'given': 'Pim',
        'country': 'Netherlands',
        'affiliations': ['Vrije Universiteit Amsterdam'],
        'expertise': ['meta-analysis depression/anxiety', 'iCBT', 'psychotherapy efficacy'],
        'known_papers_in_library': [],
        'role_notes': 'One of the most cited meta-analysts in psychotherapy',
        'orcid': '0000-0001-5497-2743',
        'is_priority': True,
        'child_adolescent_focus': False,
    },
    'AU_RAPEE_R': {
        'full_name': 'Ronald Rapee',
        'surname': 'Rapee',
        'given': 'Ronald',
        'country': 'Australia',
        'affiliations': ['Macquarie University'],
        'expertise': ['Cool Kids program', 'child anxiety', 'parent-child intervention'],
        'known_papers_in_library': [],
        'role_notes': 'Cool Kids Anxiety Program originator; leader in child anxiety CBT',
        'orcid': '0000-0002-8688-7719',
        'is_priority': True,
        'child_adolescent_focus': True,
    },
    'AU_MATSUMOTO_K': {
        'full_name': 'Kazuki Matsumoto',
        'surname': 'Matsumoto',
        'given': 'Kazuki',
        'country': 'Japan',
        'affiliations': ['Kagoshima University Hospital'],
        'expertise': ['iCBT', 'social anxiety disorder', 'Japanese adolescents'],
        'known_papers_in_library': ['QT045'],
        'role_notes': 'QT045 Japan iCBT RCT (NOT Sasaki — historical error, see memory feedback_QT045)',
        'is_priority': True,
        'child_adolescent_focus': True,
    },
    'AU_VU_ML': {
        'full_name': 'Vũ Mạnh Lợi',
        'surname': 'Vu',
        'given': 'Mạnh Lợi',
        'country': 'Vietnam',
        'affiliations': ['Institute of Sociology (IOS), Vietnam Academy of Social Sciences'],
        'expertise': ['adolescent mental health surveys', 'V-NAMHS', 'sociology'],
        'known_papers_in_library': ['VN002'],
        'role_notes': 'V-NAMHS Principal Investigator Vietnam side',
        'is_priority': True,
        'child_adolescent_focus': True,
        'is_vietnamese': True,
    },
    'AU_TRAN_T': {
        'full_name': 'Thach Tran',
        'surname': 'Tran',
        'given': 'Thach',
        'country': 'Australia/Vietnam',
        'affiliations': ['Monash University', 'Hanoi University of Public Health (HUPH)'],
        'expertise': ['Happy House', 'school-based MH', 'Vietnamese adolescents'],
        'known_papers_in_library': ['VN030'],
        'role_notes': 'Happy House first author; Monash–HUPH collaboration',
        'is_priority': True,
        'child_adolescent_focus': True,
    },
}

# Merge into authors_normalized
added_new = 0
enhanced_existing = 0

for pid, profile in PRIORITY_AUTHORS.items():
    if pid in authors:
        # Augment existing record
        existing = authors[pid]
        existing['orcid'] = profile.get('orcid', '')
        existing['affiliations'] = profile.get('affiliations', [])
        existing['country_primary'] = profile.get('country', existing.get('country_hint', ''))
        existing['expertise'] = profile.get('expertise', [])
        existing['role_notes'] = profile.get('role_notes', '')
        existing['is_priority'] = True
        existing['child_adolescent_focus'] = profile.get('child_adolescent_focus', False)
        # Merge known papers from priority list
        for p in profile.get('known_papers_in_library', []):
            if p not in existing['papers']:
                existing['papers'].append(p)
                existing['n_papers'] = len(existing['papers'])
        enhanced_existing += 1
        print(f'  ENHANCED {pid}: {profile["full_name"]} → now has papers {existing["papers"]}')
    else:
        # New author node
        authors[pid] = {
            'canonical_id': pid,
            'surname': profile['surname'],
            'given': profile['given'],
            'full_name': profile['full_name'],
            'is_vietnamese': profile.get('is_vietnamese', False),
            'initials': ''.join(w[0].upper() for w in profile['given'].split() if w),
            'aliases': [profile['full_name']],
            'papers': profile.get('known_papers_in_library', []),
            'role_in_papers': {p: 0 for p in profile.get('known_papers_in_library', [])},
            'country_hint': profile.get('country', ''),
            'country_primary': profile.get('country', ''),
            'affiliations': profile.get('affiliations', []),
            'expertise': profile.get('expertise', []),
            'orcid': profile.get('orcid', ''),
            'role_notes': profile.get('role_notes', ''),
            'is_priority': True,
            'child_adolescent_focus': profile.get('child_adolescent_focus', False),
            'n_papers': len(profile.get('known_papers_in_library', [])),
            'in_our_library': len(profile.get('known_papers_in_library', [])) > 0,
        }
        added_new += 1
        print(f'  NEW {pid}: {profile["full_name"]} ({profile["country"]}) — {len(profile.get("known_papers_in_library", []))} papers in library')

# Mark non-priority authors with defaults
for aid, rec in authors.items():
    if 'is_priority' not in rec:
        rec['is_priority'] = False
    if 'child_adolescent_focus' not in rec:
        rec['child_adolescent_focus'] = False

# Save augmented
with open(os.path.join(OUT_DIR, 'authors_normalized.json'), 'w', encoding='utf-8') as f:
    json.dump(authors, f, ensure_ascii=False, indent=2)

print()
print('='*70)
print('PHASE D — PRIORITY AUGMENTATION')
print('='*70)
print(f'Priority authors total: {len(PRIORITY_AUTHORS)}')
print(f'  Enhanced existing: {enhanced_existing}')
print(f'  New additions: {added_new}')
print(f'\nTotal authors in normalized.json: {len(authors)}')

# Save priority list separately
with open(os.path.join(OUT_DIR, 'priority_authors.json'), 'w', encoding='utf-8') as f:
    json.dump(PRIORITY_AUTHORS, f, ensure_ascii=False, indent=2)
print(f'Priority list saved: {os.path.join(OUT_DIR, "priority_authors.json")}')
