# -*- coding: utf-8 -*-
"""Update author_kg_v1 -> v2: add Karasu, Rose, Stankov, Small, Blanc, First,
Pezzella + their papers + edges. Source: verified PDF Karasu bio sketch +
WebFetch PubMed/PMC/ERIC for other 4 authors."""
import sys, io, json, os

try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
INFILE = os.path.join(ROOT, '06_Scripts', 'author_kg_data', 'author_kg_v1.json')
OUTFILE = os.path.join(ROOT, '06_Scripts', 'author_kg_data', 'author_kg_v2.json')

with open(INFILE, 'r', encoding='utf-8') as f:
    kg = json.load(f)

# === NEW AUTHOR NODES ===
new_authors = [
    {
        'id': 'AU_KARASU_TB',
        'type': 'Author', 'label': 'T. Byram Karasu',
        'surname': 'Karasu', 'given': 'T. Byram',
        'is_vietnamese': 'False',
        'n_papers': 2, 'country_hint': 'USA',
        'affiliation': 'Albert Einstein College of Medicine',
        'titles': 'Silverman Professor and Chairman, Dept of Psychiatry & '
                  'Behavioral Sciences; Distinguished Life Fellow APA; '
                  'Editor-in-Chief emeritus AJ Psychotherapy',
        'birth_year': '1935',
        'verification': 'PDF bio sketch + Encyclopedia.com (07/06/2026)',
    },
    {
        'id': 'AU_ROSE_AJ',
        'type': 'Author', 'label': 'Amanda J. Rose',
        'surname': 'Rose', 'given': 'Amanda J.',
        'is_vietnamese': 'False',
        'n_papers': 1, 'country_hint': 'USA',
        'concept': 'Co-rumination hypothesis',
        'verification': 'PubMed PMID 12487497',
    },
    {
        'id': 'AU_STANKOV_L',
        'type': 'Author', 'label': 'Lazar Stankov',
        'surname': 'Stankov', 'given': 'Lazar',
        'is_vietnamese': 'False',
        'n_papers': 1, 'country_hint': 'Australia',
        'concept': 'Confucian academic culture',
        'verification': 'ERIC EJ905763 + ScienceDirect',
    },
    {
        'id': 'AU_SMALL_S',
        'type': 'Author', 'label': 'Stuart Small',
        'surname': 'Small', 'given': 'Stuart',
        'is_vietnamese': 'False',
        'n_papers': 1, 'country_hint': 'USA',
        'verification': 'PMC PMC7820702',
    },
    {
        'id': 'AU_BLANC_J',
        'type': 'Author', 'label': 'Judite Blanc',
        'surname': 'Blanc', 'given': 'Judite',
        'is_vietnamese': 'False',
        'n_papers': 1, 'country_hint': 'USA',
        'verification': 'PMC PMC7820702',
    },
    {
        'id': 'AU_FIRST_MB',
        'type': 'Author', 'label': 'Michael B. First',
        'surname': 'First', 'given': 'Michael B.',
        'is_vietnamese': 'False',
        'n_papers': 1, 'country_hint': 'USA',
        'verification': 'PMC PMC9077590 (DSM-5-TR overview)',
    },
    {
        'id': 'AU_PEZZELLA_P',
        'type': 'Author', 'label': 'Pasquale Pezzella',
        'surname': 'Pezzella', 'given': 'Pasquale',
        'is_vietnamese': 'False',
        'n_papers': 1, 'country_hint': 'Italy',
        'verification': 'PMC PMC9077598 (ICD-11 effective announcement)',
    },
]

# === NEW PAPER NODES ===
new_papers = [
    {
        'id': 'PA_KARASU_1986_AJPSY',
        'type': 'Paper',
        'label': 'Karasu 1986 — The psychotherapies: benefits and limitations',
        'authors': ['AU_KARASU_TB'],
        'journal': 'American Journal of Psychotherapy',
        'volume': '40', 'issue': '3', 'pages': '324-342', 'year': 1986,
        'pmid': '3094389',
        'doi': '10.1176/appi.psychotherapy.1986.40.3.324',
        'key_finding': '"over 450 types polled nationwide" (page 325) — '
                       'CONFIRMED nguyên văn từ PDF',
        'pdf_path': '02_Papers-goc/karasu1986.pdf',
        'verification_status': 'PRIMARY_VERIFIED',
    },
    {
        'id': 'PA_KARASU_1986_AJP',
        'type': 'Paper',
        'label': 'Karasu 1986 — The specificity vs nonspecificity dilemma',
        'authors': ['AU_KARASU_TB'],
        'journal': 'American Journal of Psychiatry',
        'volume': '143', 'issue': '6', 'pages': '687-695', 'year': 1986,
        'pmid': '3717390',
        'note': 'Bài KHÁC cùng năm 1986 — không phải nguồn của 450+',
        'verification_status': 'PUBMED_VERIFIED',
    },
    {
        'id': 'PA_ROSE_2002',
        'type': 'Paper',
        'label': 'Rose 2002 — Co-rumination in the friendships of girls and boys',
        'authors': ['AU_ROSE_AJ'],
        'journal': 'Child Development',
        'volume': '73', 'issue': '6', 'pages': '1830-1843', 'year': 2002,
        'pmid': '12487497',
        'doi': '10.1111/1467-8624.00509',
        'key_concept': 'Co-rumination hypothesis',
        'verification_status': 'PUBMED_VERIFIED',
    },
    {
        'id': 'PA_STANKOV_2010',
        'type': 'Paper',
        'label': 'Stankov 2010 — Unforgiving Confucian culture',
        'authors': ['AU_STANKOV_L'],
        'journal': 'Learning and Individual Differences',
        'volume': '20', 'issue': '6', 'pages': '555-563', 'year': 2010,
        'doi': '10.1016/j.lindif.2010.05.003',
        'key_concept': 'Confucian academic culture, test anxiety, self-doubt',
        'verification_status': 'ERIC_VERIFIED',
    },
    {
        'id': 'PA_SMALL_BLANC_2021',
        'type': 'Paper',
        'label': 'Small & Blanc 2021 — Tam Giao Vietnam',
        'authors': ['AU_SMALL_S', 'AU_BLANC_J'],
        'journal': 'Frontiers in Psychiatry',
        'volume': '11', 'pages': '589618', 'year': 2021,
        'doi': '10.3389/fpsyt.2020.589618',
        'key_concept': 'Tam giao (Confucianism + Buddhism + Taoism); '
                       'cultural additivity',
        'verification_status': 'PMC_FULLTEXT_VERIFIED',
    },
    {
        'id': 'PA_FIRST_2022',
        'type': 'Paper',
        'label': 'First et al 2022 — DSM-5-TR overview',
        'authors': ['AU_FIRST_MB'],
        'journal': 'World Psychiatry',
        'volume': '21', 'issue': '2', 'pages': '218-219', 'year': 2022,
        'doi': '10.1002/wps.20989',
        'key_finding': 'DSM-5 + 3 new diagnoses + 70 modified criteria sets',
        'verification_status': 'PMC_VERIFIED',
    },
    {
        'id': 'PA_PEZZELLA_2022',
        'type': 'Paper',
        'label': 'Pezzella 2022 — ICD-11 now officially in effect',
        'authors': ['AU_PEZZELLA_P'],
        'journal': 'World Psychiatry',
        'volume': '21', 'issue': '2', 'pages': '331-332', 'year': 2022,
        'doi': '10.1002/wps.20982',
        'key_finding': 'ICD-11 effective globally 01/01/2022',
        'verification_status': 'PMC_VERIFIED',
    },
]

# === NEW TOPIC NODES ===
new_topics = [
    {
        'id': 'TP_PSYCHOTHERAPY_SCHOOLS',
        'type': 'Topic',
        'label': 'Number of psychotherapy schools',
        'description': 'Karasu 1986 reported "over 450 types"; common '
                       'misattribution as "400+"',
    },
    {
        'id': 'TP_CONFUCIAN_CULTURE',
        'type': 'Topic',
        'label': 'Confucian academic culture',
        'description': 'Stankov 2010: high achievement + test anxiety + '
                       'self-doubt nexus',
    },
    {
        'id': 'TP_CORUMINATION',
        'type': 'Topic',
        'label': 'Co-rumination hypothesis',
        'description': 'Rose 2002: peer disclosure amplifying worry vs '
                       'dissipating it',
    },
    {
        'id': 'TP_TAM_GIAO',
        'type': 'Topic',
        'label': 'Tam Giao (Vietnamese cultural framework)',
        'description': 'Small & Blanc 2021: Confucianism+Buddhism+Taoism '
                       'cultural additivity in Vietnam',
    },
    {
        'id': 'TP_DSM_ICD_CLASSIFICATION',
        'type': 'Topic',
        'label': 'DSM/ICD classification systems',
        'description': 'DSM-5 (237), DSM-5-TR (240+), ICD-11 Chapter 06 '
                       '(162 four-character categories, 18 groupings, '
                       'effective 01/01/2022)',
    },
]

# === NEW EDGES ===
new_edges = []
# AUTHORED_BY edges
paper_author_pairs = [
    ('PA_KARASU_1986_AJPSY', 'AU_KARASU_TB'),
    ('PA_KARASU_1986_AJP', 'AU_KARASU_TB'),
    ('PA_ROSE_2002', 'AU_ROSE_AJ'),
    ('PA_STANKOV_2010', 'AU_STANKOV_L'),
    ('PA_SMALL_BLANC_2021', 'AU_SMALL_S'),
    ('PA_SMALL_BLANC_2021', 'AU_BLANC_J'),
    ('PA_FIRST_2022', 'AU_FIRST_MB'),
    ('PA_PEZZELLA_2022', 'AU_PEZZELLA_P'),
]
for p, a in paper_author_pairs:
    new_edges.append({'type': 'AUTHORED_BY', 'source': p, 'target': a})

# CO_AUTHORED Small <-> Blanc
new_edges.append({'type': 'CO_AUTHORED', 'source': 'AU_SMALL_S',
                  'target': 'AU_BLANC_J',
                  'paper_id': 'PA_SMALL_BLANC_2021'})

# STUDIES (Paper -> Topic)
paper_topic = [
    ('PA_KARASU_1986_AJPSY', 'TP_PSYCHOTHERAPY_SCHOOLS'),
    ('PA_STANKOV_2010', 'TP_CONFUCIAN_CULTURE'),
    ('PA_ROSE_2002', 'TP_CORUMINATION'),
    ('PA_SMALL_BLANC_2021', 'TP_TAM_GIAO'),
    ('PA_FIRST_2022', 'TP_DSM_ICD_CLASSIFICATION'),
    ('PA_PEZZELLA_2022', 'TP_DSM_ICD_CLASSIFICATION'),
]
for p, t in paper_topic:
    new_edges.append({'type': 'STUDIES', 'source': p, 'target': t})

# Append to KG
kg['nodes'].extend(new_authors)
kg['nodes'].extend(new_papers)
kg['nodes'].extend(new_topics)
kg['edges'].extend(new_edges)

# Update meta
kg['meta'] = {
    'created': '2026-04-15',
    'last_updated': '2026-06-07',
    'n_nodes': len(kg['nodes']),
    'n_edges': len(kg['edges']),
    'node_types': {
        'Author': len([n for n in kg['nodes'] if n.get('type') == 'Author']),
        'Paper': len([n for n in kg['nodes'] if n.get('type') == 'Paper']),
        'Country': len([n for n in kg['nodes'] if n.get('type') == 'Country']),
        'Topic': len([n for n in kg['nodes'] if n.get('type') == 'Topic']),
    },
    'v2_changelog': [
        '2026-06-07: Added 7 new Authors (Karasu, Rose, Stankov, Small, '
        'Blanc, First, Pezzella)',
        '2026-06-07: Added 7 new Papers (Karasu 1986 AJPsy + AJP, Rose 2002, '
        'Stankov 2010, Small & Blanc 2021, First 2022, Pezzella 2022)',
        '2026-06-07: Added 5 new Topics (Psychotherapy schools, Confucian '
        'culture, Co-rumination, Tam Giao, DSM/ICD)',
        '2026-06-07: Karasu 1986 number CORRECTED to 450+ (not 400+) — '
        'verified directly from PDF page 325',
    ],
}

with open(OUTFILE, 'w', encoding='utf-8') as f:
    json.dump(kg, f, ensure_ascii=False, indent=2)

print(f'Saved: {OUTFILE}')
print(f'Total nodes: {len(kg["nodes"])} (was 241)')
print(f'Total edges: {len(kg["edges"])} (was 809)')
print(f'New nodes added: {len(new_authors) + len(new_papers) + len(new_topics)}')
print(f'New edges added: {len(new_edges)}')
