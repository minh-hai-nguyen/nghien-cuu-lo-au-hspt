# -*- coding: utf-8 -*-
"""
Phase C — Build Author Knowledge Graph v1.

Input: authors_normalized.json, paper_author_edges.json, canonical_index.json
Output: author_kg_v1.graphml + author_kg_v1.json (nodes + edges)
"""
import os, sys, io, json
from collections import defaultdict
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    import networkx as nx
except ImportError:
    print('ERROR: networkx required. pip install networkx')
    sys.exit(1)

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
OUT_DIR = os.path.join(os.path.dirname(__file__), 'author_kg_data')

with open(os.path.join(OUT_DIR, 'authors_normalized.json'), encoding='utf-8') as f:
    authors = json.load(f)
with open(os.path.join(OUT_DIR, 'paper_author_edges.json'), encoding='utf-8') as f:
    paper_author_edges = json.load(f)
with open(os.path.join(ROOT, '02_Papers-goc', 'canonical_index.json'), encoding='utf-8') as f:
    canon = json.load(f)

# Filter out noise authors (from bad PDF extraction)
def is_valid_author(aid, rec):
    # Reject if surname is abbreviation only
    if rec['surname'].lower() in ('phd', 'md', 'ma', 'ba', 'msc', 'bsc'):
        return False
    # Reject if surname too long (> 20 chars — likely garbage)
    if len(rec['surname']) > 20:
        return False
    # Reject if surname has digits
    if any(c.isdigit() for c in rec['surname']):
        return False
    # Reject if full name contains "không giải thích" or other noise markers
    noise_markers = ['không giải thích', 'nhưng ', 'theo ', 'của ', 'với ', 'bộ ',
                     'creative commons', 'plos one', 'journal of', 'wiley', 'springer']
    full_lower = rec['full_name'].lower()
    for marker in noise_markers:
        if marker in full_lower:
            return False
    return True

# Filter
clean_authors = {aid: rec for aid, rec in authors.items() if is_valid_author(aid, rec)}
rejected = len(authors) - len(clean_authors)
print(f'Rejected {rejected} noise author entries')
print(f'Valid authors: {len(clean_authors)}')

# Build graph
G = nx.MultiDiGraph()

# === NODES ===
# Author nodes
for aid, rec in clean_authors.items():
    G.add_node(aid,
               type='Author',
               label=rec['full_name'],
               surname=rec['surname'],
               given=rec.get('given', ''),
               is_vietnamese=rec.get('is_vietnamese', False),
               n_papers=rec['n_papers'],
               country_hint=rec.get('country_hint', ''))

# Paper nodes
for cid, meta in canon.items():
    G.add_node(cid,
               type='Paper',
               label=meta.get('descriptor', cid),
               year=meta.get('descriptor', '').split('_')[-2] if '_' in meta.get('descriptor', '') else '',
               pdf_folder=meta.get('pdf_folder', ''),
               has_pdf=(meta.get('pdf_folder') and meta.get('pdf_folder') != 'None'))

# Country nodes (one per unique country_hint)
country_set = set()
for rec in clean_authors.values():
    c = rec.get('country_hint', '')
    if c:
        country_set.add(c)
for c in country_set:
    cid = f'COUNTRY_{c.replace(" ", "_").upper()}'
    G.add_node(cid, type='Country', label=c)

# Topic nodes (inferred from paper descriptors)
# Keywords → topic
TOPIC_KEYWORDS = {
    'TOP_ANXIETY': ['anxiety', 'lo âu', 'GAD', 'SAD', 'SocialAnxiety', 'DASS'],
    'TOP_DEPRESSION': ['depression', 'trầm cảm', 'CESD', 'PHQ', 'MDD'],
    'TOP_CBT': ['CBT', 'iCBT', 'CA-CBT', 'Maya'],
    'TOP_MHPSS': ['MHPSS', 'MentalHealth', 'Psychosocial'],
    'TOP_ADOLESCENT': ['Adolescent', 'VTN', 'HS', 'student', 'school', 'youth'],
    'TOP_COVID': ['COVID', 'pandemic'],
    'TOP_DIGITAL': ['Digital', 'DMHI', 'app', 'online', 'iCBT', 'Mobile', 'internet'],
    'TOP_RESILIENCE': ['Resilience', 'resilient'],
    'TOP_BULLYING': ['Bullying', 'bắt nạt', 'Cyberbullying'],
    'TOP_SUICIDE': ['Suicide', 'self-harm', 'SuicideRisk'],
    'TOP_SCHOOL_INTERVENTION': ['School', 'SchoolBased', 'classroom', 'BESST', 'HappyHouse'],
    'TOP_PREVALENCE': ['Prevalence', 'Epidemiology', 'survey'],
    'TOP_GENDER': ['gender', 'female', 'nam nữ'],
    'TOP_VN': ['Vietnam', 'VN', 'Hanoi', 'TPHCM', 'Hue'],
    'TOP_LMIC': ['LMIC', 'low', 'middle income'],
}
for tid, _ in TOPIC_KEYWORDS.items():
    G.add_node(tid, type='Topic', label=tid.replace('TOP_', '').title())

# === EDGES ===
# AUTHORED_BY: Paper → Author
for edge in paper_author_edges:
    if edge['author_id'] in clean_authors:
        G.add_edge(edge['paper_id'], edge['author_id'],
                   type='AUTHORED_BY',
                   position=edge['position'],
                   is_first_author=edge['is_first_author'])

# CO_AUTHORED: Author → Author (for each pair sharing a paper)
coauthor_pairs = defaultdict(lambda: {'shared_papers': [], 'count': 0})
paper_to_authors = defaultdict(list)
for edge in paper_author_edges:
    if edge['author_id'] in clean_authors:
        paper_to_authors[edge['paper_id']].append(edge['author_id'])

for paper_id, aids in paper_to_authors.items():
    for i in range(len(aids)):
        for j in range(i+1, len(aids)):
            a, b = aids[i], aids[j]
            key = tuple(sorted([a, b]))
            coauthor_pairs[key]['shared_papers'].append(paper_id)
            coauthor_pairs[key]['count'] += 1

for (a, b), info in coauthor_pairs.items():
    G.add_edge(a, b,
               type='CO_AUTHORED',
               shared_papers=';'.join(info['shared_papers']),
               strength=info['count'])
    G.add_edge(b, a,
               type='CO_AUTHORED',
               shared_papers=';'.join(info['shared_papers']),
               strength=info['count'])

# LOCATED_IN: Author → Country (via country_hint)
for aid, rec in clean_authors.items():
    c = rec.get('country_hint', '')
    if c:
        cnode = f'COUNTRY_{c.replace(" ", "_").upper()}'
        G.add_edge(aid, cnode, type='LOCATED_IN')

# STUDIES: Author → Topic (inferred via papers → topic keywords)
author_topics = defaultdict(set)
for aid, rec in clean_authors.items():
    for pid in rec['papers']:
        paper_descr = canon.get(pid, {}).get('descriptor', '')
        for topic_id, keywords in TOPIC_KEYWORDS.items():
            if any(kw.lower() in paper_descr.lower() for kw in keywords):
                author_topics[aid].add(topic_id)

for aid, topics in author_topics.items():
    for tid in topics:
        paper_count = sum(1 for pid in clean_authors[aid]['papers']
                          if any(kw.lower() in canon.get(pid, {}).get('descriptor', '').lower()
                                 for kw in TOPIC_KEYWORDS[tid]))
        G.add_edge(aid, tid, type='STUDIES', paper_count=paper_count)

# === STATS ===
n_author = sum(1 for n, d in G.nodes(data=True) if d.get('type') == 'Author')
n_paper = sum(1 for n, d in G.nodes(data=True) if d.get('type') == 'Paper')
n_country = sum(1 for n, d in G.nodes(data=True) if d.get('type') == 'Country')
n_topic = sum(1 for n, d in G.nodes(data=True) if d.get('type') == 'Topic')

e_types = defaultdict(int)
for u, v, k, d in G.edges(keys=True, data=True):
    e_types[d.get('type', 'UNKNOWN')] += 1

print('='*70)
print('PHASE C — AUTHOR KG v1')
print('='*70)
print(f'Nodes: {G.number_of_nodes()} total')
print(f'  Authors: {n_author}')
print(f'  Papers: {n_paper}')
print(f'  Countries: {n_country}')
print(f'  Topics: {n_topic}')
print(f'Edges: {G.number_of_edges()} total')
for et, c in sorted(e_types.items(), key=lambda x: -x[1]):
    print(f'  {et}: {c}')

# === VALIDATION ===
print('\nValidation:')
# R1: every paper has ≥ 1 AUTHORED_BY
papers_no_author = []
for n, d in G.nodes(data=True):
    if d.get('type') == 'Paper':
        has_ab = any(dd.get('type') == 'AUTHORED_BY' for _, _, dd in G.out_edges(n, data=True))
        if not has_ab:
            papers_no_author.append(n)
print(f'R1 (paper has ≥ 1 author): {len(papers_no_author)} violations')
if papers_no_author:
    print(f'   Violations: {papers_no_author[:5]}...' if len(papers_no_author) > 5 else f'   {papers_no_author}')

# R4: "Sasaki" not in authors
sasaki = [n for n, d in G.nodes(data=True) if d.get('type') == 'Author' and 'SASAKI' in n]
print(f'R4 (Sasaki excluded): {"✓" if not sasaki else "✗ FOUND: " + str(sasaki)}')

# === EXPORT ===
# GraphML — convert None to empty string, bool to str (GraphML restrictions)
for n, d in G.nodes(data=True):
    for k, v in list(d.items()):
        if v is None:
            G.nodes[n][k] = ''
        elif isinstance(v, bool):
            G.nodes[n][k] = str(v)
        elif isinstance(v, (list, dict, set)):
            G.nodes[n][k] = str(v)

for u, v, k, d in G.edges(keys=True, data=True):
    for kk, vv in list(d.items()):
        if vv is None:
            G.edges[u, v, k][kk] = ''
        elif isinstance(vv, bool):
            G.edges[u, v, k][kk] = str(vv)
        elif isinstance(vv, (list, dict, set)):
            G.edges[u, v, k][kk] = str(vv)

graphml_path = os.path.join(OUT_DIR, 'author_kg_v1.graphml')
nx.write_graphml(G, graphml_path)
print(f'\nGraphML saved: {graphml_path}')

# JSON (nodes + edges list)
nodes_list = [{'id': n, **d} for n, d in G.nodes(data=True)]
edges_list = [{'source': u, 'target': v, **d} for u, v, d in G.edges(data=True)]

json_data = {
    'meta': {
        'created': '2026-04-15',
        'n_nodes': G.number_of_nodes(),
        'n_edges': G.number_of_edges(),
        'node_types': {'Author': n_author, 'Paper': n_paper, 'Country': n_country, 'Topic': n_topic},
        'edge_types': dict(e_types),
    },
    'nodes': nodes_list,
    'edges': edges_list,
}
json_path = os.path.join(OUT_DIR, 'author_kg_v1.json')
with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(json_data, f, ensure_ascii=False, indent=2, default=str)
print(f'JSON saved: {json_path}')

# Top collaboration pairs
print('\nTop 10 collaboration pairs (strength):')
top_collab = sorted(coauthor_pairs.items(), key=lambda x: -x[1]['count'])[:10]
for (a, b), info in top_collab:
    a_name = clean_authors.get(a, {}).get('full_name', a)
    b_name = clean_authors.get(b, {}).get('full_name', b)
    print(f'  {a_name} ↔ {b_name} (share {info["count"]} papers)')
