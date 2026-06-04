"""Build papers_metadata.json - augment canonical_index with extra facets.

Extracts: year, lead_author, region, doc_types, topics, method, linked_questions.
Output: tro-ly-nghien-cuu-tam-ly-light/web/data/papers_metadata.json
"""
import sys, io, json, re
from pathlib import Path
from collections import Counter
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = Path('c:/Users/HLC/OneDrive/read_books/Lo-au')
CANONICAL = ROOT / '02_Papers-goc/canonical_index.json'
QA_INDEX = ROOT / 'tro-ly-nghien-cuu-tam-ly-light/web/data/rag_questions_index.json'
OUT_LIGHT = ROOT / 'tro-ly-nghien-cuu-tam-ly-light/web/data/papers_metadata.json'
OUT_HEAVY = ROOT / 'tro-ly-nghien-cuu-tam-ly/web/data/papers_metadata.json'

# Region map: pdf_folder -> facet label
REGION_MAP = {
    'Viet-Nam': 'Việt Nam',
    'The-gioi_Au-My-Uc': 'Âu-Mỹ-Úc',
    'The-gioi_Khac': 'Khác',
    'Dong-Nam-A': 'Đông Nam Á',
    '11-bai-ban-dau-va-mo-rong': 'Tổng hợp ban đầu',
}

# Topic heuristics — keyword in descriptor → topic tag
TOPIC_KEYWORDS = {
    'Anxiety': ['anxiety', 'lo_au', 'GAD', 'gad7', 'gad-7', 'phobia', 'sad', 'sas', 'anx'],
    'Depression': ['depression', 'tram_cam', 'phq9', 'phq-9', 'cesd', 'ces-d', 'mdd', 'depress'],
    'Stress': ['stress', 'academicstress', 'pss', 'burnout', 'pressure'],
    'Intervention': ['cbt', 'mindfulness', 'rct', 'intervention', 'program', 'therapy', 'treatment',
                     'eacp', 'copingpower', 'peersupport', 'photovoice', 'app', 'digital'],
    'Methodology': ['systematicreview', 'metaanalysis', 'meta-analysis', 'scopingreview', 'narrativereview',
                    'review', 'prisma', 'sr', 'ma', 'nma'],
    'Resilience': ['resilience', 'phuc_hoi', 'coping', 'self-efficacy', 'protective'],
    'Sleep': ['sleep', 'giac_ngu', 'insomnia'],
    'Suicide': ['suicide', 'self_harm', 'tu_sat', 'tu_lam_hai'],
    'Family/Parenting': ['parenting', 'family', 'gia_dinh', 'cha_me', 'parent'],
    'School': ['school', 'student', 'truong', 'classroom', 'teacher', 'gv', 'hocsinh', 'thcs', 'thpt'],
    'Mental Health (general)': ['mentalhealth', 'sktt', 'unicef', 'vnamhs', 'gbd', 'wellbeing', 'who'],
    'Bullying': ['bully', 'bat_nat', 'cyberbully'],
    'Substance use': ['substance', 'alcohol', 'drug', 'smoking'],
    'Help-seeking': ['help_seeking', 'helpseeking', 'service', 'tham_van', 'tu_van'],
}

# Method heuristics — keyword in descriptor or summary → method tag
METHOD_KEYWORDS = {
    'RCT': ['rct', 'randomized', 'randomised', 'eacp_rct'],
    'Meta-analysis': ['metaanalysis', 'meta-analysis', 'metaanalys'],
    'Systematic review': ['systematicreview', 'systematic_review'],
    'Network meta-analysis': ['networkmeta', 'nma'],
    'Scoping review': ['scopingreview', 'scoping_review'],
    'Narrative review': ['narrativereview', 'narrative_review'],
    'Cross-sectional': ['crosssectional', 'cross-sectional', 'survey', 'khaosat'],
    'Longitudinal': ['longitudinal', 'cohort', 'thuần_tập', 'thuantap'],
    'Photovoice/Qualitative': ['photovoice', 'qualitative', 'fgd'],
    'Latent profile': ['lpa', 'latentprofile'],
}

LIGHT_DOC_DIR = ROOT / 'tro-ly-nghien-cuu-tam-ly-light/web/docs'

def extract_year(descriptor: str):
    m = re.search(r'(19\d{2}|20\d{2})', descriptor)
    return int(m.group(1)) if m else None

def extract_lead_author(descriptor: str):
    """Heuristic: first underscore-separated token that looks like a surname."""
    parts = descriptor.split('_')
    if not parts:
        return ''
    # Skip common prefixes
    skip = {'the', 'a', 'an', 'study', 'survey', 'report'}
    for p in parts:
        if p.lower() in skip: continue
        # Avoid year tokens / pure-digit
        if re.fullmatch(r'\d+', p): continue
        # Avoid abbreviation-like (all caps + short)
        if len(p) <= 2: continue
        # Avoid country/region abbreviations
        if p.lower() in ('vn', 'qt', 'usa', 'uk', 'eu'): continue
        return p
    return parts[0]

def year_bucket(year):
    if year is None: return 'Không rõ'
    if year < 2010: return 'Trước 2010'
    if year < 2016: return '2010-2015'
    if year < 2021: return '2016-2020'
    return '2021+'

def detect_topics(descriptor: str, summary_path: str):
    haystack = (descriptor + ' ' + summary_path).lower().replace('-', '').replace(' ', '')
    found = []
    for topic, kws in TOPIC_KEYWORDS.items():
        for kw in kws:
            if kw.replace('-', '').replace(' ', '') in haystack:
                found.append(topic)
                break
    if not found:
        found.append('Mental Health (general)')
    return found

def detect_method(descriptor: str, existing_method: str = None):
    if existing_method:
        # Normalize known canonical method strings
        em = existing_method.lower()
        if 'rct' in em: return 'RCT'
        if 'scoping' in em: return 'Scoping review'
        if 'narrative' in em: return 'Narrative review'
        if 'meta' in em and 'network' in em: return 'Network meta-analysis'
        if 'meta' in em: return 'Meta-analysis'
        if 'systematic' in em: return 'Systematic review'
        if 'cross' in em: return 'Cross-sectional'
        if 'longitud' in em: return 'Longitudinal'
        return existing_method
    haystack = descriptor.lower().replace('-', '').replace('_', '')
    for method, kws in METHOD_KEYWORDS.items():
        for kw in kws:
            if kw.replace('-', '').replace('_', '') in haystack:
                return method
    return 'Khác/Chưa rõ'

def doc_type_list(entry):
    types = []
    if entry.get('pdf'): types.append('original')
    if entry.get('translation'): types.append('translation')
    if entry.get('summary'): types.append('summary')
    return types

def main():
    with open(CANONICAL, 'r', encoding='utf-8') as f:
        canonical = json.load(f)
    with open(QA_INDEX, 'r', encoding='utf-8') as f:
        qa = json.load(f)

    # Build paper -> [QA_id] map from rag_questions paper_refs
    paper_to_qids = {}
    for qe in qa['entries']:
        for pref in qe.get('paper_refs', []):
            paper_to_qids.setdefault(pref, []).append(qe['id'])

    papers = []
    region_ct = Counter()
    method_ct = Counter()
    year_bucket_ct = Counter()
    topic_ct = Counter()
    doc_type_ct = Counter()
    lead_author_ct = Counter()

    for pid in sorted(canonical.keys()):
        e = canonical[pid]
        descriptor = e.get('descriptor', '')
        year = extract_year(descriptor)
        lead = extract_lead_author(descriptor)
        region = REGION_MAP.get(e.get('pdf_folder') or '', 'Không rõ')
        topics = detect_topics(descriptor, e.get('summary', ''))
        method = detect_method(descriptor, e.get('method'))
        doc_types = doc_type_list(e)
        linked_qs = paper_to_qids.get(pid, [])

        # Build files dict (relative paths in /web/docs/ — actual presence not verified here)
        files = {}
        if e.get('pdf'): files['pdf'] = e['pdf']
        if e.get('translation'): files['translation'] = e['translation']
        if e.get('summary'): files['summary'] = e['summary']

        record = {
            'id': pid,
            'descriptor': descriptor,
            'lead_author': lead,
            'year': year,
            'year_bucket': year_bucket(year),
            'region': region,
            'is_vn': pid.startswith('VN'),
            'doc_types': doc_types,
            'method': method,
            'topics': topics,
            'linked_questions': linked_qs,
            'pdf_folder': e.get('pdf_folder'),
            'doi': e.get('doi'),
            'pubmed': e.get('pubmed'),
            'files': files,
        }
        papers.append(record)

        # Tally
        region_ct[region] += 1
        method_ct[method] += 1
        year_bucket_ct[year_bucket(year)] += 1
        for t in topics: topic_ct[t] += 1
        for dt in doc_types: doc_type_ct[dt] += 1
        if lead: lead_author_ct[lead] += 1

    output = {
        'meta': {
            'created': datetime.now().strftime('%Y-%m-%d'),
            'version': 'papers_metadata-v1',
            'n_entries': len(papers),
            'description': 'Augmented metadata for faceted browser. 7 facets: region, doc_type, lead_author, linked_question, year_bucket, method, topic.',
        },
        'facet_options': {
            'region': sorted(region_ct.items(), key=lambda x: -x[1]),
            'method': sorted(method_ct.items(), key=lambda x: -x[1]),
            'year_bucket': sorted(year_bucket_ct.items(), key=lambda x: -x[1]),
            'topic': sorted(topic_ct.items(), key=lambda x: -x[1]),
            'doc_type': sorted(doc_type_ct.items(), key=lambda x: -x[1]),
            'top_authors': lead_author_ct.most_common(30),
        },
        'entries': papers,
    }

    OUT_LIGHT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT_LIGHT, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f'Saved LIGHT: {OUT_LIGHT}')
    print(f'  Size: {OUT_LIGHT.stat().st_size//1024} KB, entries: {len(papers)}')

    # Mirror to heavy if folder exists
    if OUT_HEAVY.parent.exists():
        with open(OUT_HEAVY, 'w', encoding='utf-8') as f:
            json.dump(output, f, ensure_ascii=False, indent=2)
        print(f'Saved HEAVY: {OUT_HEAVY}')

    print()
    print('=== Facet summary ===')
    print(f'Region:       {dict(region_ct)}')
    print(f'Method:       {dict(method_ct)}')
    print(f'Year bucket:  {dict(year_bucket_ct)}')
    print(f'Doc type:     {dict(doc_type_ct)}')
    print(f'Topics:       {dict(topic_ct.most_common(10))}')
    print(f'Top authors:  {lead_author_ct.most_common(10)}')

if __name__ == '__main__':
    main()
