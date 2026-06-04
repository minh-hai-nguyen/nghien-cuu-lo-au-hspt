# -*- coding: utf-8 -*-
"""
Phase E — Author RAG (keyword + TF-IDF style, no chromadb needed).

Build a lightweight retrieval index for author profiles:
- Each author → profile text (bilingual EN + VN)
- Index: tokenized, stopword-filtered, lowercase
- Query: tokenize + score via token overlap with boost for exact matches
- Ranking: score + priority boost

Output: rag_authors_index.json (can be queried with rag_authors_query.py)
"""
import os, sys, io, json, re, math
from collections import defaultdict, Counter
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
OUT_DIR = os.path.join(os.path.dirname(__file__), 'author_kg_data')

with open(os.path.join(OUT_DIR, 'authors_normalized.json'), encoding='utf-8') as f:
    authors = json.load(f)
with open(os.path.join(OUT_DIR, 'author_kg_v1.json'), encoding='utf-8') as f:
    kg = json.load(f)
with open(os.path.join(ROOT, '02_Papers-goc', 'canonical_index.json'), encoding='utf-8') as f:
    canon = json.load(f)

# Stopwords (EN + VN common)
STOPWORDS = set('''a an and or but in on at to from for of with by is are was were
be been being have has had do does did the this that these those it its they them
their we us our you your he she his her i me my
là và hoặc nhưng trong tại đến từ cho của với bởi là các một những này đó
nó họ của họ chúng ta chúng tôi tôi bạn bạn của tôi anh ấy cô ấy của anh ấy
có được bị đã đang sẽ cũng không chỉ rất đều như'''.split())

def tokenize(text):
    """Lowercase, extract word tokens."""
    if not text: return []
    text = text.lower()
    # Keep Vietnamese chars, letters, digits
    tokens = re.findall(r'\b[\wÀ-ỹđĐ]+\b', text, flags=re.UNICODE)
    return [t for t in tokens if t not in STOPWORDS and len(t) > 1]

# Build CO_AUTHORED map for profile text
coauthor_map = defaultdict(list)
for e in kg['edges']:
    if e.get('type') == 'CO_AUTHORED':
        coauthor_map[e['source']].append((e['target'], int(e.get('strength', 1))))

def build_profile_text(aid, rec):
    """Compose bilingual profile for an author."""
    name = rec['full_name']
    country = rec.get('country_primary') or rec.get('country_hint', '')
    affs = rec.get('affiliations', [])
    expertise = rec.get('expertise', [])
    papers = rec.get('papers', [])
    role = rec.get('role_notes', '')
    n_papers = rec.get('n_papers', 0)
    is_priority = rec.get('is_priority', False)
    is_child_focus = rec.get('child_adolescent_focus', False)
    is_vn = rec.get('is_vietnamese', False)

    # Top 5 collaborators
    collabs_raw = sorted(coauthor_map.get(aid, []), key=lambda x: -x[1])[:5]
    collab_names = []
    for cid, _strength in collabs_raw:
        if cid in authors:
            collab_names.append(authors[cid].get('full_name', cid))

    # Paper titles (from canonical descriptor)
    paper_info = []
    for pid in papers:
        descr = canon.get(pid, {}).get('descriptor', '')
        paper_info.append(f'{pid} ({descr})')

    # EN text
    en_lines = [f"Author: {name}."]
    if country:
        en_lines.append(f"Country: {country}.")
    if affs:
        en_lines.append(f"Affiliations: {'; '.join(affs)}.")
    if expertise:
        en_lines.append(f"Research expertise: {', '.join(expertise)}.")
    en_lines.append(f"Papers in library: {n_papers}.")
    if paper_info:
        en_lines.append(f"Papers: {'; '.join(paper_info[:8])}.")
    if collab_names:
        en_lines.append(f"Top collaborators: {', '.join(collab_names)}.")
    if role:
        en_lines.append(f"Notable role: {role}.")
    if is_priority:
        en_lines.append("This is a priority author in child/adolescent psychology.")
    if is_child_focus:
        en_lines.append("Focus: child and adolescent mental health.")
    if is_vn:
        en_lines.append("Vietnamese researcher.")

    # VN text
    vn_lines = [f"Tác giả: {name}."]
    if country:
        vn_lines.append(f"Quốc gia: {country}.")
    if affs:
        vn_lines.append(f"Đơn vị: {'; '.join(affs)}.")
    if expertise:
        vn_lines.append(f"Chuyên môn: {', '.join(expertise)}.")
    vn_lines.append(f"Số paper trong thư viện: {n_papers}.")
    if paper_info:
        vn_lines.append(f"Các paper: {'; '.join(paper_info[:8])}.")
    if collab_names:
        vn_lines.append(f"Cộng tác viên hàng đầu: {', '.join(collab_names)}.")
    if role:
        vn_lines.append(f"Vai trò nổi bật: {role}.")
    if is_priority:
        vn_lines.append("Đây là tác giả ưu tiên trong tâm lý học trẻ em/vị thành niên.")
    if is_child_focus:
        vn_lines.append("Tập trung: sức khoẻ tâm thần trẻ em và vị thành niên.")
    if is_vn:
        vn_lines.append("Nhà nghiên cứu người Việt Nam.")

    return ' '.join(en_lines), ' '.join(vn_lines), {
        'n_collaborators': len(collab_names),
        'collab_names': collab_names,
    }

# Build index entries
entries = []
df_counter = Counter()  # document frequency for IDF

for aid, rec in authors.items():
    en_text, vn_text, extras = build_profile_text(aid, rec)
    full_text = f"{en_text}\n\n{vn_text}"
    tokens = tokenize(full_text)
    tf = Counter(tokens)

    # Also add name tokens heavily (boost)
    name_tokens = tokenize(rec['full_name'])
    for t in name_tokens:
        tf[t] += 5  # boost name tokens

    entry = {
        'id': aid,
        'name': rec['full_name'],
        'text_en': en_text,
        'text_vn': vn_text,
        'tokens': dict(tf),
        'meta': {
            'surname': rec['surname'],
            'country': rec.get('country_primary') or rec.get('country_hint', ''),
            'affiliations': rec.get('affiliations', []),
            'expertise': rec.get('expertise', []),
            'papers': rec.get('papers', []),
            'n_papers': rec.get('n_papers', 0),
            'is_priority': rec.get('is_priority', False),
            'child_adolescent_focus': rec.get('child_adolescent_focus', False),
            'is_vietnamese': rec.get('is_vietnamese', False),
            'n_collaborators': extras['n_collaborators'],
            'collab_names': extras['collab_names'],
        }
    }
    entries.append(entry)

    # Update DF
    for t in set(tokens):
        df_counter[t] += 1

# Compute IDF
N = len(entries)
idf = {t: math.log(N / (df + 1)) + 1 for t, df in df_counter.items()}

# Save index
index_data = {
    'meta': {
        'created': '2026-04-15',
        'n_authors': N,
        'total_unique_tokens': len(df_counter),
    },
    'idf': idf,
    'entries': entries,
}

out = os.path.join(OUT_DIR, 'rag_authors_index.json')
with open(out, 'w', encoding='utf-8') as f:
    json.dump(index_data, f, ensure_ascii=False, indent=2)

print('='*70)
print('PHASE E — AUTHOR RAG INDEX BUILT')
print('='*70)
print(f'Authors indexed: {N}')
print(f'Unique tokens: {len(df_counter):,}')
print(f'Index file: {out}')
print(f'Index size: {os.path.getsize(out) / 1024:.1f} KB')

# Sample entry
print()
print('Sample entry (AU_MATSUMOTO_K):')
for e in entries:
    if e['id'] == 'AU_MATSUMOTO_K':
        print(f'  Name: {e["name"]}')
        print(f'  Text EN: {e["text_en"][:200]}...')
        print(f'  Text VN: {e["text_vn"][:200]}...')
        print(f'  Meta: country={e["meta"]["country"]}, priority={e["meta"]["is_priority"]}, papers={e["meta"]["papers"]}')
        print(f'  Top 10 tokens by count: {Counter(e["tokens"]).most_common(10)}')
        break
