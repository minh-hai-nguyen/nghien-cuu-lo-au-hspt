# -*- coding: utf-8 -*-
"""
Build TF-IDF keyword index cho bản lite.
Extract papers + insights từ rebuild_rag_full.py (heavy) bằng ast.literal_eval.
Output: web/data/rag_main_index.json (~200-500 KB).
"""
import os, sys, json, re, math, ast
from pathlib import Path
from collections import Counter

os.environ['PYTHONIOENCODING'] = 'utf-8'
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = Path(__file__).parent
HEAVY_REBUILD = BASE.parent / 'tro-ly-nghien-cuu-tam-ly' / 'rebuild_rag_full.py'
OUT = BASE / 'web' / 'data' / 'rag_main_index.json'

# ---- Stopwords (EN + VN), copy từ app.py heavy version ----
STOPWORDS = set('''a an and or but in on at to from for of with by is are was were be been being have has had do does did the this that these those it its they them their we us our you your he she his her i me my là và hoặc nhưng trong tại đến từ cho của với bởi là các một những này đó nó họ chúng ta chúng tôi tôi bạn anh ấy cô ấy có được bị đã đang sẽ cũng không chỉ rất đều như ai ở khi nào tới sau trước giữa ngoài trên dưới bên cùng qua vào ra lên xuống đây đó kia'''.split())

def tokenize(text):
    if not text: return []
    text = text.lower()
    tokens = re.findall(r'\b[\wÀ-ỹđĐ]+\b', text, flags=re.UNICODE)
    return [t for t in tokens if t not in STOPWORDS and len(t) > 1]

def extract_papers_and_insights(rebuild_path):
    """Extract papers + insights bằng exec với mock read_docx."""
    src = rebuild_path.read_text(encoding='utf-8')
    # Extract only the 'papers = [...]' and 'insights = [...]' blocks
    ns = {'read_docx': lambda *a, **kw: '', 'os': os, 'BASE': str(rebuild_path.parent)}  # mock — skip docx loading
    # Find block start/end for each list by brace matching
    def extract_list_block(name):
        m = re.search(rf'^{name}\s*=\s*\[', src, re.MULTILINE)
        if not m:
            return None
        i = m.end() - 1  # position of '['
        depth = 0
        while i < len(src):
            c = src[i]
            if c == '[':
                depth += 1
            elif c == ']':
                depth -= 1
                if depth == 0:
                    return src[m.start():i+1]
            i += 1
        return None
    papers_block = extract_list_block('papers')
    insights_block = extract_list_block('insights')
    if not papers_block or not insights_block:
        raise RuntimeError('Không extract được papers/insights block.')
    exec(papers_block, ns)
    exec(insights_block, ns)
    return ns['papers'], ns['insights']

def main():
    print('=== BUILD RAG MAIN INDEX (TF-IDF) ===')
    papers, insights = extract_papers_and_insights(HEAVY_REBUILD)
    print(f'Extracted {len(papers)} papers + {len(insights)} insights')

    entries = []

    # --- Papers ---
    for p in papers:
        text_parts = [
            p.get('title', ''),
            p.get('authors', ''),
            p.get('journal', ''),
            p.get('location', ''),
            p.get('tool', ''),
            f"Khu vực: {p.get('region', '')}",
            f"Mẫu: {p.get('n', '')}",
            f"Phát hiện: {p.get('key', '')}",
        ]
        text = '\n'.join([t for t in text_parts if t])
        tokens = tokenize(text)
        tf = dict(Counter(tokens))
        entries.append({
            'id': p.get('id', ''),
            'type': 'paper',
            'title': p.get('title', ''),
            'text': text,
            'tokens': tf,
            'meta': {
                'id': p.get('id', ''),
                'region': p.get('region', ''),
                'region_code': p.get('region_code', ''),
                'authors': p.get('authors', ''),
                'journal': p.get('journal', ''),
                'location': p.get('location', ''),
                'tool': p.get('tool', ''),
                'n': p.get('n', ''),
                'file_dich': p.get('file_dich', ''),
                'file_tt': p.get('file_tt', ''),
            }
        })

    # --- Insights ---
    for ins in insights:
        text = f"{ins.get('title', '')}\n{ins.get('content', '')}"
        tokens = tokenize(text)
        tf = dict(Counter(tokens))
        entries.append({
            'id': ins.get('id', ''),
            'type': 'insight',
            'title': ins.get('title', ''),
            'text': text,
            'tokens': tf,
            'meta': {
                'id': ins.get('id', ''),
                'region': ins.get('region', ''),
                'region_code': ins.get('region_code', ''),
            }
        })

    # --- Compute IDF ---
    N = len(entries)
    df = Counter()
    for e in entries:
        for tok in e['tokens']:
            df[tok] += 1
    idf = {tok: math.log((N + 1) / (dfi + 1)) + 1 for tok, dfi in df.items()}

    # --- Save ---
    OUT.parent.mkdir(parents=True, exist_ok=True)
    out_data = {
        'meta': {
            'created': '2026-04-20',
            'version': 'lite-v1',
            'n_entries': N,
            'n_papers': len(papers),
            'n_insights': len(insights),
            'total_unique_tokens': len(idf),
        },
        'idf': idf,
        'entries': entries,
    }
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(out_data, f, ensure_ascii=False, indent=1)

    size_kb = OUT.stat().st_size / 1024
    print(f'Saved: {OUT}')
    print(f'Size: {size_kb:.1f} KB')
    print(f'Stats: {N} entries ({len(papers)} papers + {len(insights)} insights), {len(idf)} unique tokens')

    # --- Sanity check ---
    q = 'CBT lo âu vị thành niên'
    q_toks = tokenize(q)
    scores = []
    for e in entries:
        s = sum(e['tokens'].get(t, 0) * idf.get(t, 1.0) for t in q_toks)
        if s > 0:
            scores.append((s, e['id'], e['title']))
    scores.sort(reverse=True)
    print(f"\nSanity query: '{q}' → top-3:")
    for s, eid, title in scores[:3]:
        print(f'  {eid} [{s:.2f}]  {title}')

if __name__ == '__main__':
    main()
