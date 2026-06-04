# -*- coding: utf-8 -*-
"""
Trợ lý Nghiên cứu Tâm lý học — Bản LITE (TF-IDF, không dùng sentence-transformers)
Phù hợp Render free tier 512 MB RAM.

Các endpoint:
- GET  /                       serve index.html
- POST /api/query              main RAG search (TF-IDF 47 entries)
- GET  /api/documents          list PDFs/DOCXs trong /docs
- POST /api/authors/query      tác giả search
- GET  /api/authors/list       tác giả list + filter
- GET  /api/authors/{id}       tác giả detail
- GET  /api/authors/stats      tác giả stats
- GET  /api/glossary           glossary all
- GET  /api/glossary/search    glossary search
- GET  /api/glossary/categories glossary categories
- GET  /api/papers/list        faceted browser — filter + group_by
- GET  /api/papers/facets      facet options + counts
- GET  /api/papers/detail/{id} paper full metadata + linked Q&A
"""
import os, sys, json, re, math
from pathlib import Path
from collections import Counter as _Counter
from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, JSONResponse

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / 'data'
DOCS_DIR = BASE_DIR / 'docs'
STATIC_DIR = BASE_DIR / 'static'

app = FastAPI(title='Trợ lý NC Tâm lý học (LITE)')

# =========================================================
# Tokenizer + stopwords (EN + VN) — dùng chung main + author
# =========================================================
STOPWORDS = set('''a an and or but in on at to from for of with by is are was were be been being
have has had do does did the this that these those it its they them their we us our you your
he she his her i me my là và hoặc nhưng trong tại đến từ cho của với bởi là các một những này đó
nó họ chúng ta chúng tôi tôi bạn anh ấy cô ấy có được bị đã đang sẽ cũng không chỉ rất đều như
ai ở khi nào tới sau trước giữa ngoài trên dưới bên cùng qua vào ra lên xuống đây đó kia'''.split())

def _tokenize(text):
    if not text: return []
    text = text.lower()
    tokens = re.findall(r'\b[\wÀ-ỹđĐ]+\b', text, flags=re.UNICODE)
    return [t for t in tokens if t not in STOPWORDS and len(t) > 1]

# =========================================================
# MAIN RAG — 47 entries (36 papers + 11 insights)
# =========================================================
_MAIN_INDEX = None

def _load_main_index():
    global _MAIN_INDEX
    try:
        p = DATA_DIR / 'rag_main_index.json'
        if p.exists():
            with open(p, encoding='utf-8') as f:
                _MAIN_INDEX = json.load(f)
            return True
    except Exception as e:
        print(f'Main index load failed: {e}')
    return False

_load_main_index()
if _MAIN_INDEX:
    print(f"Main RAG loaded: {_MAIN_INDEX['meta']['n_entries']} entries "
          f"({_MAIN_INDEX['meta']['n_papers']} papers + {_MAIN_INDEX['meta']['n_insights']} insights), "
          f"{_MAIN_INDEX['meta']['total_unique_tokens']} tokens")

def _main_score(qtokens, entry, idf):
    """TF-IDF dot product + title boost + ID match boost."""
    etok = entry['tokens']
    score = 0.0
    for qt in qtokens:
        if qt in etok:
            score += etok[qt] * idf.get(qt, 1.0)
    # Title match boost
    title_lower = entry.get('title', '').lower()
    for qt in qtokens:
        if qt in title_lower:
            score += 5
    # Exact ID match (e.g. "QT045", "VN01", "INSIGHT_06")
    eid = entry.get('id', '').lower()
    for qt in qtokens:
        if qt == eid or (len(qt) >= 3 and qt in eid):
            score += 30
    return score

@app.post('/api/query')
async def query(request: Request):
    if not _MAIN_INDEX:
        return JSONResponse({'error': 'Main index not loaded'}, status_code=500)
    data = await request.json()
    q = data.get('question', '').strip()
    top_k = int(data.get('n_results', 5))
    if not q:
        return JSONResponse({'error': 'Empty'}, status_code=400)

    qtokens = _tokenize(q)
    idf = _MAIN_INDEX.get('idf', {})
    scored = []
    for e in _MAIN_INDEX['entries']:
        s = _main_score(qtokens, e, idf)
        if s > 0:
            scored.append((s, e))
    scored.sort(key=lambda x: -x[0])

    top = scored[:top_k]
    max_s = top[0][0] if top else 1.0

    sources = []
    for s, e in top:
        meta = e.get('meta', {})
        # Build links
        links = {}
        file_dich = meta.get('file_dich', '')
        if file_dich and (DOCS_DIR / file_dich).exists():
            links['translation_docx'] = f'/docs/{file_dich}'
        file_tt = meta.get('file_tt', '')
        if file_tt and (DOCS_DIR / file_tt).exists():
            links['summary_docx'] = f'/docs/{file_tt}'
        sources.append({
            'text': e.get('text', '')[:600],
            'relevance': round(s / max_s * 100, 1),
            'title': e.get('title', ''),
            'journal': meta.get('journal', ''),
            'authors': meta.get('authors', ''),
            'region': meta.get('region', ''),
            'n': meta.get('n', ''),
            'tool': meta.get('tool', ''),
            'links': links,
            'page_ref': '',
            'next_question': '',
            'type': e.get('type', 'paper'),
            'id': e.get('id', ''),
        })

    return JSONResponse({'question': q, 'sources': sources})


@app.get('/api/documents')
async def list_documents():
    docs = []
    if DOCS_DIR.exists():
        for fname in sorted(os.listdir(DOCS_DIR)):
            if not fname.endswith(('.pdf', '.docx', '.md')):
                continue
            docs.append({
                'filename': fname,
                'title': fname,
                'download': f'/docs/{fname}',
            })
    return JSONResponse(docs)


# =========================================================
# AUTHOR KG + RAG — copy nguyên pattern từ bản heavy
# =========================================================
_AUTHOR_INDEX = None
_AUTHOR_KG = None
_AUTHOR_NORM = None

def _load_author_data():
    global _AUTHOR_INDEX, _AUTHOR_KG, _AUTHOR_NORM
    try:
        idx_path = DATA_DIR / 'rag_authors_index.json'
        kg_path = DATA_DIR / 'author_kg_v1.json'
        norm_path = DATA_DIR / 'authors_normalized.json'
        if idx_path.exists():
            with open(idx_path, encoding='utf-8') as f:
                _AUTHOR_INDEX = json.load(f)
        if kg_path.exists():
            with open(kg_path, encoding='utf-8') as f:
                _AUTHOR_KG = json.load(f)
        if norm_path.exists():
            with open(norm_path, encoding='utf-8') as f:
                _AUTHOR_NORM = json.load(f)
        return _AUTHOR_INDEX is not None
    except Exception as e:
        print(f'Author data load failed: {e}')
        return False

def _author_score(qtokens, entry, idf):
    etok = entry['tokens']
    score = 0.0
    for qt in qtokens:
        if qt in etok:
            score += etok[qt] * idf.get(qt, 1.0)
    surname_lower = entry['meta'].get('surname', '').lower()
    if surname_lower:
        for qt in qtokens:
            if qt == surname_lower or surname_lower in qt or qt in surname_lower:
                score += 20
    if entry['meta'].get('is_priority'):
        score += 3
    if entry['meta'].get('child_adolescent_focus'):
        score += 2
    return score

_load_author_data()
print(f'Author data loaded: {len(_AUTHOR_INDEX["entries"]) if _AUTHOR_INDEX else 0} authors')


@app.post('/api/authors/query')
async def authors_query(request: Request):
    if not _AUTHOR_INDEX:
        return JSONResponse({'error': 'Author index not loaded'}, status_code=500)
    data = await request.json()
    q = data.get('question', '').strip()
    top_k = int(data.get('n_results', 8))
    if not q:
        return JSONResponse({'error': 'Empty query'}, status_code=400)

    qtokens = _tokenize(q)
    idf = _AUTHOR_INDEX.get('idf', {})
    scored = []
    for e in _AUTHOR_INDEX['entries']:
        s = _author_score(qtokens, e, idf)
        if s > 0:
            scored.append((s, e))
    scored.sort(key=lambda x: -x[0])
    results = []
    for s, e in scored[:top_k]:
        results.append({
            'id': e['id'],
            'name': e['name'],
            'score': round(s, 2),
            'text_en': e.get('text_en', ''),
            'text_vn': e.get('text_vn', ''),
            'meta': e['meta'],
        })
    return JSONResponse({
        'question': q,
        'n_results': len(results),
        'total_scored': len(scored),
        'results': results,
    })


@app.get('/api/authors/list')
async def authors_list(priority_only: bool = False, child_focus: bool = False, country: str = ''):
    if not _AUTHOR_INDEX:
        return JSONResponse({'error': 'Author index not loaded'}, status_code=500)
    items = []
    for e in _AUTHOR_INDEX['entries']:
        m = e['meta']
        if priority_only and not m.get('is_priority'):
            continue
        if child_focus and not m.get('child_adolescent_focus'):
            continue
        if country and m.get('country', '').lower() != country.lower():
            continue
        items.append({
            'id': e['id'],
            'name': e['name'],
            'country': m.get('country', ''),
            'affiliations': m.get('affiliations', []),
            'expertise': m.get('expertise', []),
            'n_papers': m.get('n_papers', 0),
            'papers': m.get('papers', []),
            'is_priority': m.get('is_priority', False),
            'child_adolescent_focus': m.get('child_adolescent_focus', False),
            'is_vietnamese': m.get('is_vietnamese', False),
            'n_collaborators': m.get('n_collaborators', 0),
            'collab_names': m.get('collab_names', []),
        })
    items.sort(key=lambda x: (-int(x['is_priority']), -x['n_papers'], x['name']))
    return JSONResponse({'total': len(items), 'authors': items})


@app.get('/api/authors/stats')
async def authors_stats():
    if not _AUTHOR_INDEX:
        return JSONResponse({'error': 'not loaded'}, status_code=500)
    country_ct = _Counter()
    priority_n = 0
    child_focus_n = 0
    multi_paper_n = 0
    for e in _AUTHOR_INDEX['entries']:
        m = e['meta']
        if m.get('country'):
            country_ct[m['country']] += 1
        if m.get('is_priority'): priority_n += 1
        if m.get('child_adolescent_focus'): child_focus_n += 1
        if m.get('n_papers', 0) >= 2: multi_paper_n += 1
    return JSONResponse({
        'total_authors': len(_AUTHOR_INDEX['entries']),
        'priority_authors': priority_n,
        'child_adolescent_focused': child_focus_n,
        'multi_paper_authors': multi_paper_n,
        'top_countries': country_ct.most_common(15),
    })


@app.get('/api/authors/{author_id}')
async def author_detail(author_id: str):
    if not _AUTHOR_INDEX:
        return JSONResponse({'error': 'not loaded'}, status_code=500)
    for e in _AUTHOR_INDEX['entries']:
        if e['id'] == author_id:
            return JSONResponse({
                'id': e['id'],
                'name': e['name'],
                'text_en': e.get('text_en', ''),
                'text_vn': e.get('text_vn', ''),
                'meta': e['meta'],
            })
    return JSONResponse({'error': 'not found'}, status_code=404)


# =========================================================
# QUESTIONS (Q&A) — RAG search over questions thầy đã hỏi + KG
# =========================================================
_QUESTIONS_INDEX = None
_QUESTIONS_KG = None

def _load_questions():
    global _QUESTIONS_INDEX, _QUESTIONS_KG
    try:
        qidx = DATA_DIR / 'rag_questions_index.json'
        qkg = DATA_DIR / 'questions_kg.json'
        if qidx.exists():
            with open(qidx, encoding='utf-8') as f:
                _QUESTIONS_INDEX = json.load(f)
        if qkg.exists():
            with open(qkg, encoding='utf-8') as f:
                _QUESTIONS_KG = json.load(f)
        return _QUESTIONS_INDEX is not None
    except Exception as e:
        print(f'Questions load failed: {e}')
        return False

def _q_score(qtokens, entry, idf):
    etok = entry.get('tokens', {})
    score = 0.0
    for qt in qtokens:
        if qt in etok:
            score += etok[qt] * idf.get(qt, 1.0)
    for qt in qtokens:
        for c in entry.get('concepts', []):
            cl = c.lower()
            if qt == cl or qt in cl or cl in qt:
                score += 5
                break
    topic_l = entry.get('topic', '').lower()
    for qt in qtokens:
        if qt in topic_l:
            score += 3
    return score

_load_questions()
if _QUESTIONS_INDEX:
    print(f"Questions loaded: {_QUESTIONS_INDEX['meta']['n_entries']} Q&A, "
          f"{_QUESTIONS_INDEX['meta'].get('total_unique_tokens', 0)} tokens")

@app.post('/api/questions/query')
async def questions_query(request: Request):
    if not _QUESTIONS_INDEX:
        return JSONResponse({'error': 'Questions index not loaded'}, status_code=500)
    data = await request.json()
    q = data.get('question', '').strip()
    top_k = int(data.get('n_results', 5))
    if not q:
        return JSONResponse({'error': 'Empty query'}, status_code=400)
    qtokens = _tokenize(q)
    idf = _QUESTIONS_INDEX.get('idf', {})
    scored = []
    for e in _QUESTIONS_INDEX['entries']:
        s = _q_score(qtokens, e, idf)
        if s > 0:
            scored.append((s, e))
    scored.sort(key=lambda x: -x[0])
    results = []
    for s, e in scored[:top_k]:
        results.append({
            'id': e['id'],
            'question': e['question'],
            'short_answer': e['short_answer'],
            'topic': e['topic'],
            'concepts': e.get('concepts', []),
            'doc_path': e.get('doc_path', ''),
            'paper_refs': e.get('paper_refs', []),
            'date_asked': e.get('date_asked', ''),
            'score': round(s, 2),
        })
    return JSONResponse({
        'question': q,
        'n_results': len(results),
        'total_scored': len(scored),
        'results': results,
    })

@app.get('/api/questions/list')
async def questions_list(topic: str = ''):
    if not _QUESTIONS_INDEX:
        return JSONResponse({'error': 'Not loaded'}, status_code=500)
    items = []
    for e in _QUESTIONS_INDEX['entries']:
        if topic and e.get('topic', '').lower() != topic.lower():
            continue
        items.append({
            'id': e['id'],
            'question': e['question'],
            'topic': e.get('topic', ''),
            'concepts': e.get('concepts', []),
            'date_asked': e.get('date_asked', ''),
        })
    items.sort(key=lambda x: x['date_asked'], reverse=True)
    return JSONResponse({'total': len(items), 'questions': items})

@app.get('/api/questions/stats')
async def questions_stats():
    if not _QUESTIONS_INDEX:
        return JSONResponse({'error': 'Not loaded'}, status_code=500)
    topic_ct = _Counter()
    date_ct = _Counter()
    for e in _QUESTIONS_INDEX['entries']:
        topic_ct[e.get('topic', 'Other')] += 1
        date_ct[e.get('date_asked', 'Unknown')] += 1
    kg_stats = {}
    if _QUESTIONS_KG:
        kg_stats = _QUESTIONS_KG.get('meta', {})
    return JSONResponse({
        'total': len(_QUESTIONS_INDEX['entries']),
        'topics': topic_ct.most_common(),
        'dates': sorted(date_ct.items(), reverse=True),
        'kg': kg_stats,
    })

@app.get('/api/questions/{qid}')
async def questions_detail(qid: str):
    if not _QUESTIONS_INDEX:
        return JSONResponse({'error': 'Not loaded'}, status_code=500)
    for e in _QUESTIONS_INDEX['entries']:
        if e['id'] == qid:
            related = []
            if _QUESTIONS_KG:
                nmap = {n['id']: n for n in _QUESTIONS_KG.get('nodes', [])}
                for edge in _QUESTIONS_KG.get('edges', []):
                    if edge.get('source') == qid and edge.get('target') in nmap:
                        n = nmap[edge['target']]
                        related.append({
                            'id': n['id'],
                            'type': n['type'],
                            'label': n['label'],
                            'relation': edge['relation'],
                        })
            return JSONResponse({
                'id': e['id'],
                'question': e['question'],
                'short_answer': e['short_answer'],
                'topic': e['topic'],
                'concepts': e.get('concepts', []),
                'doc_path': e.get('doc_path', ''),
                'paper_refs': e.get('paper_refs', []),
                'date_asked': e.get('date_asked', ''),
                'related': related,
            })
    return JSONResponse({'error': 'not found'}, status_code=404)

# =========================================================
# PAPERS METADATA — Faceted browser ("Sắp xếp đa diện")
# =========================================================
_PAPERS_META = None

def _load_papers_metadata():
    global _PAPERS_META
    try:
        p = DATA_DIR / 'papers_metadata.json'
        if p.exists():
            with open(p, encoding='utf-8') as f:
                _PAPERS_META = json.load(f)
            return True
    except Exception as e:
        print(f'Papers metadata load failed: {e}')
    return False

_load_papers_metadata()
if _PAPERS_META:
    print(f"Papers metadata loaded: {_PAPERS_META['meta']['n_entries']} papers, "
          f"{len(_PAPERS_META.get('facet_options', {}).get('region', []))} regions, "
          f"{len(_PAPERS_META.get('facet_options', {}).get('top_authors', []))} top authors")


def _paper_match(p, region=None, doc_type=None, author=None, topic=None,
                 year_bucket=None, method=None, problem=None, search=None):
    """Check if paper matches all active filters (AND logic)."""
    if region and p.get('region') != region:
        return False
    if doc_type and doc_type not in p.get('doc_types', []):
        return False
    if author and author.lower() not in (p.get('lead_author') or '').lower():
        return False
    if topic and topic not in p.get('topics', []):
        return False
    if year_bucket and p.get('year_bucket') != year_bucket:
        return False
    if method and p.get('method') != method:
        return False
    if problem:
        # match by linked_questions OR concept tag
        lq = p.get('linked_questions') or []
        if problem not in lq:
            return False
    if search:
        s = search.lower()
        hay = (p.get('id', '') + ' ' + p.get('descriptor', '') + ' '
               + (p.get('lead_author') or '') + ' ' + ' '.join(p.get('topics', []))).lower()
        if s not in hay:
            return False
    return True


@app.get('/api/papers/list')
async def papers_list(region: str = '', doc_type: str = '', author: str = '',
                      topic: str = '', year_bucket: str = '', method: str = '',
                      problem: str = '', search: str = '', group_by: str = ''):
    """Filter + (optional) group papers. group_by ∈ region/method/year_bucket/topic/doc_type/lead_author/problem."""
    if not _PAPERS_META:
        return JSONResponse({'error': 'Papers metadata not loaded'}, status_code=500)

    items = []
    for p in _PAPERS_META['entries']:
        if _paper_match(p, region or None, doc_type or None, author or None,
                        topic or None, year_bucket or None, method or None,
                        problem or None, search or None):
            items.append({
                'id': p['id'],
                'descriptor': p['descriptor'],
                'lead_author': p.get('lead_author', ''),
                'year': p.get('year'),
                'year_bucket': p.get('year_bucket', ''),
                'region': p.get('region', ''),
                'is_vn': p.get('is_vn', False),
                'doc_types': p.get('doc_types', []),
                'method': p.get('method', ''),
                'topics': p.get('topics', []),
                'linked_questions': p.get('linked_questions', []),
                'doi': p.get('doi'),
                'pdf_folder': p.get('pdf_folder'),
                'files': p.get('files', {}),
            })

    # Sort default: VN first, then by year desc
    items.sort(key=lambda x: (not x['is_vn'], -(x['year'] or 0), x['id']))

    grouped = None
    if group_by:
        groups = {}
        for it in items:
            key = None
            if group_by == 'region':
                key = it['region']
            elif group_by == 'method':
                key = it['method']
            elif group_by == 'year_bucket':
                key = it['year_bucket']
            elif group_by == 'doc_type':
                # Use first doc_type or "Khác"
                key = it['doc_types'][0] if it['doc_types'] else 'Khác'
            elif group_by == 'lead_author':
                key = it['lead_author'] or 'Khác'
            elif group_by == 'topic':
                # Each paper can be in multiple topic groups
                for t in (it['topics'] or ['Khác']):
                    groups.setdefault(t, []).append(it)
                continue
            elif group_by == 'problem':
                lqs = it['linked_questions'] or ['Chưa có Q&A']
                for q in lqs:
                    groups.setdefault(q, []).append(it)
                continue
            else:
                key = 'Tất cả'
            groups.setdefault(key or 'Khác', []).append(it)
        grouped = [
            {'group': g, 'count': len(papers), 'items': papers}
            for g, papers in sorted(groups.items(), key=lambda x: -len(x[1]))
        ]

    return JSONResponse({
        'total': len(items),
        'group_by': group_by or None,
        'groups': grouped,
        'items': items if not group_by else None,
    })


@app.get('/api/papers/facets')
async def papers_facets():
    """Return available facet options + counts."""
    if not _PAPERS_META:
        return JSONResponse({'error': 'Papers metadata not loaded'}, status_code=500)
    return JSONResponse({
        'meta': _PAPERS_META.get('meta', {}),
        'facet_options': _PAPERS_META.get('facet_options', {}),
    })


@app.get('/api/papers/detail/{paper_id}')
async def paper_detail(paper_id: str):
    if not _PAPERS_META:
        return JSONResponse({'error': 'not loaded'}, status_code=500)
    for p in _PAPERS_META['entries']:
        if p['id'] == paper_id:
            # Resolve linked Q&A details
            linked_q_full = []
            if _QUESTIONS_INDEX:
                qmap = {q['id']: q for q in _QUESTIONS_INDEX['entries']}
                for qid in p.get('linked_questions', []):
                    q = qmap.get(qid)
                    if q:
                        linked_q_full.append({
                            'id': q['id'],
                            'question': q['question'],
                            'topic': q['topic'],
                            'doc_path': q.get('doc_path', ''),
                        })
            return JSONResponse({
                **p,
                'linked_questions_full': linked_q_full,
            })
    return JSONResponse({'error': 'not found'}, status_code=404)


# =========================================================
# GLOSSARY
# =========================================================
_GLOSSARY = None

def _load_glossary():
    global _GLOSSARY
    try:
        p = DATA_DIR / 'glossary.json'
        if p.exists():
            with open(p, encoding='utf-8') as f:
                _GLOSSARY = json.load(f)
            return True
    except Exception as e:
        print(f'Glossary load failed: {e}')
    return False

_load_glossary()
if _GLOSSARY:
    print(f'Glossary loaded: {len(_GLOSSARY.get("abbreviations", {}))} abbrev + '
          f'{len(_GLOSSARY.get("detailed_terms", {}))} detailed + '
          f'{len(_GLOSSARY.get("organizations", {}))} orgs')


@app.get('/api/glossary')
async def glossary_all():
    if not _GLOSSARY:
        return JSONResponse({'error': 'not loaded'}, status_code=500)
    return JSONResponse(_GLOSSARY)


@app.get('/api/glossary/search')
async def glossary_search(q: str = ''):
    if not _GLOSSARY:
        return JSONResponse({'error': 'not loaded'}, status_code=500)
    q = q.strip()
    if not q:
        return JSONResponse({'results': []})
    ql = q.lower()
    results = {'abbreviations': {}, 'detailed_terms': {}, 'organizations': {}}
    for key, val in _GLOSSARY['abbreviations'].items():
        if ql == key.lower() or ql in key.lower() or ql in val.lower():
            results['abbreviations'][key] = val
    for key, val in _GLOSSARY.get('detailed_terms', {}).items():
        haystack = (key + ' ' + val.get('name_vn', '') + ' ' +
                    val.get('explanation_vn', '') + ' ' + val.get('category', '')).lower()
        if ql in haystack:
            results['detailed_terms'][key] = val
    for key, val in _GLOSSARY.get('organizations', {}).items():
        haystack = (key + ' ' + val.get('name_full', '') + ' ' +
                    val.get('name_vn', '') + ' ' + val.get('role', '')).lower()
        if ql in haystack:
            results['organizations'][key] = val
    return JSONResponse({
        'query': q,
        'total_matches': (len(results['abbreviations']) + len(results['detailed_terms'])
                          + len(results['organizations'])),
        'results': results,
    })


@app.get('/api/glossary/categories')
async def glossary_categories():
    if not _GLOSSARY:
        return JSONResponse({'error': 'not loaded'}, status_code=500)
    cat_ct = {}
    for key, val in _GLOSSARY.get('detailed_terms', {}).items():
        c = val.get('category', 'Other')
        cat_ct[c] = cat_ct.get(c, 0) + 1
    return JSONResponse({'categories': cat_ct})


# =========================================================
# Static + Index
# =========================================================
app.mount('/static', StaticFiles(directory=str(STATIC_DIR)), name='static')
if DOCS_DIR.exists():
    app.mount('/docs', StaticFiles(directory=str(DOCS_DIR)), name='docs')


@app.get('/', response_class=HTMLResponse)
async def index():
    idx = STATIC_DIR / 'index.html'
    with open(idx, 'r', encoding='utf-8') as f:
        return f.read()


if __name__ == '__main__':
    import uvicorn
    port = int(os.environ.get('PORT', 8000))
    uvicorn.run(app, host='0.0.0.0', port=port)
