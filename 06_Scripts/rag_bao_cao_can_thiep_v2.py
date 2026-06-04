# -*- coding: utf-8 -*-
"""
RAG cho Bao cao Can thiep Tam ly RLLA VTN — phien ban 2 (200 queries).
- Chunking: paragraph_group (gop consecutive paragraphs trong cung subsection)
- Tables: gop voi caption + headers + rows thanh 1 chunk lon
- Embedding: paraphrase-multilingual-MiniLM-L12-v2
- 200 test queries: tu sinh tu noi dung bao cao
"""
import sys, os, io, re, json, random
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn
import chromadb
from sentence_transformers import SentenceTransformer

random.seed(42)

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
DOC_PATH = os.path.join(ROOT, 'Bao cao Can thiep tam ly RLLA VTN - 10042026 v2.docx')
DB_PATH = os.path.join(ROOT, 'rag_bao_cao_can_thiep')

print(f'Doc reader: {os.path.basename(DOC_PATH)}')
doc = Document(DOC_PATH)

# ============================================================
# 1. CHUNKING — paragraph_group + caption-merged tables
# ============================================================
chunks = []
section_map = {
    'BÁO CÁO TỔNG HỢP': 'Tieu de',
    'TÓM LƯỢC ĐIỀU HÀNH': '0. Tom luoc',
    'PHẦN I — VIỆT NAM': 'I. Viet Nam',
    'PHẦN II — CHÂU Á': 'II. Chau A',
    'PHẦN III — CHÂU ÂU': 'III. Chau Au - Uc - My',
    'PHẦN IV — TỔNG KẾT': 'IV. Tong ket',
    'TÀI LIỆU THAM KHẢO': 'V. Tai lieu tham khao',
}
def detect_section(text):
    for k, v in section_map.items():
        if text.startswith(k):
            return v
    return None

def detect_paper(text):
    pats = [
        ('TranNguyenNgoc2018', r'Trần Nguyễn Ngọc'),
        ('Bress2024', r'Bress et al\. 2024|Maya'),
        ('Walder2025', r'Walder et al\. 2025|DMHI'),
        ('Xian2024', r'Xian.*2024|NMA SAD'),
        ('DeSilva2024', r'De Silva.*2024|Sri Lanka'),
        ('Brown2025', r'Brown.*Carter.*2025|BESST|PLACES'),
        ('Cao2025', r'Cao.*2025|resilience.*MA'),
        ('Zhang2026', r'Zhang.*Liang.*Kang.*2026'),
        ('Qiaochu2025', r'Qiaochu.*2025|Mobile CBT'),
        ('Menon2025', r'Menon.*2025|Scoping.*LMIC'),
        ('Li2025', r'Li et al\. 2025|BMC.*NMA'),
        ('Sasaki2024', r'Sasaki.*2024'),
        ('Praptomojati2024', r'Praptomojati|CA-CBT'),
        ('Walkup2008', r'Walkup.*2008|CAMS'),
        ('Zugman2024', r'Zugman.*2024'),
        ('Dong2025', r'Dong.*Wang.*Lin.*2025|0328785'),
        ('Chen2025', r'Chen et al'),
    ]
    out = []
    for pid, pat in pats:
        if re.search(pat, text):
            out.append(pid)
    return out

current_section = 'Tieu de'
current_subsection = ''
last_paper = ''
chunk_idx = 0
last_caption = None
group_buffer = []  # accumulate paragraphs in current subsection

def flush_group():
    global chunk_idx, group_buffer
    if not group_buffer:
        return
    chunk_idx += 1
    full = '\n'.join(group_buffer)
    chunks.append({
        'text': full,
        'id': f'g{chunk_idx}',
        'type': 'paragraph_group',
        'section': current_section,
        'subsection': current_subsection,
        'papers': '|'.join(detect_paper(full)),
    })
    group_buffer = []

body = doc.element.body
for child in body.iterchildren():
    if child.tag == qn('w:p'):
        para = Paragraph(child, doc)
        text = para.text.strip()
        if not text:
            continue
        style = para.style.name
        sec = detect_section(text)
        if sec:
            flush_group()
            current_section = sec
            current_subsection = text
            chunk_idx += 1
            chunks.append({
                'text': text, 'id': f'h{chunk_idx}', 'type': 'heading',
                'section': current_section, 'subsection': text, 'papers': '',
            })
            last_caption = None
            continue
        if style.startswith('Heading'):
            flush_group()
            current_subsection = text
            chunk_idx += 1
            chunks.append({
                'text': text, 'id': f'h{chunk_idx}', 'type': 'heading',
                'section': current_section, 'subsection': text, 'papers': '',
            })
            last_caption = None
            continue
        # If this paragraph is a CAPTION (Bảng X. or Biểu đồ X.) — track it
        if re.match(r'^(Bảng|Biểu đồ)\s*\d+\.', text):
            last_caption = text
        else:
            last_caption = None
        group_buffer.append(text)
    elif child.tag == qn('w:tbl'):
        # Flush paragraph group first (caption is in buffer)
        # But keep caption in chunk for the table
        tb = Table(child, doc)
        if len(tb.rows) < 2:
            continue
        headers = [c.text.strip() for c in tb.rows[0].cells]
        parts = []
        # If caption is the last item in group_buffer, use it (but DON'T remove)
        if last_caption:
            parts.append(f'CAPTION: {last_caption}')
        elif group_buffer and re.match(r'^(Bảng|Biểu đồ)\s*\d+\.', group_buffer[-1]):
            parts.append(f'CAPTION: {group_buffer[-1]}')
        parts.append('HEADERS: ' + ' | '.join(headers))
        for ri, row in enumerate(tb.rows[1:], start=1):
            cells = [c.text.strip() for c in row.cells]
            pairs = [f'{h}={v}' for h, v in zip(headers, cells) if v]
            if pairs:
                parts.append(' | '.join(pairs))
        full = '\n'.join(parts)
        ctx_section = f'[{current_section}] [{current_subsection}]'
        full_ctx = f'{ctx_section}\n{full}'
        chunk_idx += 1
        chunks.append({
            'text': full_ctx, 'id': f't{chunk_idx}', 'type': 'table',
            'section': current_section, 'subsection': current_subsection,
            'papers': '|'.join(detect_paper(full)),
        })
        last_caption = None

flush_group()

print(f'Total chunks: {len(chunks)}')
print(f'  - Headings: {sum(1 for c in chunks if c["type"] == "heading")}')
print(f'  - Paragraph groups: {sum(1 for c in chunks if c["type"] == "paragraph_group")}')
print(f'  - Tables: {sum(1 for c in chunks if c["type"] == "table")}')
total_chars = sum(len(c['text']) for c in chunks)
print(f'  - Total chars: {total_chars:,}')
print(f'  - Avg chars/chunk: {total_chars // len(chunks)}')

# ============================================================
# 2. EMBEDDING + STORE
# ============================================================
print('\nLoading multilingual model...')
model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')

print(f'ChromaDB at: {DB_PATH}')
client = chromadb.PersistentClient(path=DB_PATH)
try:
    client.delete_collection('bao_cao_can_thiep_v2')
except Exception:
    pass

col = client.create_collection(
    'bao_cao_can_thiep_v2',
    metadata={'hnsw:space': 'cosine', 'description': 'Bao cao Can thiep tam ly RLLA VTN v2'}
)

texts = [c['text'] for c in chunks]
ids = [c['id'] for c in chunks]
metas = [{'type': c['type'], 'section': c['section'],
          'subsection': c.get('subsection', ''), 'papers': c['papers']} for c in chunks]

print(f'Encoding {len(texts)} chunks...')
embs = model.encode(texts, show_progress_bar=False, batch_size=32).tolist()
col.add(documents=texts, embeddings=embs, ids=ids, metadatas=metas)
print(f'Added: {col.count()} chunks')

# ============================================================
# 3. AUTO-GENERATE 200 TEST QUERIES from the report content
# ============================================================
print('\nGenerating 200 test queries from report content...')

test_queries = []  # list of (query, expected_keywords)

# A) Manual high-value queries (30) — like before, but tweaked
manual = [
    ('Trần Nguyễn Ngọc luận án 2018 đối tượng người lớn', ['170', 'Bạch Mai']),
    ('hiệu quả thư giãn luyện tập sau 4 tuần HAM-A', ['11,1', '45,5']),
    ('cỡ mẫu can thiệp Trần Nguyễn Ngọc bao nhiêu', ['99', 'hoàn thành']),
    ('Việt Nam có RCT cho VTN không', ['0 RCT', 'khoảng trống']),
    ('triệu chứng tim đập nhanh giảm sau điều trị', ['88,9', '43,4']),
    ('Sri Lanka cluster RCT 720 học sinh', ['720', 'De Silva']),
    ('iCBT hạng 1 lo âu xã hội VTN SUCRA', ['71,2', 'Xian']),
    ('CBT do giáo viên LMIC khả thi', ['Sri Lanka', 'giáo viên']),
    ('Nhật Bản subthreshold SAD iCBT Sasaki', ['77', 'Sasaki']),
    ('CA-CBT thích ứng văn hoá Đông Nam Á 7 nghiên cứu', ['Praptomojati', 'Việt Nam']),
    ('mobile CBT 9 RCT trầm cảm vs lo âu', ['7/8', 'Qiaochu']),
    ('khoảng trống cộng đồng gia đình LMIC Menon', ['Menon', '69']),
    ('NMA Bayesian SUCRA CBT trẻ em Li', ['0,66', '16 RCT']),
    ('CAMS Walkup 2008 đáp ứng kết hợp CBT SSRI', ['80,7', 'Walkup']),
    ('Walder DMHI Hedges g SAD specific', ['0,878', 'Walder']),
    ('Maya app HAM-A Cohen d 12 tuần', ['1,04', 'Bress']),
    ('mindfulness UK 8376 học sinh thất bại', ['8.376', 'Brown']),
    ('Bayesian MA universal CBT 19865', ['19.865', 'Zhang']),
    ('resilience trường meta-analysis Cao', ['Cao', 'Frontiers']),
    ('Zugman pediatric anxiety treatment AJP', ['Zugman', 'AJP']),
    ('xếp hạng 13 phương pháp can thiệp', ['CBT', 'iCBT']),
    ('5 thành phần can thiệp đề cương Việt Nam 12 tuần', ['12 tuần', 'CBT NHÓM']),
    ('PLACES self-referral mô hình UK', ['PLACES', 'BESST']),
    ('khoảng trống số 1 Việt Nam', ['0 RCT', 'VTN']),
    ('Dong PLOS kênh giao tiếp gia đình OR', ['0,22', 'Dong']),
    ('JAMA Network Open Bress NCT05130281', ['NCT05130281', 'Bress']),
    ('Walkup 2008 NEJM placebo response rate', ['23,7', 'placebo']),
    ('Trần Nguyễn Ngọc Bảng 3.20', ['3.20', 'HAM-A']),
    ('De Silva 2024 cluster RCT page reference', ['18, 108', 'De Silva']),
    ('Xian 2024 J Affect Disord pages', ['614', 'Xian']),
]
test_queries.extend(manual)

# B) Auto-generate from each paragraph_group containing statistics
# Pattern: find sentences with numeric statistics
all_text_chunks = [c for c in chunks if c['type'] in ('paragraph_group', 'table')]

stat_pat = re.compile(r'(\d+[,\.]?\d*\s*%|p\s*[<>=]\s*0[,\.]\d+|OR\s*=\s*\d+[,\.]?\d*|β\s*=\s*[−\-]?\d+[,\.]?\d*|n\s*=\s*\d+|SUCRA\s*\d+[,\.]?\d*|d\s*=\s*\d+[,\.]?\d*|g\s*=\s*\d+[,\.]?\d*)')

# Auto queries — extract a stat + nearest 5-6 keywords as query
import unicodedata
def shorten(text, n):
    return text[:n] + ('...' if len(text) > n else '')

auto_queries = []
seen = set()
for c in all_text_chunks:
    txt = c['text']
    # Get sentences
    sents = re.split(r'(?<=[\.!?])\s+', txt)
    for s in sents:
        s = s.strip()
        if len(s) < 30 or len(s) > 400:
            continue
        m = stat_pat.search(s)
        if not m:
            continue
        stat = m.group(0).strip()
        # Build a query from the surrounding context (remove the stat to make it a question)
        # Strategy: use 6-10 keyword tokens from the sentence
        tokens = re.findall(r'\b[\w%/áàảãạâầấẩẫậăằắẳẵặéèẻẽẹêềếểễệíìỉĩịóòỏõọôồốổỗộơờớởỡợúùủũụưừứửữựýỳỷỹỵđÁÀẢÃẠÂẦẤẨẪẬĂẰẮẲẴẶÉÈẺẼẸÊỀẾỂỄỆÍÌỈĨỊÓÒỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÚÙỦŨỤƯỪỨỬỮỰÝỲỶỸỴĐ]+\b', s)
        if len(tokens) < 6:
            continue
        # Take first 8-10 meaningful tokens (skip pure numbers)
        kw_tokens = [t for t in tokens if not re.fullmatch(r'\d+[,\.]?\d*%?', t)][:10]
        if len(kw_tokens) < 4:
            continue
        query = ' '.join(kw_tokens[:8])
        # The expected keyword is the stat
        # Also add 1-2 distinctive content words from the sentence
        content_words = [t for t in kw_tokens if len(t) > 5][:2]
        expected = [stat] + content_words[:1]
        key = query[:60]
        if key in seen:
            continue
        seen.add(key)
        auto_queries.append((query, expected))
        if len(auto_queries) >= 170:  # cap at 170 to combine with 30 manual = 200
            break
    if len(auto_queries) >= 170:
        break

test_queries.extend(auto_queries[:170])
print(f'Total queries: {len(test_queries)} ({len(manual)} manual + {len(auto_queries[:170])} auto)')

# ============================================================
# 4. RUN VERIFICATION
# ============================================================
print('\n' + '=' * 70)
print(f'KIEM TRA RAG / 200 QUERIES')
print('=' * 70)

passed = 0
failed = 0
results_log = []

for i, (q, expected_kws) in enumerate(test_queries, 1):
    q_emb = model.encode([q]).tolist()
    res = col.query(
        query_embeddings=q_emb,
        n_results=5,
        include=['documents', 'distances', 'metadatas'],
    )
    docs = res['documents'][0]
    dists = res['distances'][0]
    metas_r = res['metadatas'][0]
    top_text = ' '.join(docs).lower()
    found_kws = [kw for kw in expected_kws if kw.lower() in top_text]
    n_found = len(found_kws)
    n_expected = len(expected_kws)
    rel = max(0, (1 - dists[0]) * 100)
    success = n_found >= max(1, n_expected - 1) and rel >= 30

    status = 'PASS' if success else 'FAIL'
    if success:
        passed += 1
    else:
        failed += 1

    if i <= 30 or status == 'FAIL':  # In first 30 + all FAILs
        prefix = 'M' if i <= 30 else 'A'
        print(f'[{status}] {prefix}{i:03d} [{rel:.0f}%] {q[:55]:55s} {n_found}/{n_expected}')

    results_log.append({
        'idx': i,
        'kind': 'manual' if i <= 30 else 'auto',
        'query': q,
        'expected': expected_kws,
        'found': found_kws,
        'top_doc_id': res['ids'][0][0] if 'ids' in res else '',
        'top_doc_excerpt': docs[0][:200] if docs else '',
        'top_section': metas_r[0].get('section', '') if metas_r else '',
        'top_paper': metas_r[0].get('papers', '') if metas_r else '',
        'relevance': round(rel, 1),
        'status': status,
    })

print()
print('=' * 70)
print(f'KET QUA: {passed}/{len(test_queries)} PASS ({passed/len(test_queries)*100:.1f}%) | {failed} FAIL')
print('=' * 70)

# Stats by kind
m_pass = sum(1 for r in results_log if r['kind'] == 'manual' and r['status'] == 'PASS')
a_pass = sum(1 for r in results_log if r['kind'] == 'auto' and r['status'] == 'PASS')
m_total = sum(1 for r in results_log if r['kind'] == 'manual')
a_total = sum(1 for r in results_log if r['kind'] == 'auto')
print(f'  Manual:  {m_pass}/{m_total} ({m_pass/max(m_total,1)*100:.0f}%)')
print(f'  Auto:    {a_pass}/{a_total} ({a_pass/max(a_total,1)*100:.0f}%)')

# Save log
log_path = os.path.join(os.path.dirname(__file__), 'rag_verification_200.json')
with open(log_path, 'w', encoding='utf-8') as f:
    json.dump({
        'doc': os.path.basename(DOC_PATH),
        'n_chunks': len(chunks),
        'n_queries': len(test_queries),
        'passed': passed,
        'failed': failed,
        'pass_rate': round(passed / len(test_queries) * 100, 1),
        'manual_pass': m_pass,
        'manual_total': m_total,
        'auto_pass': a_pass,
        'auto_total': a_total,
        'results': results_log,
    }, f, ensure_ascii=False, indent=2)
print(f'\nLog saved: {log_path}')

# Show first 5 FAIL details
fails = [r for r in results_log if r['status'] == 'FAIL']
if fails:
    print(f'\n--- FIRST 5 FAILS (out of {len(fails)}) ---')
    for r in fails[:5]:
        print(f'\n[{r["idx"]}] {r["query"][:80]}')
        print(f'  Expected: {r["expected"]}')
        print(f'  Found: {r["found"]}')
        print(f'  Top section: {r["top_section"]}')
        print(f'  Top excerpt: {r["top_doc_excerpt"][:150]}')

db_size = sum(os.path.getsize(os.path.join(dp, f))
              for dp, dn, fn in os.walk(DB_PATH) for f in fn) // 1024
print(f'\nDB size: {db_size} KB')
