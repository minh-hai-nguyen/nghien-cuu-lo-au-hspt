# -*- coding: utf-8 -*-
"""
Tạo RAG cho Báo cáo Can thiệp Tâm lý RLLA VTN — phiên bản 2.
Đồng thời dùng RAG làm công cụ KIỂM TRA tính nhất quán nội bộ của báo cáo:
1) Trích heading + đoạn văn (gộp consecutive trong cùng section) + bảng (mỗi bảng là 1 chunk lớn)
2) Encode bằng paraphrase-multilingual-MiniLM-L12-v2 (đa ngôn ngữ — tối ưu cho Việt-Anh)
3) Lưu vào ChromaDB persistent
4) Chạy 30 truy vấn kiểm tra → đối chiếu với kỳ vọng (key facts)
5) Tổng kết PASS/FAIL → phát hiện lỗi inconsistency / thiếu thông tin
"""
import sys, os, io, re, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn
import chromadb
from sentence_transformers import SentenceTransformer

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
DOC_PATH = os.path.join(ROOT, 'Bao cao Can thiep tam ly RLLA VTN - 10042026 v2.docx')
DB_PATH = os.path.join(ROOT, 'rag_bao_cao_can_thiep')

print(f'Đọc báo cáo: {DOC_PATH}')
doc = Document(DOC_PATH)

# ============================================================
# 1. CHUNK STRATEGY
# ============================================================
chunks = []  # list of dicts: text, chunk_id, section, type, source_paper

current_section = 'TIÊU ĐỀ'
current_subsection = ''

# Heading mapping (manually mapped from the doc structure)
section_map = {
    'BÁO CÁO TỔNG HỢP': 'Tiêu đề',
    'TÓM LƯỢC ĐIỀU HÀNH': '0. Tóm lược',
    'PHẦN I — VIỆT NAM': 'I. Việt Nam',
    'PHẦN II — CHÂU Á': 'II. Châu Á',
    'PHẦN III — CHÂU ÂU': 'III. Châu Âu - Úc - Mỹ',
    'PHẦN IV — TỔNG KẾT': 'IV. Tổng kết',
    'TÀI LIỆU THAM KHẢO': 'V. Tài liệu tham khảo',
}

def detect_section(text):
    for key, val in section_map.items():
        if text.startswith(key):
            return val
    return None

def detect_paper(text):
    """Heuristic: identify which paper(s) this chunk references"""
    papers = []
    patterns = [
        ('TranNguyenNgoc2018', r'Trần Nguyễn Ngọc'),
        ('Bress2024', r'Bress et al\. 2024|Maya'),
        ('Walder2025', r'Walder et al\. 2025|DMHI'),
        ('Xian2024', r'Xian.*2024|NMA SAD'),
        ('DeSilva2024', r'De Silva.*2024|Sri Lanka'),
        ('Brown2025', r'Brown.*Carter.*2025|BESST|PLACES'),
        ('Cao2025', r'Cao.*2025|resilience.*MA'),
        ('Zhang2026', r'Zhang.*Liang.*Kang.*2026|Bayesian MA'),
        ('Qiaochu2025', r'Qiaochu.*2025|Mobile CBT'),
        ('Menon2025', r'Menon.*2025|Scoping.*LMIC'),
        ('Li2025', r'Li et al\. 2025|BMC.*NMA'),
        ('Sasaki2024', r'Sasaki.*2024|iCBT.*Nhật'),
        ('Praptomojati2024', r'Praptomojati|CA-CBT'),
        ('Walkup2008', r'Walkup.*2008|CAMS'),
        ('Zugman2024', r'Zugman.*2024|AJP.*pediatric'),
        ('Dong2025', r'Dong.*Wang.*Lin.*2025|PLOS ONE.*0328785'),
        ('Chen2025', r'Chen et al\. \(?2025'),
    ]
    for pid, pat in patterns:
        if re.search(pat, text):
            papers.append(pid)
    return papers

# 1.1 CHUNKING: walk through document body in order — mix paragraphs + tables.
# Strategy: HYBRID
#  - Mỗi paragraph riêng = 1 chunk (đảm bảo retrieval chính xác từng câu)
#  - PREPEND context header (Section + Subsection + Last paper mentioned) vào đầu mỗi chunk
#  - Mỗi bảng = 1 chunk lớn, gộp với caption đứng trước nó
#  - Heading riêng cũng thành chunk (giúp tìm theo tên section)
# Walk in document order using XML body
from docx.oxml.ns import qn as _qn

current_section = 'Tiêu đề'
current_subsection = ''
last_paper_in_text = ''
chunk_idx = 0
last_caption = None  # text của caption bảng (đoạn ngay trước bảng)

# Helper: walk body elements in order
body = doc.element.body

def make_context(section, subsection, paper):
    parts = [f'[{section}]']
    if subsection and subsection != section:
        parts.append(f'[{subsection}]')
    if paper:
        parts.append(f'[{paper}]')
    return ' '.join(parts)

# Get list of paragraph elements and table elements in order
para_elems = list(body.iter(_qn('w:p')))
# Build map of paragraph element -> docx Paragraph object
from docx.text.paragraph import Paragraph
from docx.table import Table

# Iterate body's direct children to maintain order
for child in body.iterchildren():
    if child.tag == _qn('w:p'):
        para = Paragraph(child, doc)
        text = para.text.strip()
        if not text:
            continue
        style = para.style.name
        sec = detect_section(text)
        if sec:
            current_section = sec
            current_subsection = text
            chunk_idx += 1
            chunks.append({
                'text': text,
                'id': f'h{chunk_idx}',
                'type': 'heading',
                'section': current_section,
                'subsection': current_subsection,
                'papers': '',
            })
            last_caption = None
            continue
        if style.startswith('Heading'):
            current_subsection = text
            chunk_idx += 1
            chunks.append({
                'text': text,
                'id': f'h{chunk_idx}',
                'type': 'heading',
                'section': current_section,
                'subsection': current_subsection,
                'papers': '',
            })
            last_caption = None
            continue
        # Detect papers in this paragraph
        papers_here = detect_paper(text)
        if papers_here:
            last_paper_in_text = papers_here[0]
        # Build augmented chunk with context prepended
        ctx = make_context(current_section, current_subsection, last_paper_in_text)
        aug_text = f'{ctx}\n{text}'
        chunk_idx += 1
        chunks.append({
            'text': aug_text,
            'id': f'p{chunk_idx}',
            'type': 'paragraph',
            'section': current_section,
            'subsection': current_subsection,
            'papers': '|'.join(papers_here) if papers_here else last_paper_in_text,
        })
        # Caption tracking — italic + size 10 paragraphs that start with "Bảng N." or "Biểu đồ N."
        if re.match(r'^(Bảng|Biểu đồ)\s*\d+\.', text):
            last_caption = text
        else:
            last_caption = None
    elif child.tag == _qn('w:tbl'):
        tb = Table(child, doc)
        if len(tb.rows) < 2:
            continue
        headers = [c.text.strip() for c in tb.rows[0].cells]
        parts = []
        # Include caption if exists
        if last_caption:
            parts.append(f'CAPTION: {last_caption}')
        parts.append('HEADERS: ' + ' | '.join(headers))
        for ri, row in enumerate(tb.rows[1:], start=1):
            cells = [c.text.strip() for c in row.cells]
            pairs = [f'{h}={v}' for h, v in zip(headers, cells) if v]
            if pairs:
                parts.append(' | '.join(pairs))
        full = '\n'.join(parts)
        ctx = make_context(current_section, current_subsection, last_paper_in_text)
        aug_text = f'{ctx}\n{full}'
        chunk_idx += 1
        chunks.append({
            'text': aug_text,
            'id': f't{chunk_idx}',
            'type': 'table',
            'section': current_section,
            'subsection': current_subsection,
            'papers': '|'.join(detect_paper(full)) or last_paper_in_text,
        })
        last_caption = None

print(f'Tổng số chunks: {len(chunks)}')
print(f'  - Headings: {sum(1 for c in chunks if c["type"] == "heading")}')
print(f'  - Paragraphs: {sum(1 for c in chunks if c["type"] == "paragraph")}')
print(f'  - Tables: {sum(1 for c in chunks if c["type"] == "table")}')
total_chars = sum(len(c['text']) for c in chunks)
print(f'  - Tổng chars: {total_chars:,}')
print(f'  - TB chars/chunk: {total_chars // len(chunks)}')

# ============================================================
# 2. EMBEDDING + STORE
# ============================================================
print('\nLoading embedding model paraphrase-multilingual-MiniLM-L12-v2 (đa ngôn ngữ)...')
model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')

print(f'Tạo ChromaDB persistent tại: {DB_PATH}')
client = chromadb.PersistentClient(path=DB_PATH)
# Reset collection
try:
    client.delete_collection('bao_cao_can_thiep_v2')
except Exception:
    pass

col = client.create_collection(
    'bao_cao_can_thiep_v2',
    metadata={'hnsw:space': 'cosine', 'description': 'Bao cao Can thiep tam ly RLLA VTN v2 - 10042026'}
)

texts = [c['text'] for c in chunks]
ids = [c['id'] for c in chunks]
metas = [{'type': c['type'], 'section': c['section'],
          'subsection': c.get('subsection', ''), 'papers': c['papers']} for c in chunks]

print(f'Encoding {len(texts)} chunks...')
embeddings = model.encode(texts, show_progress_bar=False, batch_size=32).tolist()
col.add(documents=texts, embeddings=embeddings, ids=ids, metadatas=metas)
print(f'Đã add {col.count()} chunks vào collection.')

# ============================================================
# 3. VERIFICATION QUERIES — kiểm tra báo cáo có chứa đúng key facts
# ============================================================
print('\n' + '='*70)
print('KIỂM TRA RAG / ĐỐI CHIẾU NỘI BỘ BÁO CÁO')
print('='*70)

# Mỗi truy vấn có expected_keywords — phải xuất hiện trong top-3 results
test_queries = [
    # Việt Nam
    ('Trần Nguyễn Ngọc luận án 2018 đối tượng', ['170', 'người lớn', 'Bạch Mai']),
    ('hiệu quả thư giãn luyện tập sau 4 tuần', ['11,1', 'nặng', '45,5']),
    ('cỡ mẫu can thiệp Trần Nguyễn Ngọc', ['99', 'hoàn thành']),
    ('lo âu Việt Nam có RCT cho VTN không', ['0 RCT', 'khoảng trống', 'không có']),
    ('triệu chứng tim đập nhanh giảm sau điều trị', ['88,9', '43,4', 'p < 0,0001']),

    # Châu Á
    ('Sri Lanka cluster RCT trường', ['720', 'De Silva', '36 trường']),
    ('iCBT hạng 1 lo âu xã hội VTN', ['71,2', 'Xian', 'SUCRA']),
    ('CBT do giáo viên LMIC', ['Sri Lanka', 'giáo viên', 'khả thi']),
    ('Nhật Bản subthreshold SAD iCBT', ['77', 'Sasaki', '4,97']),
    ('CA-CBT thích ứng văn hoá Đông Nam Á', ['Praptomojati', '7 NC', 'Việt Nam']),
    ('mobile CBT trầm cảm vs lo âu', ['7/8', '2/6', 'Qiaochu']),
    ('khoảng trống cộng đồng gia đình LMIC', ['Menon', '69', '12 quốc gia']),
    ('NMA Bayesian SUCRA CBT trẻ em Li', ['0,66', '16 RCT', 'BMC']),

    # Châu Âu - Mỹ
    ('CAMS đáp ứng CBT SSRI', ['80,7', 'Walkup', '488']),
    ('Walder DMHI Hedges g SAD specific', ['0,878', 'Walder', 'JMIR']),
    ('Maya app HAM-A Cohen d 12 tuần', ['1,04', 'Bress', 'JAMA']),
    ('mindfiness UK 8376 học sinh thất bại', ['8.376', 'Brown', 'mindfulness']),
    ('Bayesian MA universal CBT 19865', ['19.865', 'Zhang', '31 RCT']),
    ('resilience trường meta-analysis', ['Cao', 'Frontiers', 'resilience']),
    ('Zugman pediatric anxiety treatment', ['Zugman', 'AJP', '189']),

    # Tổng hợp / khuyến nghị
    ('xếp hạng can thiệp tâm lý cho VN', ['CBT', 'iCBT', 'gCBT']),
    ('5 thành phần can thiệp đề cương VN', ['12 tuần', 'CBT NHÓM', 'GIA ĐÌNH']),
    ('PLACES self-referral mô hình UK', ['PLACES', 'BESST', 'self-referral']),
    ('khoảng trống số 1 Việt Nam', ['0 RCT', 'VTN', 'khoảng trống']),
    ('Dong PLOS kênh giao tiếp gia đình OR', ['0,22', 'Dong', '78']),

    # Citation accuracy checks
    ('JAMA Network Open Bress NCT05130281', ['NCT05130281', 'Bress', 'JAMA']),
    ('Walkup 2008 NEJM placebo response rate', ['23,7', 'placebo', 'Walkup']),
    ('Trần Nguyễn Ngọc Bảng 3.20 trang 76', ['3.20', '76', 'HAM-A']),
    ('De Silva 2024 cluster RCT page reference', ['18.108', 'De Silva', '2024']),
    ('Xian 2024 J Affect Disord pages', ['614', '627', 'Xian']),
]

passed = 0
failed = 0
results_log = []

for q, expected_kws in test_queries:
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
    success = n_found >= max(1, n_expected - 1) and rel >= 30  # cho phép thiếu 1 keyword

    status = 'PASS' if success else 'FAIL'
    if success:
        passed += 1
    else:
        failed += 1
    print(f'[{status}] [{rel:.0f}%] {q[:50]:50s} {n_found}/{n_expected} kw → {found_kws}')

    results_log.append({
        'query': q,
        'expected': expected_kws,
        'found': found_kws,
        'top_doc': docs[0][:200] if docs else '',
        'top_section': metas_r[0].get('section', '') if metas_r else '',
        'top_paper': metas_r[0].get('papers', '') if metas_r else '',
        'relevance': round(rel, 1),
        'status': status,
    })

print('\n' + '='*70)
print(f'KẾT QUẢ KIỂM TRA: {passed}/{len(test_queries)} PASS '
      f'({passed/len(test_queries)*100:.0f}%) | {failed} FAIL')
print('='*70)

# Lưu log JSON
log_path = os.path.join(os.path.dirname(__file__), 'rag_verification_log.json')
with open(log_path, 'w', encoding='utf-8') as f:
    json.dump({
        'doc': os.path.basename(DOC_PATH),
        'n_chunks': len(chunks),
        'n_queries': len(test_queries),
        'passed': passed,
        'failed': failed,
        'pass_rate': round(passed / len(test_queries) * 100, 1),
        'results': results_log,
    }, f, ensure_ascii=False, indent=2)
print(f'\nLog đã lưu: {log_path}')

# Hiển thị FAIL chi tiết
if failed > 0:
    print('\n--- FAIL DETAILS ---')
    for r in results_log:
        if r['status'] == 'FAIL':
            print(f'\n? {r["query"]}')
            print(f'  Expected: {r["expected"]}')
            print(f'  Found: {r["found"]}')
            print(f'  Top doc ({r["relevance"]}%): {r["top_doc"]}')

# Database stats
db_size = sum(os.path.getsize(os.path.join(dp, f))
              for dp, dn, fn in os.walk(DB_PATH) for f in fn) // 1024
print(f'\nKích thước RAG DB: {db_size} KB')
print(f'Đường dẫn: {DB_PATH}')
