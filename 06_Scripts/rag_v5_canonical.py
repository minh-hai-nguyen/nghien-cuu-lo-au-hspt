# -*- coding: utf-8 -*-
"""
RAG v5 — bao cao can thiep v3 (11/04/2026) voi canonical IDs.

Cai tien so voi v4:
- Doc tu 01_Bao-cao/Bao cao Can thiep tam ly RLLA VTN - 11042026 v3.docx
- Paper detection tra ve canonical ID (VN030, QT040, ...) thay vi author+year
- Support nhieu bai moi tim duoc qua web (Happy House, CoolMinds, ...)
- Them metadata 'canonical_id' cho tung chunk
"""
import sys, os, io, re, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn
import chromadb
from sentence_transformers import SentenceTransformer

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
DOC_PATH = os.path.join(ROOT, '01_Bao-cao', 'Bao cao Can thiep tam ly RLLA VTN - 11042026 v3.docx')
DB_PATH = os.path.join(ROOT, 'rag_bao_cao_can_thiep')

print(f'Doc: {os.path.basename(DOC_PATH)}')
doc = Document(DOC_PATH)

# ============================================================
# 1. CHUNKING with canonical paper IDs
# ============================================================
chunks = []
section_map = {
    'BÁO CÁO TỔNG HỢP': 'Tieu de',
    'NHỮNG CẬP NHẬT': '0. Update notes',
    'TÓM LƯỢC ĐIỀU HÀNH': '0. Tom luoc',
    'PHẦN I — VIỆT NAM': 'I. Viet Nam',
    'PHẦN II — CHÂU Á': 'II. Chau A',
    'PHẦN III — CHÂU ÂU': 'III. Chau Au - Uc - My',
    'PHẦN IV — TỔNG KẾT': 'IV. Tong ket',
    'TÀI LIỆU THAM KHẢO': 'V. Tai lieu tham khao',
    'PHỤ LỤC A': 'VI. Phu luc A - Tu khoa',
    'PHỤ LỤC B': 'VII. Phu luc B - Ban do VN',
    'KẾT LUẬN VỀ BẢN V3': 'VIII. Ket luan v3',
}
def detect_section(text):
    for k, v in section_map.items():
        if text.startswith(k):
            return v
    return None

# Canonical paper patterns: regex → VN###/QT### canonical ID
PAPER_PATS = [
    # Vietnam papers
    ('VN001', r'Hoa 2024.*Frontiers|Hoa et al\.? 2024'),
    ('VN002', r'V-NAMHS 2022|V-NAMHS.*khảo sát quốc gia'),
    ('VN003', r'Pham.*2024.*Huế|Phạm.*2024.*hỗ trợ xã hội'),
    ('VN014', r'Hoàng Trung Học.*2025'),
    ('VN015', r'Ngô Anh Vinh.*2024|NgoAnhVinh'),
    ('VN019', r'Hồ Thị Trúc Quỳnh|Trần Thảo Vi 2025.*TLH'),
    ('VN020', r'Trần Hồ Vĩnh Lộc.*2024'),
    ('VN021', r'Trần Thảo Vi 2024|TranThaoVi 2024.*JRuralMed'),
    ('VN022', r'UNICEF VN 2022|UNICEF.*School Factors'),
    ('VN023', r'Nguyen LX 2023|Nguyễn LX 2023.*COVID'),
    ('VN024', r'Nguyễn Thanh Truyền|VinhLong 2024'),
    ('VN025', r'Phạm Thị Ngọc.*2024|Hải Phòng.*Vĩnh Bảo'),
    ('VN026', r'Trần Đức Sĩ.*2025|Long An.*PNT'),
    ('VN027', r'Dinh.*2021.*School.*VN'),
    ('VN028', r'Đào Thị Ngoãn.*2025|TCNCYH 2025.*SV Y'),
    ('VN029', r'Duong.*2025.*TPHCM|Thai.*Duong.*2025'),
    ('VN030', r'Tran.*Nguyen.*Shochet.*2023|Happy House|RAP-?A'),
    # International — CBT / interventions
    ('QT028', r'Zugman.*2024|AJP.*pediatric.*anxiety'),
    ('QT029', r'Li et al\.? 2025.*BMC|BMC.*NMA.*2025'),
    ('QT037', r'Praptomojati|CA-CBT.*ĐNA|CA-CBT.*SEA'),
    ('QT038', r'De Silva.*2024|Sri Lanka.*CBT.*720'),
    ('QT039', r'Xian.*2024|NMA SAD|Xian.*Zhang.*Jiang'),
    ('QT040', r'Walder.*2025|DMHI.*SAD|JMIR.*DMHI'),
    ('QT041', r'Zheng.*2025|MXH.*Stress.*Sleep'),
    ('QT042', r'Brown.*Carter.*2025|BESST|PLACES|MHST|Mindfulness.*Kuyken'),
    ('QT043', r'Bress.*2024|Maya.*app|JAMA.*mobile'),
    ('QT044', r'Cai.*2025.*resilience|Chenyi Cai|Resilience.*MA.*Frontiers'),
    ('QT045', r'Sasaki.*2024|Japan.*iCBT|UMIN'),
    ('QT046', r'Jagiello.*2025|Academic Stress.*SR'),
    ('QT047', r'Dong.*2025.*PLOS|Ya.an.*DASS'),
    ('QT048', r'Chen.*2025.*141|Chen et al\.? 2025.*meta'),
    ('QT049', r'Zhang.*Liang.*Kang.*2026|Bayesian MA.*31 RCT'),
    ('QT050', r'Qiaochu|Mobile CBT.*9 RCT'),
    ('QT051', r'Menon.*2025|Scoping.*LMIC.*East Asia'),
    # Legacy
    ('Walkup2008_CAMS', r'Walkup.*2008|CAMS.*NEJM'),
    ('TranNguyenNgoc2018', r'Trần Nguyễn Ngọc'),
    # New from v3 web search
    ('CoolMinds2025', r'CoolMinds|Studsgaard.*2025'),
    ('SteppedCare2025', r'Baumgart.*2025|ICBT.*SC.*VC|stepped care.*videoconferencing'),
    ('SchoolNurse2025', r'school nurse.*anxiety.*2025|CALM.*Imondi'),
    ('VRET2025', r'VRET.*MA|virtual reality.*anxiety.*2025'),
    ('JAACAP_PA2025', r'JAACAP.*2025.*physical activity|umbrella.*2025.*PA'),
    ('MIXCS_Japan2025', r'MIXCS|Multi-.*Inter-.*Cross-cultural'),
    ('ZapfPCC2024', r'Zapf.*2024|parent-child communication'),
]
def detect_paper(text):
    found = set()
    for pid, pat in PAPER_PATS:
        if re.search(pat, text, re.IGNORECASE):
            found.add(pid)
    return sorted(found)

current_section = 'Tieu de'
current_subsection = ''
last_paper = ''
chunk_idx = 0
last_caption = None

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
            current_section = sec
            current_subsection = text
            chunk_idx += 1
            chunks.append({
                'text': text, 'id': f'h{chunk_idx}', 'type': 'heading',
                'section': current_section, 'subsection': text, 'papers': '',
            })
            last_caption = None
            last_paper = ''
            continue
        if style.startswith('Heading'):
            current_subsection = text
            chunk_idx += 1
            chunks.append({
                'text': text, 'id': f'h{chunk_idx}', 'type': 'heading',
                'section': current_section, 'subsection': text, 'papers': '',
            })
            last_caption = None
            ps = detect_paper(text)
            if ps:
                last_paper = ps[0]
            continue
        # TLTK entry: detect paper ONLY from local text
        is_tltk = bool(re.match(r'^Bài\s*\d+|^\d+\.\s', text))
        ps_local = detect_paper(text)
        if is_tltk:
            chunk_paper = ps_local[0] if ps_local else ''
        else:
            if ps_local:
                last_paper = ps_local[0]
            chunk_paper = last_paper
        if re.match(r'^(Bảng|Biểu đồ)\s*\d+\.', text):
            last_caption = text
        ctx_parts = []
        if current_subsection:
            ctx_parts.append(f'[{current_subsection}]')
        if chunk_paper:
            ctx_parts.append(f'[Paper:{chunk_paper}]')
        ctx_str = ' '.join(ctx_parts)
        full_text = f'{ctx_str}\n{text}' if ctx_str else text
        chunk_idx += 1
        chunks.append({
            'text': full_text, 'id': f'p{chunk_idx}', 'type': 'paragraph',
            'section': current_section, 'subsection': current_subsection,
            'papers': chunk_paper,
        })
    elif child.tag == qn('w:tbl'):
        tb = Table(child, doc)
        if len(tb.rows) < 2:
            continue
        headers = [c.text.strip() for c in tb.rows[0].cells]
        parts = []
        if last_caption:
            parts.append(f'CAPTION: {last_caption}')
        parts.append('HEADERS: ' + ' | '.join(headers))
        for row in tb.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            pairs = [f'{h}={v}' for h, v in zip(headers, cells) if v]
            if pairs:
                parts.append(' | '.join(pairs))
        full = '\n'.join(parts)
        ctx_parts = []
        if current_subsection:
            ctx_parts.append(f'[{current_subsection}]')
        if last_paper:
            ctx_parts.append(f'[Paper:{last_paper}]')
        full_ctx = f'{" ".join(ctx_parts)}\n{full}'
        chunk_idx += 1
        chunks.append({
            'text': full_ctx, 'id': f't{chunk_idx}', 'type': 'table',
            'section': current_section, 'subsection': current_subsection,
            'papers': last_paper or '|'.join(detect_paper(full)),
        })
        last_caption = None

print(f'\nChunks: {len(chunks)} '
      f'(h={sum(1 for c in chunks if c["type"]=="heading")}, '
      f'p={sum(1 for c in chunks if c["type"]=="paragraph")}, '
      f't={sum(1 for c in chunks if c["type"]=="table")})')

# ============================================================
# 2. Build RAG
# ============================================================
print('\nLoading multilingual embedding model...')
model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
client = chromadb.PersistentClient(path=DB_PATH)

# Delete old collection and create new
try:
    client.delete_collection('bao_cao_can_thiep_v3')
except Exception:
    pass
col = client.create_collection(
    'bao_cao_can_thiep_v3',
    metadata={'hnsw:space': 'cosine',
              'description': 'Bao cao Can thiep tam ly RLLA VTN v3 11042026 - canonical IDs'}
)

texts = [c['text'] for c in chunks]
ids = [c['id'] for c in chunks]
metas = [{'type': c['type'], 'section': c['section'],
          'subsection': c.get('subsection', ''), 'papers': c['papers']} for c in chunks]

print(f'Encoding {len(texts)} chunks...')
embs = model.encode(texts, show_progress_bar=False, batch_size=32).tolist()
col.add(documents=texts, embeddings=embs, ids=ids, metadatas=metas)
print(f'Added {col.count()} chunks to collection "bao_cao_can_thiep_v3"')

# Also keep old v2 collection for compatibility — dont delete

# Print per-paper coverage
paper_counts = {}
for c in chunks:
    p = c.get('papers', '')
    if p:
        paper_counts[p] = paper_counts.get(p, 0) + 1

print(f'\nPaper coverage ({len(paper_counts)} unique papers):')
for p, n in sorted(paper_counts.items(), key=lambda x: -x[1])[:20]:
    print(f'  {p}: {n} chunks')

# Save chunks manifest
manifest_path = os.path.join(os.path.dirname(__file__), 'rag_v5_manifest.json')
with open(manifest_path, 'w', encoding='utf-8') as f:
    json.dump({
        'doc': os.path.basename(DOC_PATH),
        'collection': 'bao_cao_can_thiep_v3',
        'n_chunks': len(chunks),
        'embedding_model': 'paraphrase-multilingual-MiniLM-L12-v2',
        'papers': list(paper_counts.keys()),
        'paper_counts': paper_counts,
    }, f, ensure_ascii=False, indent=2)
print(f'\nManifest: {manifest_path}')

# Quick smoke test
print('\n=== SMOKE TEST (5 queries) ===')
test_queries = [
    ('Happy House RCT Việt Nam', ['Happy House', 'Tran', 'RAP-A', '1.084']),
    ('Walder DMHI SAD specific Hedges g', ['0,878', 'Walder']),
    ('Dong giao tiếp gia đình OR 0,22', ['0,22', 'Dong']),
    ('CoolMinds iCBT Đan Mạch', ['CoolMinds', 'Đan Mạch']),
    ('xếp hạng 15 phương pháp can thiệp', ['xếp hạng', '15']),
]
n_pass = 0
for q, expected in test_queries:
    q_emb = model.encode([q]).tolist()
    res = col.query(query_embeddings=q_emb, n_results=5,
                    include=['documents', 'distances'])
    top = ' '.join(res['documents'][0]).lower()
    found = [kw for kw in expected if kw.lower() in top]
    status = 'PASS' if found else 'FAIL'
    if found: n_pass += 1
    rel = (1 - res['distances'][0][0]) * 100
    print(f'  [{status}] [{rel:.0f}%] "{q[:50]}" found={found}')

print(f'\nSmoke test: {n_pass}/{len(test_queries)} PASS')
print(f'\nDB path: {DB_PATH}')
print('To query: chromadb.PersistentClient("rag_bao_cao_can_thiep").get_collection("bao_cao_can_thiep_v3")')
