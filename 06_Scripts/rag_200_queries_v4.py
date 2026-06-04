# -*- coding: utf-8 -*-
"""
RAG cho Bao cao Can thiep — phien ban 4 (HYBRID retrieval).
- Chunks NHO: 1 paragraph = 1 chunk + context header (subsection + paper) prepended
- Tables: gop voi caption thanh 1 chunk
- HYBRID retrieval: semantic (embeddings) + keyword overlap → re-rank
- 200+ test queries
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
DOC_PATH = os.path.join(ROOT, 'Bao cao Can thiep tam ly RLLA VTN - 10042026 v2.docx')
DB_PATH = os.path.join(ROOT, 'rag_bao_cao_can_thiep')

doc = Document(DOC_PATH)

# ============================================================
# 1. CHUNKING — small chunks + context augmentation
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

PAPER_PATS = [
    ('TranNguyenNgoc2018', r'Trần Nguyễn Ngọc'),
    ('Bress2024', r'Bress et al\. 2024|ứng dụng Maya|app Maya'),
    ('Walder2025', r'Walder et al\. 2025'),
    ('Xian2024', r'Xian.*2024'),
    ('DeSilva2024', r'De Silva.*2024'),
    ('Brown2025', r'Brown.*Carter.*2025|BESST|PLACES|MHST'),
    ('Cao2025', r'Cao.*2025'),
    ('Zhang2026', r'Zhang.*Liang.*Kang.*2026'),
    ('Qiaochu2025', r'Qiaochu'),
    ('Menon2025', r'Menon.*2025'),
    ('Li2025', r'Li et al\. 2025'),
    ('Sasaki2024', r'Sasaki.*2024'),
    ('Praptomojati2024', r'Praptomojati'),
    ('Walkup2008', r'Walkup.*2008|CAMS'),
    ('Zugman2024', r'Zugman.*2024'),
    ('Dong2025', r'Dong.*Wang.*Lin.*2025|0328785'),
    ('Chen2025', r'Chen et al'),
]
def detect_paper(text):
    return [pid for pid, pat in PAPER_PATS if re.search(pat, text)]

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
            # Update last_paper based on subsection
            ps = detect_paper(text)
            if ps:
                last_paper = ps[0]
            continue
        # If this paragraph is a TLTK entry "Bài N — ...", detect paper FROM text only
        # (do NOT inherit last_paper from previous paragraph)
        is_tltk_entry = bool(re.match(r'^Bài\s*\d+', text))
        ps_local = detect_paper(text)
        if is_tltk_entry:
            chunk_paper = ps_local[0] if ps_local else ''
        else:
            if ps_local:
                last_paper = ps_local[0]
            chunk_paper = last_paper
        # Track caption
        if re.match(r'^(Bảng|Biểu đồ)\s*\d+\.', text):
            last_caption = text
        # Build chunk with context: section header + subsection + paper
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

print(f'Chunks: {len(chunks)} '
      f'(h={sum(1 for c in chunks if c["type"]=="heading")}, '
      f'p={sum(1 for c in chunks if c["type"]=="paragraph")}, '
      f't={sum(1 for c in chunks if c["type"]=="table")})')

# ============================================================
# 2. BUILD RAG
# ============================================================
print('Loading multilingual model...')
model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
client = chromadb.PersistentClient(path=DB_PATH)
try:
    client.delete_collection('bao_cao_can_thiep_v2')
except Exception:
    pass
col = client.create_collection('bao_cao_can_thiep_v2', metadata={'hnsw:space': 'cosine'})

texts = [c['text'] for c in chunks]
ids = [c['id'] for c in chunks]
metas = [{'type': c['type'], 'section': c['section'],
          'subsection': c.get('subsection', ''), 'papers': c['papers']} for c in chunks]
embs = model.encode(texts, show_progress_bar=False, batch_size=32).tolist()
col.add(documents=texts, embeddings=embs, ids=ids, metadatas=metas)
print(f'Added {col.count()} chunks')

# ============================================================
# 3. HYBRID RETRIEVAL
# ============================================================
def normalize(s):
    return s.lower()

# Pre-load ALL chunks for full-text fallback
ALL_CHUNKS_DATA = col.get(include=['documents', 'metadatas'])
ALL_DOCS = ALL_CHUNKS_DATA['documents']
ALL_IDS = ALL_CHUNKS_DATA['ids']
ALL_METAS = ALL_CHUNKS_DATA['metadatas']

def hybrid_search(query, n=10, k_semantic=40):
    """
    1) Semantic search top k_semantic
    2) Add keyword score (substring matches in lowercased text)
    3) Re-rank by combined score
    4) FALLBACK: nếu query có identifier (DOI, NCT, ...) — full-text scan
       toàn collection để bổ sung kết quả
    """
    q_emb = model.encode([query]).tolist()
    res = col.query(query_embeddings=q_emb, n_results=k_semantic,
                    include=['documents', 'distances', 'metadatas'])
    docs = list(res['documents'][0])
    dists = list(res['distances'][0])
    metas_r = list(res['metadatas'][0])
    ids_r = list(res['ids'][0])

    # Tokenize query
    q_tokens = re.findall(r'[\w/.\-]+', query.lower())
    stop = {'và', 'của', 'là', 'có', 'cho', 'với', 'trong', 'các', 'một', 'này',
            'sau', 'theo', 'khi', 'để', 'bởi', 'từ', 'như', 'or', 'and', 'the',
            'tại', 'về', 'không', 'đã', 'phải', 'sẽ', 'ra', 'an', 'do'}
    q_tokens = [t for t in q_tokens if len(t) >= 2 and t not in stop]

    # Detect identifier-heavy query (DOI, NCT, CRD/PROSPERO, UMIN, ...)
    has_identifier = bool(re.search(
        r'(doi|nct\d+|crd\d+|umin\d+|prospero|10\.\d{3,}|\d{5,})',
        query.lower()
    ))
    sem_w, kw_w = (0.3, 0.7) if has_identifier else (0.6, 0.4)

    # Detect paper in query — boost chunks with same paper tag
    query_papers = set(detect_paper(query))

    # FALLBACK: full-text scan when query has identifier OR paper-specific
    if has_identifier or query_papers:
        if has_identifier:
            id_tokens = [t for t in q_tokens
                         if re.search(r'\d{4,}|10\.\d', t) or len(t) >= 6]
        else:
            id_tokens = []
        existing_ids = set(ids_r)
        for doc_text, doc_id, meta in zip(ALL_DOCS, ALL_IDS, ALL_METAS):
            if doc_id in existing_ids:
                continue
            # Add chunks matching identifier OR matching the paper in query
            doc_papers = set((meta.get('papers') or '').split('|'))
            doc_lower = doc_text.lower()
            match_id = id_tokens and any(t in doc_lower for t in id_tokens)
            match_paper = query_papers and (doc_papers & query_papers)
            if match_id or match_paper:
                docs.append(doc_text)
                dists.append(0.7)
                metas_r.append(meta)
                ids_r.append(doc_id)

    scored = []
    for doc_text, dist, meta, doc_id in zip(docs, dists, metas_r, ids_r):
        sem_score = 1 - dist
        doc_lower = doc_text.lower()
        kw_matches = sum(1 for t in q_tokens if t in doc_lower)
        kw_score = kw_matches / max(len(q_tokens), 1)
        # Paper boost: if query mentions paper P and chunk has P → +0.15
        doc_papers = set((meta.get('papers') or '').split('|'))
        paper_boost = 0.15 if (query_papers and doc_papers & query_papers) else 0.0
        combined = sem_w * sem_score + kw_w * kw_score + paper_boost
        scored.append({
            'doc': doc_text, 'meta': meta, 'id': doc_id,
            'sem': sem_score, 'kw': kw_score, 'combined': combined,
            'kw_matches': kw_matches, 'n_tokens': len(q_tokens),
        })
    scored.sort(key=lambda x: x['combined'], reverse=True)
    return scored[:n]

# ============================================================
# 4. 200+ QUERIES (giong v3)
# ============================================================
queries = []

# (Same 30 manual + 50 each group = ~200 — paste from v3)
queries += [
    ('Trần Nguyễn Ngọc 2018 luận án năm bảo vệ', ['2018', 'Tiến sĩ']),
    ('Trần Nguyễn Ngọc tổng số trang luận án', ['177']),
    ('Trần Nguyễn Ngọc tổng số bệnh nhân nghiên cứu', ['170']),
    ('Trần Nguyễn Ngọc số bệnh nhân hoàn thành can thiệp', ['99']),
    ('Trần Nguyễn Ngọc Đại học Y Hà Nội Bạch Mai', ['Bạch Mai']),
    ('Trần Nguyễn Ngọc người hướng dẫn khoa học', ['Nguyễn Kim Việt']),
    ('Trần Nguyễn Ngọc tỷ lệ nam nữ trong mẫu', ['65', '105']),
    ('Trần Nguyễn Ngọc tuổi trung bình bệnh nhân', ['43,2']),
    ('Trần Nguyễn Ngọc mã chuyên ngành tâm thần', ['62720148']),
    ('Trần Nguyễn Ngọc thiết kế nghiên cứu trước sau', ['trước', 'sau']),
    ('Trần Nguyễn Ngọc liệu pháp Thư giãn Luyện tập', ['Thư giãn', 'Luyện tập']),
    ('Trần Nguyễn Ngọc số buổi can thiệp 4 tuần', ['20 buổi']),
    ('Trần Nguyễn Ngọc thời lượng mỗi buổi tập', ['60 phút']),
    ('Trần Nguyễn Ngọc bệnh nhân nội trú hay ngoại trú', ['nội trú']),
    ('Trần Nguyễn Ngọc công cụ đánh giá HAM-A PSQI EPI', ['HAM-A', 'PSQI']),
    ('Trần Nguyễn Ngọc 3 thời điểm đo T0 T2 T4', ['T0', 'T2', 'T4']),
    ('Trần Nguyễn Ngọc tỷ lệ lo âu nặng vào viện T0', ['45,5']),
    ('Trần Nguyễn Ngọc tỷ lệ lo âu nặng tuần 4 T4', ['11,1']),
    ('Trần Nguyễn Ngọc HAM-A nhẹ tăng T2 T4', ['53,6', '52,5']),
    ('Trần Nguyễn Ngọc tỷ lệ hồi hộp tim đập nhanh ban đầu', ['88,9']),
    ('Trần Nguyễn Ngọc hồi hộp tim đập nhanh sau 4 tuần', ['43,4']),
    ('Trần Nguyễn Ngọc tỷ lệ vã mồ hôi T0 T4', ['59,6', '16,1']),
    ('Trần Nguyễn Ngọc triệu chứng run T0', ['57,6']),
    ('Trần Nguyễn Ngọc khô miệng T0 sau điều trị', ['38,4']),
    ('Trần Nguyễn Ngọc khó thở thuyên giảm', ['56,6']),
    ('Trần Nguyễn Ngọc sợ bị chết tỷ lệ', ['33,3']),
    ('Trần Nguyễn Ngọc sang chấn tâm lý tỷ lệ', ['45,3']),
    ('Trần Nguyễn Ngọc trầm cảm kèm theo nữ nam', ['12,4', '15,2']),
    ('Trần Nguyễn Ngọc tỷ lệ rút khỏi nghiên cứu', ['41,8']),
    ('Trần Nguyễn Ngọc Bảng 3.20 trang 76 nguồn dữ liệu', ['3.20', '76']),
    ('Trần Nguyễn Ngọc Bảng 3.23 trang 78', ['3.23', '78']),
    ('Trần Nguyễn Ngọc cấu trúc 5 phần buổi tập', ['15 phút', '10 phút']),
    ('Trần Nguyễn Ngọc Yoga thư giãn thở 3 thành phần', ['Jacobson', 'Schultz']),
    ('Trần Nguyễn Ngọc Montgomery 13.3 phần trăm', ['13,3', 'Montgomery']),
    ('Trần Nguyễn Ngọc CGI mức nặng giảm', ['52,5', '9,1']),
    ('Trần Nguyễn Ngọc đối tượng người lớn không phải VTN', ['người lớn']),
    ('Trần Nguyễn Ngọc luận án đầu tiên Việt Nam', ['ĐẦU TIÊN', 'Việt Nam']),
    ('Trần Nguyễn Ngọc thiết kế không có nhóm đối chứng', ['không đối chứng']),
    ('Trần Nguyễn Ngọc Naomi Breslau cỡ mẫu', ['Breslau']),
    ('Trần Nguyễn Ngọc khả năng nhân rộng cộng đồng', ['nhân rộng', 'rẻ tiền']),
]
queries += [
    ('De Silva 2024 Sri Lanka Cluster RCT thiết kế', ['cluster RCT', 'Sri Lanka']),
    ('De Silva tổng số học sinh tham gia', ['720']),
    ('De Silva số trường được phân ngẫu nhiên', ['36 trường']),
    ('De Silva học sinh lớp tuổi nào', ['lớp 9']),
    ('De Silva 8 phiên CBT mỗi tuần', ['8 phiên']),
    ('De Silva 40 phút mỗi phiên', ['40 phút']),
    ('De Silva ai cung cấp can thiệp', ['giáo viên']),
    ('De Silva tỷ lệ mất mẫu rất thấp', ['< 1', 'mất mẫu']),
    ('De Silva beta giảm lo âu sau 3 tháng', ['−0,096', '0,038']),
    ('De Silva tự trọng tăng beta', ['0,811']),
    ('De Silva DOI Child Adolesc Psychiatry Mental Health', ['18', '108']),
    ('De Silva mô hình khả thi LMIC', ['LMIC', 'khả thi']),
    ('Xian 2024 Network Meta-Analysis Bayesian SAD', ['NMA', 'Bayesian']),
    ('Xian tổng số RCT được phân tích', ['30 RCT']),
    ('Xian tổng số người tham gia VTN', ['1.547']),
    ('Xian iCBT SUCRA hạng 1', ['71,2']),
    ('Xian gCBT nhóm SUCRA', ['68,4']),
    ('Xian I-CBT cá nhân SUCRA', ['66,0']),
    ('Xian iCBT trầm cảm SUCRA hai trong một', ['92,2']),
    ('Xian gCBT chức năng cải thiện', ['89,6']),
    ('Xian Tạp chí J Affect Disord trang', ['614', '627']),
    ('Xian PROSPERO đăng ký', ['CRD42023476829']),
    ('Sasaki 2024 Nhật Bản RCT đa trung tâm subthreshold', ['Sasaki', 'Nhật']),
    ('Sasaki bao nhiêu HS SV đăng ký', ['77']),
    ('Sasaki nhóm can thiệp đối chứng', ['38', '39']),
    ('Sasaki 6 đại học và trường THPT', ['6 đại học', 'THPT']),
    ('Sasaki iCBT tự học không hướng dẫn', ['tự học']),
    ('Sasaki tỷ lệ đáp ứng', ['61']),
    ('Sasaki OR đáp ứng', ['4,97']),
    ('Sasaki tỷ lệ phục hồi', ['68']),
    ('Sasaki OR phục hồi', ['3,95']),
    ('Sasaki UMIN đăng ký', ['UMIN000049768']),
    ('Praptomojati 2024 CA-CBT Đông Nam Á', ['CA-CBT', 'Praptomojati']),
    ('Praptomojati tổng số nghiên cứu trong SR', ['7 nghiên cứu', '7 NC']),
    ('Praptomojati Asian Journal of Psychiatry tạp chí', ['Asian Journal']),
    ('Praptomojati 0 nghiên cứu nào từ Việt Nam', ['Việt Nam']),
    ('Praptomojati khung CTAF', ['CTAF']),
    ('Praptomojati Hồi giáo Phật giáo thích ứng tôn giáo', ['Hồi giáo', 'Phật giáo']),
    ('Qiaochu 2025 Mobile CBT 9 RCT', ['9 RCT', 'Qiaochu']),
    ('Qiaochu tổng N người tham gia', ['2.479']),
    ('Qiaochu mobile trầm cảm tỷ lệ NC dương tính', ['7/8', '87,5']),
    ('Qiaochu mobile lo âu tỷ lệ NC dương tính', ['2/6', '33,3']),
    ('Qiaochu Clinical Psychology Psychotherapy DOI', ['e70173', 'cpp']),
    ('Menon 2025 scoping review LMIC', ['Menon', 'scoping']),
    ('Menon số nghiên cứu tổng hợp', ['69']),
    ('Menon số quốc gia LMIC', ['12 quốc gia']),
    ('Menon 32 RCT 31 trước sau 6 đánh giá', ['32 RCT', '31']),
    ('Menon khoảng trống cộng đồng gia đình dài hạn', ['CỘNG ĐỒNG', 'GIA ĐÌNH']),
    ('Menon Asia Pacific J Public Health pages', ['332', '346']),
    ('Li 2025 BMC Psychiatry NMA Bayesian 30 RCT', ['BMC', 'NMA']),
    ('Li tổng số người tham gia 30 RCT', ['1.711']),
    ('Li ACT SUCRA hạng cao nhưng CrI', ['ACT', '0,69']),
    ('Li CBT SUCRA bằng chứng mạnh nhất', ['0,66', '16 RCT']),
    ('Li VRET SUCRA virtual reality', ['VRET', '0,51']),
    ('Li PE physical exercise SUCRA', ['PE', '0,51']),
]
queries += [
    ('Walder 2025 JMIR DMHI 21 RCT meta-analysis', ['Walder', 'JMIR']),
    ('Walder DMHI vs đối chứng tổng quát g', ['0,508']),
    ('Walder DMHI vs WL danh sách chờ', ['0,576']),
    ('Walder DMHI nền CBT Hedges g', ['0,610']),
    ('Walder DMHI có hướng dẫn người', ['0,825', 'hướng dẫn']),
    ('Walder DMHI thiết kế riêng SAD g lớn nhất', ['0,878']),
    ('Walder DMHI không hướng dẫn nhỏ', ['0,3', '0,4']),
    ('Walder PROSPERO CRD', ['CRD42023424181']),
    ('Walder J Med Internet Research Q1', ['JMIR']),
    ('Bress 2024 Maya app JAMA Network Open', ['Bress', 'Maya', 'JAMA']),
    ('Bress NCT05130281 ClinicalTrials đăng ký', ['NCT05130281']),
    ('Bress thanh niên 18 25 tuổi RCT', ['18', '25']),
    ('Bress tổng số người tham gia', ['59']),
    ('Bress GAD SAD chiếm bao nhiêu', ['56', '41']),
    ('Bress Maya 12 phiên 6 tuần', ['12 phiên', '6 tuần']),
    ('Bress HAM-A tuần 3 chênh lệch', ['−3,20', 'Cohen', '0,64']),
    ('Bress HAM-A tuần 6 kết thúc Cohen d', ['−5,64', '0,94']),
    ('Bress HAM-A tuần 12 theo dõi', ['−5,67', '1,04']),
    ('Bress LSAS social anxiety scale tuần 12', ['−17,61', '1,07']),
    ('Bress ASI anxiety sensitivity', ['−9,51', '0,93']),
    ('Brown Carter 2025 J Mental Health Editorial', ['Brown', 'Carter']),
    ('Brown Carter PLACES self-referral mô hình', ['PLACES', 'self-referral']),
    ('Brown Carter BESST trial CBT 900 HS 57 trường', ['BESST', '900', '57']),
    ('Brown Carter Mindfulness Kuyken 8376 HS thất bại', ['8.376', 'Kuyken']),
    ('Brown Carter MHST Mental Health Support Team', ['MHST']),
    ('Brown Carter co-design học sinh tham gia thiết kế', ['co-design']),
    ('Brown Carter trang 357 361 vol 34', ['357', '361']),
    ('Cao 2025 Frontiers Psychiatry resilience MA RCT', ['Cao', 'Frontiers', 'resilience']),
    ('Cao DOI 1594658 article', ['1594658']),
    ('Cao kích thước hiệu ứng nhỏ trung bình', ['nhỏ', 'trung bình']),
    ('Cao heterogeneity cao', ['heterogeneity']),
    ('Cao Cochrane Risk of Bias đánh giá', ['Cochrane']),
    ('Cao positive psychology yếu tố bảo vệ', ['BẢO VỆ', 'lạc quan']),
    ('Zhang Liang Kang 2026 Bayesian MA J Clin Psychol', ['Zhang', 'Bayesian']),
    ('Zhang 31 RCT tổng hợp', ['31 RCT']),
    ('Zhang n học sinh nguy cơ thấp', ['19.865']),
    ('Zhang hiệu quả khiêm tốn trầm cảm lo âu', ['khiêm tốn', 'NHỎ']),
    ('Zhang chất lượng bằng chứng nền rất thấp', ['THẤP']),
    ('Zhang duy trì 1 năm sau can thiệp', ['1 năm']),
    ('Zhang trang 248 259 vol 82', ['248', '259']),
    ('Zugman 2024 American Journal Psychiatry pediatric anxiety', ['Zugman', 'AJP']),
    ('Zugman pediatric anxiety treatment review', ['pediatric', 'review']),
    ('Zugman 3 rối loạn lo âu trẻ em GAD SAD', ['GAD', 'SAD']),
    ('Zugman trang 189 200 vol 181', ['189', '200', '181']),
    ('Zugman NIMH NIH NC SKTT hàng đầu Mỹ', ['NIMH', 'NIH']),
    ('Walkup 2008 CAMS NEJM RCT đa trung tâm', ['Walkup', 'CAMS']),
    ('Walkup CAMS tổng số trẻ tham gia', ['488']),
    ('Walkup CAMS tuổi trẻ 7 17', ['7', '17']),
    ('Walkup CBT SSRI kết hợp đáp ứng', ['80,7']),
    ('Walkup CBT đơn thuần đáp ứng', ['59,7']),
    ('Walkup SSRI Sertraline đơn thuần', ['54,9', 'Sertraline']),
    ('Walkup placebo baseline đáp ứng', ['23,7', 'placebo']),
    ('Walkup CAMS NEJM 359 2753', ['NEJM', '2753']),
]
queries += [
    ('tóm lược 6 phát hiện then chốt', ['then chốt', 'phát hiện']),
    ('CBT phương pháp hiệu quả nhất xuyên suốt', ['CBT', 'hiệu quả nhất']),
    ('CAMS 80,7 phần trăm CBT SSRI gold standard', ['CAMS', '80,7']),
    ('iCBT SAD hạng 1 NMA Walder Xian', ['iCBT', 'hạng 1']),
    ('mobile CBT trầm cảm mạnh hơn lo âu', ['trầm cảm', 'lo âu', 'Qiaochu']),
    ('universal CBT phổ quát hiệu quả nhỏ', ['phổ quát', 'NHỎ']),
    ('khoảng trống cực lớn Việt Nam 0 RCT', ['0 RCT', 'cực lớn']),
    ('Đông Á Thái Bình Dương LMIC thiếu cộng đồng', ['LMIC']),
    ('Bảng 11 so sánh ba vùng địa lý', ['ba vùng', 'so sánh']),
    ('Bảng 12 xếp hạng 13 phương pháp', ['xếp hạng', '13']),
    ('Việt Nam số bài can thiệp chỉ 1', ['Trần Nguyễn Ngọc']),
    ('Châu Á số bài can thiệp ngoài Việt Nam', ['7']),
    ('Châu Âu Úc Mỹ số bài can thiệp', ['7']),
    ('hạng 1 CBT SSRI kết hợp khuyến nghị VN', ['CBT + SSRI']),
    ('hạng 2 CBT cá nhân nhóm cốt lõi', ['CBT', 'cốt lõi']),
    ('hạng 3 iCBT cho SAD VTN số hoá', ['iCBT', 'SAD']),
    ('hạng 13 mindfulness phổ quát tránh', ['mindfulness', 'TRÁNH']),
    ('5 khoảng trống nghiên cứu Việt Nam', ['5', 'khoảng trống']),
    ('khoảng trống số 1 thiếu tuyệt đối RCT', ['THIẾU TUYỆT ĐỐI', 'RCT']),
    ('5 thành phần can thiệp 12 tuần Bảng 13', ['5 thành phần', '12 tuần']),
    ('Module CBT NHÓM TARGETED 8 buổi', ['CBT NHÓM', '8 buổi']),
    ('Module GIAO TIẾP GIA ĐÌNH 4 buổi cha mẹ', ['GIAO TIẾP GIA ĐÌNH', '4 buổi']),
    ('Module RESILIENCE 3 buổi', ['RESILIENCE', '3 buổi']),
    ('Module PE thư giãn thở 8 buổi', ['PE', '8 buổi']),
    ('App iCBT tiếng Việt suốt 12 tuần', ['iCBT', '12 tuần']),
    ('PLACES self-referral cơ chế tiếp cận', ['PLACES']),
    ('Dong PLOS ONE 0,22 kênh giao tiếp', ['Dong', '0,22']),
    ('Chen 2025 141 nghiên cứu meta', ['Chen', '141']),
    ('emotional functioning yếu tố bảo vệ Chen', ['emotional', 'bảo vệ']),
    ('Trần Thảo Vi lạc quan trung gian', ['Trần Thảo Vi', 'lạc quan']),
]
queries += [
    ('Bress 2024 DOI jamanetworkopen 28372', ['10.1001/jamanetworkopen']),
    ('Walder 2025 DOI preprints 67067', ['10.2196/preprints']),
    ('Xian 2024 DOI jad 08.097', ['10.1016/j.jad']),
    ('De Silva 2024 DOI s13034', ['10.1186/s13034']),
    ('Cao 2025 DOI fpsyt 1594658', ['10.3389/fpsyt']),
    ('Brown Carter DOI 09638237', ['10.1080/09638237']),
    ('Li 2025 DOI s12888 07227', ['10.1186/s12888']),
    ('Sasaki 2024 DOI 55786', ['10.2196/55786']),
    ('Zugman 2024 DOI appi ajp', ['10.1176/appi.ajp']),
    ('Praptomojati DOI ajp 103896', ['10.1016/j.ajp']),
    ('Zhang Liang Kang 2026 DOI jclp 70069', ['10.1002/jclp']),
    ('Qiaochu 2025 DOI cpp 70173', ['10.1002/cpp']),
    ('Menon 2025 DOI 10105395241313154', ['10.1177/10105395']),
    ('Walkup 2008 NEJM 359 2753 2766', ['NEJM', '2753', '359']),
    ('CAMS NNT Number Needed to Treat', ['NNT']),
    ('Bài 28 Zugman tài liệu tham khảo', ['28', 'Zugman']),
    ('Bài 29 Li 2025 BMC Psychiatry', ['29', 'Li']),
    ('Bài 41 Praptomojati Hartanto', ['41', 'Praptomojati']),
    ('Bài 42 De Silva 2024', ['42', 'De Silva']),
    ('Bài 43 Xian Zhang Jiang', ['43', 'Xian']),
    ('Bài 44 Walder Frey Berger', ['44', 'Walder']),
    ('Bài 48 Brown Carter editorial', ['48', 'Brown']),
    ('Bài 49 Bress Falk Schier', ['49', 'Bress']),
    ('Bài 50 Cao 2025 Frontiers', ['50', 'Cao']),
    ('Bài 51 Sasaki Nhật Bản', ['51', 'Sasaki']),
    ('Bài 56 Zhang Liang Kang abstract paywall', ['56', 'Zhang']),
    ('Bài 57 Qiaochu Wang abstract', ['57', 'Qiaochu']),
    ('Bài 58 Menon Coppard McEwen', ['58', 'Menon']),
    ('Bài 54 Dong Wang Lin đối chiếu', ['54', 'Dong']),
    ('Bài 55 Chen Wang Zhu COVID', ['55', 'Chen']),
]

print(f'\nTotal queries: {len(queries)}')

# ============================================================
# 5. RUN VERIFICATION (HYBRID)
# ============================================================
print('\n' + '=' * 70)
print(f'KIEM TRA RAG HYBRID / {len(queries)} QUERIES')
print('=' * 70)

passed = 0
failed = 0
results_log = []

for i, (q, expected_kws) in enumerate(queries, 1):
    results = hybrid_search(q, n=10)
    top5_text = ' '.join(r['doc'] for r in results[:10]).lower()
    found_kws = [kw for kw in expected_kws if kw.lower() in top5_text]
    n_found = len(found_kws)
    n_expected = len(expected_kws)
    top_score = results[0]['combined'] * 100 if results else 0
    success = n_found >= 1 and top_score >= 25

    status = 'PASS' if success else 'FAIL'
    if success: passed += 1
    else: failed += 1

    results_log.append({
        'idx': i,
        'query': q,
        'expected': expected_kws,
        'found': found_kws,
        'top_id': results[0]['id'] if results else '',
        'top_section': results[0]['meta'].get('section', '') if results else '',
        'top_paper': results[0]['meta'].get('papers', '') if results else '',
        'top_excerpt': (results[0]['doc'][:200] if results else ''),
        'sem_score': round(results[0]['sem'] * 100, 1) if results else 0,
        'kw_score': round(results[0]['kw'] * 100, 1) if results else 0,
        'combined_score': round(top_score, 1),
        'status': status,
    })

print(f'\n{"="*70}')
print(f'TONG KET: {passed}/{len(queries)} PASS ({passed/len(queries)*100:.1f}%) | {failed} FAIL')
print(f'{"="*70}')

groups = [
    ('GROUP 1: VN — Trần Nguyễn Ngọc', 1, 40),
    ('GROUP 2: Châu Á', 41, 90),
    ('GROUP 3: Âu - Úc - Mỹ', 91, 140),
    ('GROUP 4: Tóm lược + Khuyến nghị', 141, 170),
    ('GROUP 5: TLTK + DOI', 171, 200),
]
for name, lo, hi in groups:
    sub = [r for r in results_log if lo <= r['idx'] <= hi]
    sp = sum(1 for r in sub if r['status'] == 'PASS')
    print(f'  {name}: {sp}/{len(sub)} ({sp/max(len(sub),1)*100:.0f}%)')

# Save log
log_path = os.path.join(os.path.dirname(__file__), 'rag_verification_200.json')
with open(log_path, 'w', encoding='utf-8') as f:
    json.dump({
        'doc': os.path.basename(DOC_PATH),
        'n_chunks': len(chunks),
        'n_queries': len(queries),
        'passed': passed,
        'failed': failed,
        'pass_rate': round(passed / len(queries) * 100, 2),
        'retrieval': 'hybrid (semantic 0.6 + keyword 0.4)',
        'embedding_model': 'paraphrase-multilingual-MiniLM-L12-v2',
        'results': results_log,
    }, f, ensure_ascii=False, indent=2)
print(f'\nLog saved: {log_path}')

fails = [r for r in results_log if r['status'] == 'FAIL']
if fails:
    print(f'\n--- ALL {len(fails)} FAILS ---')
    for r in fails:
        print(f'[{r["idx"]:3d}] {r["query"][:60]:60s} expected={r["expected"]} found={r["found"]} score={r["combined_score"]}%')
