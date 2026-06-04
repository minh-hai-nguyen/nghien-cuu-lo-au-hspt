# -*- coding: utf-8 -*-
"""
RAG cho Bao cao Can thiep — phien ban 3:
- Chunking giu nguyen (paragraph_group + table)
- 200 queries CHAT LUONG: 30 manual + 170 template-based generated from report content
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
# 1. CHUNKING (giong v2)
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
        ('Xian2024', r'Xian.*2024'),
        ('DeSilva2024', r'De Silva.*2024|Sri Lanka'),
        ('Brown2025', r'Brown.*Carter.*2025|BESST|PLACES'),
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
    return [pid for pid, pat in pats if re.search(pat, text)]

current_section = 'Tieu de'
current_subsection = ''
chunk_idx = 0
last_caption = None
group_buffer = []

def flush_group():
    global chunk_idx, group_buffer
    if not group_buffer:
        return
    chunk_idx += 1
    full = '\n'.join(group_buffer)
    chunks.append({
        'text': full, 'id': f'g{chunk_idx}', 'type': 'paragraph_group',
        'section': current_section, 'subsection': current_subsection,
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
        if re.match(r'^(Bảng|Biểu đồ)\s*\d+\.', text):
            last_caption = text
        group_buffer.append(text)
    elif child.tag == qn('w:tbl'):
        tb = Table(child, doc)
        if len(tb.rows) < 2:
            continue
        headers = [c.text.strip() for c in tb.rows[0].cells]
        parts = []
        cap = last_caption
        if not cap and group_buffer and re.match(r'^(Bảng|Biểu đồ)\s*\d+\.', group_buffer[-1]):
            cap = group_buffer[-1]
        if cap:
            parts.append(f'CAPTION: {cap}')
        parts.append('HEADERS: ' + ' | '.join(headers))
        for row in tb.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            pairs = [f'{h}={v}' for h, v in zip(headers, cells) if v]
            if pairs:
                parts.append(' | '.join(pairs))
        full = '\n'.join(parts)
        ctx = f'[{current_section}] [{current_subsection}]'
        chunk_idx += 1
        chunks.append({
            'text': f'{ctx}\n{full}', 'id': f't{chunk_idx}', 'type': 'table',
            'section': current_section, 'subsection': current_subsection,
            'papers': '|'.join(detect_paper(full)),
        })
        last_caption = None

flush_group()

print(f'Chunks: {len(chunks)} '
      f'(headings={sum(1 for c in chunks if c["type"]=="heading")}, '
      f'paragraph_groups={sum(1 for c in chunks if c["type"]=="paragraph_group")}, '
      f'tables={sum(1 for c in chunks if c["type"]=="table")})')

# ============================================================
# 2. BUILD RAG
# ============================================================
print('Loading model paraphrase-multilingual-MiniLM-L12-v2...')
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
# 3. 200 QUERIES — template-based, manually curated for high quality
# ============================================================
# Each query: (question, expected_keywords[at_least_one_must_match])
# Designed to test SPECIFIC facts that should be retrievable from the report

queries = []

# === GROUP 1: VIỆT NAM — Trần Nguyễn Ngọc 2018 (40 queries) ===
TNN = 'Trần Nguyễn Ngọc 2018'
queries += [
    ('Trần Nguyễn Ngọc 2018 luận án năm bảo vệ', ['2018', 'Tiến sĩ']),
    ('Trần Nguyễn Ngọc tổng số trang luận án', ['177']),
    ('Trần Nguyễn Ngọc tổng số bệnh nhân nghiên cứu', ['170']),
    ('Trần Nguyễn Ngọc số bệnh nhân hoàn thành can thiệp', ['99']),
    ('Trần Nguyễn Ngọc Đại học Y Hà Nội Bạch Mai', ['Bạch Mai', 'Y Hà Nội']),
    ('Trần Nguyễn Ngọc người hướng dẫn khoa học', ['Nguyễn Kim Việt']),
    ('Trần Nguyễn Ngọc tỷ lệ nam nữ trong mẫu', ['65', '105']),
    ('Trần Nguyễn Ngọc tuổi trung bình bệnh nhân', ['43,2']),
    ('Trần Nguyễn Ngọc mã chuyên ngành tâm thần', ['62720148']),
    ('Trần Nguyễn Ngọc thiết kế nghiên cứu trước sau', ['trước', 'sau']),
    ('Trần Nguyễn Ngọc liệu pháp gì không kết hợp thuốc', ['Thư giãn', 'Luyện tập']),
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
    ('Trần Nguyễn Ngọc cấu trúc 5 phần buổi tập', ['15 phút', '10 phút', '20 phút']),
    ('Trần Nguyễn Ngọc Yoga thư giãn thở 3 thành phần', ['Jacobson', 'Schultz']),
    ('Trần Nguyễn Ngọc Montgomery 13.3 phần trăm', ['13,3', 'Montgomery']),
    ('Trần Nguyễn Ngọc CGI mức nặng giảm', ['52,5', '9,1']),
    ('Trần Nguyễn Ngọc đối tượng người lớn không phải VTN', ['người lớn', 'không phải VTN']),
    ('Trần Nguyễn Ngọc luận án đầu tiên Việt Nam', ['ĐẦU TIÊN', 'Việt Nam']),
    ('Trần Nguyễn Ngọc thiết kế không có nhóm đối chứng', ['không đối chứng', 'không có']),
    ('Trần Nguyễn Ngọc Naomi Breslau cỡ mẫu', ['Breslau']),
    ('Trần Nguyễn Ngọc khả năng nhân rộng cộng đồng', ['rẻ tiền', 'nhân rộng']),
]

# === GROUP 2: CHÂU Á — De Silva, Xian, Sasaki, Praptomojati, Qiaochu, Menon, Li (50 queries) ===
queries += [
    # De Silva 2024
    ('De Silva 2024 Sri Lanka Cluster RCT thiết kế', ['cluster RCT', 'Sri Lanka']),
    ('De Silva tổng số học sinh tham gia', ['720']),
    ('De Silva số trường được phân ngẫu nhiên', ['36 trường']),
    ('De Silva học sinh lớp tuổi nào', ['lớp 9', '14']),
    ('De Silva 8 phiên CBT mỗi tuần', ['8 phiên']),
    ('De Silva 40 phút mỗi phiên', ['40 phút']),
    ('De Silva ai cung cấp can thiệp', ['giáo viên']),
    ('De Silva tỷ lệ mất mẫu rất thấp', ['< 1', 'mất mẫu']),
    ('De Silva beta giảm lo âu sau 3 tháng', ['−0,096', '0,038']),
    ('De Silva tự trọng tăng beta', ['0,811']),
    ('De Silva DOI Child Adolesc Psychiatry Mental Health', ['18', '108']),
    ('De Silva mô hình khả thi LMIC', ['LMIC', 'khả thi']),
    # Xian 2024
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
    # Sasaki 2024
    ('Sasaki 2024 Nhật Bản RCT đa trung tâm subthreshold', ['Sasaki', 'Nhật']),
    ('Sasaki bao nhiêu HS SV đăng ký', ['77']),
    ('Sasaki nhóm can thiệp đối chứng', ['38', '39']),
    ('Sasaki 6 đại học và trường THPT', ['6 đại học', 'THPT']),
    ('Sasaki iCBT tự học không hướng dẫn', ['tự học', 'không']),
    ('Sasaki tỷ lệ đáp ứng', ['61']),
    ('Sasaki OR đáp ứng', ['4,97']),
    ('Sasaki tỷ lệ phục hồi', ['68']),
    ('Sasaki OR phục hồi', ['3,95']),
    ('Sasaki UMIN đăng ký', ['UMIN000049768']),
    # Praptomojati 2024
    ('Praptomojati 2024 CA-CBT Đông Nam Á', ['CA-CBT', 'Praptomojati']),
    ('Praptomojati tổng số nghiên cứu trong SR', ['7 nghiên cứu', '7 NC']),
    ('Praptomojati Asian Journal of Psychiatry tạp chí', ['Asian Journal', 'Psychiatry']),
    ('Praptomojati 0 nghiên cứu nào từ Việt Nam', ['Việt Nam', '0']),
    ('Praptomojati khung CTAF', ['CTAF']),
    ('Praptomojati Hồi giáo Phật giáo thích ứng tôn giáo', ['Hồi giáo', 'Phật giáo']),
    # Qiaochu 2025
    ('Qiaochu 2025 Mobile CBT 9 RCT', ['9 RCT', 'Qiaochu']),
    ('Qiaochu tổng N người tham gia', ['2.479']),
    ('Qiaochu mobile trầm cảm tỷ lệ NC dương tính', ['7/8', '87,5']),
    ('Qiaochu mobile lo âu tỷ lệ NC dương tính', ['2/6', '33,3']),
    ('Qiaochu Clinical Psychology Psychotherapy DOI', ['e70173', 'cpp']),
    # Menon 2025
    ('Menon 2025 scoping review LMIC', ['Menon', 'scoping']),
    ('Menon số nghiên cứu tổng hợp', ['69']),
    ('Menon số quốc gia LMIC', ['12 quốc gia']),
    ('Menon 32 RCT 31 trước sau 6 đánh giá', ['32 RCT', '31']),
    ('Menon khoảng trống cộng đồng gia đình dài hạn', ['CỘNG ĐỒNG', 'GIA ĐÌNH']),
    ('Menon Asia Pacific J Public Health pages', ['332', '346']),
    # Li 2025 BMC NMA
    ('Li 2025 BMC Psychiatry NMA Bayesian 30 RCT', ['BMC', 'NMA']),
    ('Li tổng số người tham gia 30 RCT', ['1.711']),
    ('Li ACT SUCRA hạng cao nhưng CrI', ['ACT', '0,69']),
    ('Li CBT SUCRA bằng chứng mạnh nhất', ['0,66', '16 RCT']),
    ('Li VRET SUCRA virtual reality', ['VRET', '0,51']),
    ('Li PE physical exercise SUCRA', ['PE', '0,51']),
]

# === GROUP 3: ÂU - ÚC - MỸ — Walder, Bress, Brown, Cao, Zhang, Zugman, Walkup (50 queries) ===
queries += [
    # Walder 2025
    ('Walder 2025 JMIR DMHI 21 RCT meta-analysis', ['Walder', 'JMIR']),
    ('Walder DMHI vs đối chứng tổng quát g', ['0,508']),
    ('Walder DMHI vs WL danh sách chờ', ['0,576']),
    ('Walder DMHI nền CBT Hedges g', ['0,610']),
    ('Walder DMHI có hướng dẫn người', ['0,825', 'hướng dẫn']),
    ('Walder DMHI thiết kế riêng SAD g lớn nhất', ['0,878']),
    ('Walder DMHI không hướng dẫn nhỏ', ['0,3', '0,4']),
    ('Walder PROSPERO CRD', ['CRD42023424181']),
    ('Walder J Med Internet Research Q1', ['JMIR']),
    # Bress 2024 — Maya app
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
    # Brown & Carter 2025
    ('Brown Carter 2025 J Mental Health Editorial', ['Brown', 'Carter']),
    ('Brown Carter PLACES self-referral mô hình', ['PLACES', 'self-referral']),
    ('Brown Carter BESST trial CBT 900 HS 57 trường', ['BESST', '900', '57']),
    ('Brown Carter Mindfulness Kuyken 8376 HS thất bại', ['8.376', 'Kuyken']),
    ('Brown Carter MHST Mental Health Support Team', ['MHST']),
    ('Brown Carter co-design học sinh tham gia thiết kế', ['co-design']),
    ('Brown Carter trang 357 361 vol 34', ['357', '361']),
    # Cao 2025 Resilience
    ('Cao 2025 Frontiers Psychiatry resilience MA RCT', ['Cao', 'Frontiers', 'resilience']),
    ('Cao DOI 1594658 article', ['1594658']),
    ('Cao kích thước hiệu ứng nhỏ trung bình', ['nhỏ', 'trung bình']),
    ('Cao heterogeneity cao', ['heterogeneity', 'CAO']),
    ('Cao Cochrane Risk of Bias đánh giá', ['Cochrane']),
    ('Cao positive psychology yếu tố bảo vệ', ['BẢO VỆ', 'lạc quan']),
    # Zhang 2026 Bayesian MA
    ('Zhang Liang Kang 2026 Bayesian MA J Clin Psychol', ['Zhang', 'Bayesian']),
    ('Zhang 31 RCT tổng hợp', ['31 RCT']),
    ('Zhang n học sinh nguy cơ thấp', ['19.865']),
    ('Zhang hiệu quả khiêm tốn trầm cảm lo âu', ['khiêm tốn', 'NHỎ']),
    ('Zhang chất lượng bằng chứng nền rất thấp', ['THẤP', 'chất lượng']),
    ('Zhang duy trì 1 năm sau can thiệp', ['1 năm', 'duy trì']),
    ('Zhang trang 248 259 vol 82', ['248', '259']),
    # Zugman 2024 AJP
    ('Zugman 2024 American Journal Psychiatry pediatric anxiety', ['Zugman', 'AJP']),
    ('Zugman pediatric anxiety treatment review', ['pediatric', 'review']),
    ('Zugman 3 rối loạn lo âu trẻ em GAD SAD', ['GAD', 'SAD']),
    ('Zugman trang 189 200 vol 181', ['189', '200', '181']),
    ('Zugman NIMH NIH NC SKTT hàng đầu Mỹ', ['NIMH', 'NIH']),
    # Walkup 2008 CAMS
    ('Walkup 2008 CAMS NEJM RCT đa trung tâm', ['Walkup', 'CAMS']),
    ('Walkup CAMS tổng số trẻ tham gia', ['488']),
    ('Walkup CAMS tuổi trẻ 7 17', ['7', '17']),
    ('Walkup CBT SSRI kết hợp đáp ứng', ['80,7']),
    ('Walkup CBT đơn thuần đáp ứng', ['59,7']),
    ('Walkup SSRI Sertraline đơn thuần', ['54,9', 'Sertraline']),
    ('Walkup placebo baseline đáp ứng', ['23,7', 'placebo']),
    ('Walkup CAMS NEJM 359 2753', ['NEJM', '2753']),
]

# === GROUP 4: TÓM LƯỢC + KHUYẾN NGHỊ + KẾT LUẬN (30 queries) ===
queries += [
    ('tóm lược 6 phát hiện then chốt', ['then chốt', 'phát hiện']),
    ('CBT phương pháp hiệu quả nhất xuyên suốt', ['CBT', 'hiệu quả nhất']),
    ('CAMS 80,7 phần trăm CBT SSRI gold standard', ['CAMS', '80,7']),
    ('iCBT SAD hạng 1 NMA Walder Xian', ['iCBT', 'SAD', 'hạng 1']),
    ('mobile CBT trầm cảm mạnh hơn lo âu', ['trầm cảm', 'lo âu', 'Qiaochu']),
    ('universal CBT phổ quát hiệu quả nhỏ', ['phổ quát', 'NHỎ', 'Zhang']),
    ('khoảng trống cực lớn Việt Nam 0 RCT', ['0 RCT', 'cực lớn']),
    ('Đông Á Thái Bình Dương LMIC thiếu cộng đồng', ['LMIC', 'cộng đồng', 'gia đình']),
    ('Bảng 11 so sánh ba vùng địa lý', ['ba vùng', 'so sánh']),
    ('Bảng 12 xếp hạng 13 phương pháp', ['xếp hạng', '13']),
    ('Việt Nam số bài can thiệp', ['1', 'Trần Nguyễn Ngọc']),
    ('Châu Á số bài can thiệp ngoài Việt Nam', ['7']),
    ('Châu Âu Úc Mỹ số bài can thiệp', ['7']),
    ('hạng 1 CBT SSRI kết hợp khuyến nghị VN', ['CBT + SSRI', 'kết hợp']),
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
    ('PLACES self-referral cơ chế tiếp cận', ['PLACES', 'self-referral']),
    ('Dong PLOS ONE 0,22 kênh giao tiếp', ['Dong', '0,22']),
    ('Chen 2025 141 nghiên cứu meta', ['Chen', '141']),
    ('emotional functioning yếu tố bảo vệ Chen', ['emotional', 'bảo vệ']),
    ('Trần Thảo Vi lạc quan trung gian', ['Trần Thảo Vi', 'lạc quan']),
]

# === GROUP 5: TLTK + DOI + Page references (30 queries) ===
queries += [
    ('Bress 2024 DOI 10.1001 jamanetworkopen 2024 28372', ['10.1001', '28372']),
    ('Walder 2025 DOI 10.2196 preprints 67067', ['10.2196', '67067']),
    ('Xian 2024 DOI 10.1016 jad', ['10.1016/j.jad', '08.097']),
    ('De Silva 2024 DOI 10.1186 s13034', ['10.1186', 's13034']),
    ('Cao 2025 DOI 10.3389 fpsyt 1594658', ['10.3389', '1594658']),
    ('Brown Carter DOI 10.1080 09638237', ['10.1080', '09638237']),
    ('Li 2025 DOI 10.1186 s12888', ['s12888', '07227']),
    ('Sasaki 2024 DOI 10.2196 55786', ['10.2196', '55786']),
    ('Zugman 2024 DOI 10.1176 appi ajp', ['10.1176', '20231037']),
    ('Praptomojati DOI 10.1016 ajp 103896', ['ajp', '103896']),
    ('Zhang Liang Kang 2026 DOI 10.1002 jclp', ['jclp', '70069']),
    ('Qiaochu 2025 DOI 10.1002 cpp', ['cpp', '70173']),
    ('Menon 2025 DOI 10.1177 10105395', ['10.1177', '10105395']),
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
# 4. RUN VERIFICATION
# ============================================================
print('\n' + '=' * 70)
print(f'KIEM TRA RAG / {len(queries)} QUERIES')
print('=' * 70)

passed = 0
failed = 0
results_log = []

for i, (q, expected_kws) in enumerate(queries, 1):
    q_emb = model.encode([q]).tolist()
    res = col.query(query_embeddings=q_emb, n_results=5,
                    include=['documents', 'distances', 'metadatas'])
    docs = res['documents'][0]
    dists = res['distances'][0]
    metas_r = res['metadatas'][0]
    top_text = ' '.join(docs).lower()
    found_kws = [kw for kw in expected_kws if kw.lower() in top_text]
    n_found = len(found_kws)
    n_expected = len(expected_kws)
    rel = max(0, (1 - dists[0]) * 100)
    # PASS: at least 1 expected keyword found AND relevance >= 25
    success = n_found >= 1 and rel >= 25

    status = 'PASS' if success else 'FAIL'
    if success: passed += 1
    else: failed += 1

    results_log.append({
        'idx': i,
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

print(f'\n{"="*70}')
print(f'TONG KET: {passed}/{len(queries)} PASS ({passed/len(queries)*100:.1f}%) | {failed} FAIL')
print(f'{"="*70}')

# Stats per group
groups = [
    ('GROUP 1: VN — Trần Nguyễn Ngọc', 1, 40),
    ('GROUP 2: Châu Á', 41, 90),
    ('GROUP 3: Âu - Úc - Mỹ', 91, 140),
    ('GROUP 4: Tóm lược + Khuyến nghị', 141, 170),
    ('GROUP 5: TLTK + DOI', 171, 200),
]
for name, lo, hi in groups:
    sub = [r for r in results_log if lo <= r['idx'] <= hi]
    sub_pass = sum(1 for r in sub if r['status'] == 'PASS')
    print(f'  {name}: {sub_pass}/{len(sub)} ({sub_pass/max(len(sub),1)*100:.0f}%)')

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
        'results': results_log,
    }, f, ensure_ascii=False, indent=2)
print(f'\nLog saved: {log_path}')

# Show all FAILS
fails = [r for r in results_log if r['status'] == 'FAIL']
if fails:
    print(f'\n--- ALL {len(fails)} FAILS ---')
    for r in fails:
        print(f'\n[{r["idx"]:3d}] {r["query"][:75]}')
        print(f'      Expected: {r["expected"]}')
        print(f'      Found: {r["found"]}, rel={r["relevance"]}%')
        print(f'      Top section: {r["top_section"]}')
        print(f'      Top excerpt: {r["top_doc_excerpt"][:120]}')

db_size = sum(os.path.getsize(os.path.join(dp, f))
              for dp, dn, fn in os.walk(DB_PATH) for f in fn) // 1024
print(f'\nDB size: {db_size} KB')
