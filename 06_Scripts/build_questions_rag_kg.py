"""
Build RAG + KG cho các câu hỏi thầy đã hỏi trong tuần qua.
Output:
- web/data/rag_questions_index.json  (TF-IDF index)
- web/data/questions_kg.json         (Knowledge graph)
- 06_Scripts/questions_kg_data/*.graphml (networkx export)
"""
import sys, io, os, json, re, math
from pathlib import Path
from collections import Counter
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = Path(__file__).resolve().parent.parent
OUT_DATA = BASE / '06_Scripts' / 'questions_kg_data'
OUT_DATA.mkdir(parents=True, exist_ok=True)

# Các destination để deploy vào chatbot
WEB_HEAVY = BASE / 'tro-ly-nghien-cuu-tam-ly' / 'web' / 'data'
WEB_LIGHT = BASE / 'tro-ly-nghien-cuu-tam-ly-light' / 'web' / 'data'

# ================================================================
# DỮ LIỆU 15 CÂU HỎI + TRẢ LỜI + METADATA
# ================================================================
QUESTIONS = [
    {
        'id': 'QA_01',
        'date_asked': '2026-04-20',
        'question': "Cohen's d, hoặc Hedges g gọi là gì? Chỉ số hay thước đo, hay công thức? Khi nào dùng từng cái một? Nó có tương đương với SMD không?",
        'short_answer': ("Cả 2 đều là THƯỚC ĐO độ lớn hiệu ứng, cụ thể là 2 công thức tính SMD. "
                         "Cohen's d = (M₁ − M₂) / SD_pooled. Hedges' g = d × hệ số hiệu chỉnh J cho mẫu nhỏ. "
                         "Mẫu lớn (n ≥ 20/nhóm) dùng Cohen's d, mẫu nhỏ dùng Hedges' g. "
                         "Ngưỡng Cohen (1988): 0,2 nhỏ, 0,5 trung bình, 0,8 lớn."),
        'topic': 'Effect size',
        'concepts': ["Cohen's d", "Hedges' g", 'SMD', 'Effect size', 'Cochrane'],
        'doc_path': 'glossary.json (v3.1)',
        'paper_refs': ['QT049', 'QT029', 'QT040'],
    },
    {
        'id': 'QA_02',
        'date_asked': '2026-04-20',
        'question': ("Walder et al. (2025) — Meta-analysis 21 RCT về DMHI cho lo âu xã hội. "
                     "Các tác giả lọc ra 22 nghiên cứu, sau đó lại đưa về 21 nghiên cứu. "
                     "Tiêu chí nào để lọc ra 22 NC từ hàng trăm dữ liệu? Tiêu chí nào để gộp lại 21 NC?"),
        'short_answer': ("Thực tế 2.149 records → loại 474 duplicates → title screening loại 1.357 → "
                         "abstract screening loại 205 → 113 full-text → 22 đạt tiêu chí PICOS vào SR. "
                         "PICOS: P (TB <25 tuổi), I (DMHI remote), C (active/waitlist/CAU/TAU), "
                         "O (lo âu xã hội), S (RCT peer-reviewed Anh/Đức). "
                         "1 bài (Vigerland et al. [93]) bị loại khỏi MA vì thiếu dữ liệu post-assessment."),
        'topic': 'SR/MA methodology',
        'concepts': ['PICOS', 'Eligibility criteria', 'PRISMA flow', 'SR vs MA'],
        'doc_path': '01_Bao-cao/QT040_Walder_2025_Quy_trinh_loc_22_vs_21.docx',
        'paper_refs': ['QT040'],
    },
    {
        'id': 'QA_03',
        'date_asked': '2026-04-20',
        'question': "Waitlist cũng đạt tiêu chuẩn chọn lọc nhưng vì sao phải chờ?",
        'short_answer': ("Waitlist là nhóm ĐỐI CHỨNG trong RCT — họ đạt tiêu chí nhưng được phân ngẫu nhiên "
                         "vào nhóm chờ 8-12 tuần, sau đó mới nhận can thiệp. Mục đích: làm 'mốc so sánh' "
                         "để đo effect thực của can thiệp (trừ đi tự hồi phục + Hawthorne effect). "
                         "Đạo đức hơn no-treatment vì cuối cùng ai cũng được nhận can thiệp."),
        'topic': 'RCT design',
        'concepts': ['Waitlist', 'Control group', 'RCT', 'Effect size'],
        'doc_path': 'chat only',
        'paper_refs': ['QT040'],
    },
    {
        'id': 'QA_04',
        'date_asked': '2026-04-20',
        'question': "CAMS là RCT đa trung tâm 488 trẻ 7–17 tuổi, từ 'đa trung tâm' này nghĩa là gì?",
        'short_answer': ("RCT đa trung tâm = 1 nghiên cứu triển khai đồng thời tại NHIỀU địa điểm (bệnh viện, "
                         "trường đại học) theo CÙNG 1 protocol. Lợi ích: (1) tuyển đủ mẫu nhanh hơn; "
                         "(2) đa dạng dân số (generalizability cao); (3) giảm site-specific bias; "
                         "(4) chia sẻ chi phí; (5) chứng minh can thiệp khả thi nhân rộng. "
                         "CAMS = Child/Adolescent Anxiety Multimodal Study (Walkup 2008 NEJM)."),
        'topic': 'RCT design',
        'concepts': ['Multicenter RCT', 'CAMS', 'Generalizability', 'PROSPERO'],
        'doc_path': '01_Bao-cao/CAMS_RCT_da_trung_tam_la_gi.docx',
        'paper_refs': ['QT028', 'CAMS_Walkup_2008'],
    },
    {
        'id': 'QA_05',
        'date_asked': '2026-04-20',
        'question': ("ACT ra đời như thế nào? Vì sao chỉ có một nghiên cứu xếp hiệu quả can thiệp "
                     "của ACT cao hơn CBT?"),
        'short_answer': ("ACT do Steven C. Hayes phát triển từ đầu 1980s, hệ thống hoá năm 1999 "
                         "(Hayes/Strosahl/Wilson). Là 'thế hệ thứ 3' của CBT, dựa trên RFT "
                         "(Relational Frame Theory). 6 tiến trình hexaflex: acceptance, defusion, "
                         "present moment, self-as-context, values, committed action. "
                         "Chỉ Li 2025 BMC NMA xếp ACT > CBT vì: (1) ACT ít RCT hơn CBT nhiều lần; "
                         "(2) CrI của ACT chứa số dương (không chắc chắn hơn chứng); "
                         "(3) effect size ACT và CBT gần bằng nhau; "
                         "(4) authors tự ghi 'low quality of evidence'; "
                         "(5) 4/5 NMA khác vẫn xếp CBT hạng 1."),
        'topic': 'Psychotherapy history',
        'concepts': ['ACT', 'CBT', 'RFT', 'Steven Hayes', 'Hexaflex', 'Third wave', 'Bayesian NMA'],
        'doc_path': '01_Bao-cao/ACT_ra_doi_va_vi_sao_xep_hang_cao_hon_CBT.docx',
        'paper_refs': ['QT029', 'QT028'],
    },
    {
        'id': 'QA_06',
        'date_asked': '2026-04-20',
        'question': "NNT diễn giải cụ thể như thế nào? Vì sao chỉ số này càng nhỏ thì can thiệp càng hiệu quả?",
        'short_answer': ("NNT (Number Needed to Treat) = số người cần điều trị để có THÊM 1 người đáp ứng "
                         "so với đối chứng. Công thức: NNT = 1 / ARR (Absolute Risk Reduction). "
                         "Càng nhỏ = cần ít người = hiệu suất cao. "
                         "CAMS combination (CBT+Sertraline vs Placebo): NNT = 3 là 'xuất sắc'. "
                         "Ngưỡng: 1-3 xuất sắc, 4-10 rất tốt, 11-20 tốt, 21-50 TB, >50 thấp."),
        'topic': 'Clinical statistics',
        'concepts': ['NNT', 'ARR', 'Response rate', 'CAMS', 'Effect size'],
        'doc_path': '01_Bao-cao/NNT_dien_giai_va_vi_sao_cang_nho_cang_hieu_qua.docx',
        'paper_refs': ['QT028', 'CAMS_Walkup_2008'],
    },
    {
        'id': 'QA_07',
        'date_asked': '2026-04-20',
        'question': ("Bài Jenkins 2023: 'học sinh nữ có điểm GAD10 cao hơn học sinh nam "
                     "(p=0,016, kiểm định Mann-Whitney), với chênh lệch trung vị là 3 điểm'. "
                     "Chênh lệch trung vị 3 điểm, theo kiểm định này nghĩa là gì?"),
        'short_answer': ("Mann-Whitney là kiểm định PHI THAM SỐ so 2 nhóm độc lập theo thứ hạng (ranks), "
                         "dùng khi dữ liệu không phân phối chuẩn. p=0,016<0,05 → có khác biệt thật "
                         "giữa nữ/nam. 'Chênh lệch trung vị 3 điểm' thường tính bằng Hodges-Lehmann "
                         "estimator — trung vị của TẤT CẢ các cặp hiệu (nữ_i − nam_j). "
                         "Nghĩa là phân phối điểm GAD-10 của nữ dịch lên cao hơn nam ~3 điểm (trên thang 0-30). "
                         "3/30 ≈ 10% thang, có ý nghĩa lâm sàng khiêm tốn."),
        'topic': 'Statistical test',
        'concepts': ['Mann-Whitney', 'Hodges-Lehmann', 'Median difference', 'Non-parametric', 'GAD-10'],
        'doc_path': '01_Bao-cao/NNT_va_MannWhitney_chenh_lech_trung_vi.docx',
        'paper_refs': ['QT001'],
    },
    {
        'id': 'QA_08',
        'date_asked': '2026-04-20',
        'question': "AOR (Adjusted Odds Ratio) — tỷ số chênh lệch đã điều chỉnh được tính toán như thế nào?",
        'short_answer': ("AOR = Adjusted Odds Ratio — OR đã điều chỉnh cho confounders. "
                         "Tính từ HỒI QUY LOGISTIC ĐA BIẾN: log(odds) = β₀ + β₁X₁ + β₂X₂ + ... "
                         "AOR = exp(β). KTC 95% = exp(β ± 1,96·SE). "
                         "Ngưỡng: AOR>1 nguy cơ, AOR<1 bảo vệ, =1 không liên hệ, KTC chứa 1 không ý nghĩa. "
                         "Ví dụ: Zhu 2025 ngủ<5h AOR=13,71; Islam 2025 cha mẹ quan tâm AOR=0,75 (bảo vệ)."),
        'topic': 'Regression',
        'concepts': ['AOR', 'OR', 'Logistic regression', 'Multivariable', 'Confounding', 'β coefficient'],
        'doc_path': '01_Bao-cao/AOR_Adjusted_Odds_Ratio_cach_tinh.docx',
        'paper_refs': ['QT05_Zhu', 'QT23_JAACAP', 'QT31_Islam'],
    },
    {
        'id': 'QA_09',
        'date_asked': '2026-04-22',
        'question': "Đọc PRISMA flow là đọc như thế nào?",
        'short_answer': ("PRISMA 2020 flow có 4 giai đoạn: (I) Identification - tìm records + loại duplicates; "
                         "(II) Screening - đọc title+abstract, 2 reviewer độc lập; "
                         "(III) Eligibility - full-text assessment, PHẢI LIỆT KÊ LÝ DO LOẠI; "
                         "(IV) Included - số bài vào SR và MA. "
                         "3 điều cần kiểm: (1) số học cộng đúng; (2) có lý do loại cụ thể; "
                         "(3) có 2 reviewer độc lập. Thiếu 1 = cờ đỏ chất lượng."),
        'topic': 'SR methodology',
        'concepts': ['PRISMA', 'Flow diagram', 'Screening', 'Eligibility', 'Risk of bias'],
        'doc_path': '01_Bao-cao/PRISMA_flow_cach_doc_tung_buoc.docx',
        'paper_refs': ['QT040', 'QT029'],
    },
    {
        'id': 'QA_10',
        'date_asked': '2026-04-22',
        'question': "Chi bình phương là hồi quy đơn biến, phân tích logistic là hồi quy đa biến — cách gọi này có đúng không?",
        'short_answer': ("KHÔNG ĐÚNG. Chi-square là KIỂM ĐỊNH (không phải hồi quy) — test liên hệ 2 biến "
                         "phân loại qua bảng chéo. Hồi quy logistic LÀ hồi quy, có 2 dạng: "
                         "univariable (1 biến X) và multivariable (nhiều biến X). "
                         "Chi-square và logistic đơn biến cho p-value tương đương khi X là phân loại, "
                         "nhưng logistic thêm OR + KTC 95%. "
                         "Phân biệt: univariate = 1 outcome, multivariate = nhiều outcome (≠ univariable/multivariable)."),
        'topic': 'Statistical terminology',
        'concepts': ['Chi-square', 'Logistic regression', 'Univariable', 'Multivariable', 'Hypothesis testing'],
        'doc_path': 'chat only',
        'paper_refs': [],
    },
    {
        'id': 'QA_11',
        'date_asked': '2026-04-22',
        'question': "Trong Chen 2023: 'Hiệu ứng cụm được điều chỉnh bằng gói survey của R để tính tỷ lệ có trọng số' nghĩa là gì?",
        'short_answer': ("Lấy mẫu cụm (63.205 HS, 2 quận + 1 huyện Tự Cống) khiến HS cùng trường tương quan "
                         "cao (ICC > 0). Nếu không điều chỉnh, SE bị đánh giá thấp → p sai. "
                         "Gói `survey` của Thomas Lumley (R) xử lý cluster + weight: "
                         "`svydesign(ids=~school, strata=~district, weights=~w)`. "
                         "Tỷ lệ có trọng số p = Σ(wᵢ×yᵢ) / Σ(wᵢ), với wᵢ = 1/xác suất được chọn. "
                         "Mục đích: ngoại suy đúng cho quần thể, không bị lệch do design."),
        'topic': 'Survey methodology',
        'concepts': ['Cluster effect', 'ICC', 'Weighted prevalence', 'survey package R', 'Design effect'],
        'doc_path': 'chat only + 01_Bao-cao/Weighted_prevalence_ty_le_co_trong_so.docx',
        'paper_refs': ['QT007_Chen_2023'],
    },
    {
        'id': 'QA_12',
        'date_asked': '2026-04-23',
        'question': "SE có phải là 'size' không? DEFF có phải là 'design effect' không?",
        'short_answer': ("SE = Standard Error (sai số chuẩn), KHÔNG phải 'size'. "
                         "SE(mean) = SD/√n; dùng để tính KTC = ước tính ± 1,96·SE. "
                         "'Size' thường là n (sample size) hoặc ES (effect size). "
                         "DEFF = Design Effect — thầy nhớ đúng. "
                         "DEFF = 1 + (m̄−1)·ρ với m̄ cỡ cụm TB, ρ là ICC. "
                         "SE_cluster = SE_SRS × √DEFF."),
        'topic': 'Statistical terminology',
        'concepts': ['SE', 'DEFF', 'Design effect', 'Sample size', 'Effect size', 'ICC'],
        'doc_path': 'chat only',
        'paper_refs': [],
    },
    {
        'id': 'QA_13',
        'date_asked': '2026-04-23',
        'question': ("Câu: 'Bỏ qua cấu trúc phân cấp của dữ liệu dẫn đến sai số chuẩn bị đánh giá thấp "
                     "và giá trị p bị phóng đại' — cụm 'sai số chuẩn bị đánh giá' nghĩa là gì?"),
        'short_answer': ("Parse đúng: 'sai số chuẩn | BỊ đánh giá thấp' = SE underestimated = "
                         "công thức cho ra SE NHỎ HƠN thực tế. Logic: cluster không adjust → SE nhỏ giả → "
                         "|t| = estimate/SE lớn giả → p nhỏ giả → kết luận 'có ý nghĩa' sai "
                         "(false positive). 'Giá trị p bị phóng đại' = 'inflated significance/Type I error' "
                         "= tỷ lệ kết quả có ý nghĩa giả tăng lên."),
        'topic': 'Statistical interpretation',
        'concepts': ['Standard error', 'Underestimation', 'Inflated significance', 'Type I error', 'Cluster effect'],
        'doc_path': 'chat only',
        'paper_refs': ['QT007_Chen_2023'],
    },
    {
        'id': 'QA_14',
        'date_asked': '2026-04-23',
        'question': "Tỷ lệ có trọng số (weighted prevalence) nghĩa là như thế nào?",
        'short_answer': ("Weighted prevalence = tỷ lệ điều chỉnh bằng trọng số. "
                         "Công thức: p = Σ(wᵢ × yᵢ) / Σ(wᵢ), trong đó wᵢ = 1/xác suất chọn "
                         "(có thể bổ sung non-response adjustment + post-stratification). "
                         "Cần dùng khi: survey quốc gia, stratified cluster sampling, over-sampling nhóm thiểu số. "
                         "Không cần với RCT hoặc simple random sampling. "
                         "Trong R: gói `survey` với svydesign + svyciprop."),
        'topic': 'Survey methodology',
        'concepts': ['Weighted prevalence', 'Sampling weight', 'Non-response', 'Post-stratification', 'survey package R'],
        'doc_path': '01_Bao-cao/Weighted_prevalence_ty_le_co_trong_so.docx',
        'paper_refs': ['QT007_Chen_2023'],
    },
    {
        'id': 'QA_15',
        'date_asked': '2026-04-24',
        'question': ("Trong bài 07 có câu: 'khoảng tin cậy khá rộng cho trầm cảm (19,6–27,0%) và "
                     "lo âu (11,2–17,0%)'. KTC thế nào là vừa, là đủ, là trung bình?"),
        'short_answer': ("Ngưỡng thực hành dịch tễ (width = cận trên − cận dưới): "
                         "< 2 điểm % hẹp (rất chính xác); 2–5 vừa (chuẩn Q1); "
                         "5–10 rộng; > 10 rất rộng. Chen 2023: trầm cảm width 7,4 và lo âu width 5,8 "
                         "→ thuộc mức RỘNG. Rộng dù n=63.205 vì DEFF cluster làm effective n ~4-12k. "
                         "Khi đọc KTC rộng: ước tính 'có thể nằm đâu đó trong khoảng', "
                         "không so sánh 2 nhóm nếu KTC chồng lấn."),
        'topic': 'Statistical interpretation',
        'concepts': ['Confidence interval', 'CI width', 'Precision', 'Design effect', 'Effective sample size'],
        'doc_path': '01_Bao-cao/KTC_rong_hep_vua_du_Chen_2023.docx',
        'paper_refs': ['QT007_Chen_2023'],
    },
]

# ================================================================
# TOKENIZE + BUILD TF-IDF
# ================================================================
STOPWORDS = set('''a an and or but in on at to from for of with by is are was were be been being
have has had do does did the this that these those it its they them their we us our you your
he she his her i me my là và hoặc nhưng trong tại đến từ cho của với bởi là các một những này đó
nó họ chúng ta chúng tôi tôi bạn anh ấy cô ấy có được bị đã đang sẽ cũng không chỉ rất đều như
ai ở khi nào tới sau trước giữa ngoài trên dưới bên cùng qua vào ra lên xuống đây đó kia gì thì mà'''.split())

def tokenize(text):
    if not text: return []
    text = text.lower()
    tokens = re.findall(r'\b[\wÀ-ỹđĐ]+\b', text, flags=re.UNICODE)
    return [t for t in tokens if t not in STOPWORDS and len(t) > 1]

def build_tf(entry):
    parts = [
        entry['question'],
        entry['short_answer'],
        entry['topic'],
        ' '.join(entry['concepts']),
    ]
    text = '\n'.join(parts)
    return dict(Counter(tokenize(text)))

# Build TF cho từng entry
for q in QUESTIONS:
    q['tokens'] = build_tf(q)

# Build IDF global
N = len(QUESTIONS)
df = Counter()
for q in QUESTIONS:
    for tok in q['tokens']:
        df[tok] += 1
idf = {tok: math.log((N + 1) / (dfi + 1)) + 1 for tok, dfi in df.items()}

# RAG index
rag_index = {
    'meta': {
        'created': '2026-04-24',
        'version': 'questions-v1',
        'n_entries': N,
        'total_unique_tokens': len(idf),
        'description': '15 câu hỏi thầy đã hỏi về thống kê, methodology, psychotherapy (20-24/04/2026)',
    },
    'idf': idf,
    'entries': QUESTIONS,
}

# Save RAG
for dest_dir in [OUT_DATA, WEB_HEAVY, WEB_LIGHT]:
    if dest_dir.exists():
        dest = dest_dir / 'rag_questions_index.json'
        with open(dest, 'w', encoding='utf-8') as f:
            json.dump(rag_index, f, ensure_ascii=False, indent=1)
        print(f'RAG saved: {dest} ({dest.stat().st_size/1024:.1f} KB)')

# ================================================================
# BUILD KNOWLEDGE GRAPH
# ================================================================
kg_nodes = {}
kg_edges = []

def add_node(node_id, node_type, label, **attrs):
    if node_id not in kg_nodes:
        kg_nodes[node_id] = {'id': node_id, 'type': node_type, 'label': label, **attrs}

def add_edge(source, target, relation, **attrs):
    kg_edges.append({'source': source, 'target': target, 'relation': relation, **attrs})

# 1. Node: Questions
for q in QUESTIONS:
    add_node(q['id'], 'Question', q['question'][:80],
             topic=q['topic'], date_asked=q['date_asked'])

# 2. Node: Concepts — gom từ concepts list của các Q
concept_to_id = {}
def slug(s):
    return 'CONCEPT_' + re.sub(r'\W+', '_', s.upper()).strip('_')

for q in QUESTIONS:
    for c in q['concepts']:
        cid = slug(c)
        concept_to_id[c] = cid
        add_node(cid, 'Concept', c)
        add_edge(q['id'], cid, 'EXPLAINS')

# 3. Node: Topics
topic_to_id = {}
for q in QUESTIONS:
    topic = q['topic']
    tid = 'TOPIC_' + re.sub(r'\W+', '_', topic.upper()).strip('_')
    topic_to_id[topic] = tid
    add_node(tid, 'Topic', topic)
    add_edge(q['id'], tid, 'BELONGS_TO')

# 4. Node: Docs
for q in QUESTIONS:
    doc = q['doc_path']
    if doc and doc != 'chat only' and 'chat only' not in doc.lower():
        did = 'DOC_' + re.sub(r'\W+', '_', doc).strip('_')[:50]
        add_node(did, 'Doc', doc)
        add_edge(q['id'], did, 'ANSWERED_IN')

# 5. Node: Papers
for q in QUESTIONS:
    for p in q['paper_refs']:
        pid = 'PAPER_' + re.sub(r'\W+', '_', p).strip('_')
        add_node(pid, 'Paper', p)
        add_edge(q['id'], pid, 'CITES')

# 6. Edge: Concept ↔ Concept (co-occurrence trong cùng Q)
from itertools import combinations
concept_pairs = Counter()
for q in QUESTIONS:
    cids = [concept_to_id[c] for c in q['concepts']]
    for a, b in combinations(sorted(cids), 2):
        concept_pairs[(a, b)] += 1

for (a, b), w in concept_pairs.items():
    add_edge(a, b, 'RELATED_TO', weight=w)

# Stats
type_counter = Counter(n['type'] for n in kg_nodes.values())
rel_counter = Counter(e['relation'] for e in kg_edges)
print()
print('=== KG STATS ===')
print(f'Nodes: {len(kg_nodes)}')
for t, c in type_counter.most_common():
    print(f'  {t}: {c}')
print(f'Edges: {len(kg_edges)}')
for r, c in rel_counter.most_common():
    print(f'  {r}: {c}')

# Save KG
kg_data = {
    'meta': {
        'created': '2026-04-24',
        'version': 'questions-kg-v1',
        'n_nodes': len(kg_nodes),
        'n_edges': len(kg_edges),
        'node_types': dict(type_counter),
        'edge_relations': dict(rel_counter),
    },
    'nodes': list(kg_nodes.values()),
    'edges': kg_edges,
}

for dest_dir in [OUT_DATA, WEB_HEAVY, WEB_LIGHT]:
    if dest_dir.exists():
        dest = dest_dir / 'questions_kg.json'
        with open(dest, 'w', encoding='utf-8') as f:
            json.dump(kg_data, f, ensure_ascii=False, indent=1)
        print(f'KG saved: {dest} ({dest.stat().st_size/1024:.1f} KB)')

# Export GraphML qua networkx
try:
    import networkx as nx
    G = nx.DiGraph()
    for n in kg_nodes.values():
        attrs = {k: (str(v) if not isinstance(v, (str, int, float, bool)) else v)
                 for k, v in n.items() if k != 'id'}
        G.add_node(n['id'], **attrs)
    for e in kg_edges:
        w = e.get('weight', 1)
        G.add_edge(e['source'], e['target'], relation=e['relation'], weight=w)
    gml = OUT_DATA / 'questions_kg_v1.graphml'
    nx.write_graphml(G, str(gml))
    print(f'GraphML: {gml} ({gml.stat().st_size/1024:.1f} KB)')
except Exception as e:
    print(f'GraphML export skipped: {e}')

# ================================================================
# VERIFY với smoke test
# ================================================================
print()
print('=== SMOKE TEST RAG ===')
def search(q, top_k=3):
    qt = tokenize(q)
    scored = []
    for e in rag_index['entries']:
        s = sum(e['tokens'].get(t, 0) * idf.get(t, 1.0) for t in qt)
        # Title/concept boost
        for t in qt:
            if any(t in c.lower() for c in e['concepts']):
                s += 5
        if s > 0:
            scored.append((s, e))
    scored.sort(reverse=True, key=lambda x: x[0])
    return scored[:top_k]

for q in ['NNT là gì?', 'AOR tính thế nào?', 'Mann-Whitney median', 'ACT ra đời', 'KTC rộng']:
    print(f"Q: '{q}' →")
    for s, e in search(q):
        print(f'  [{s:.2f}] {e["id"]}: {e["question"][:70]}')
    print()

print('DONE')
