"""Update KG + RAG cho doc phan bien VN020 + bo sung VN020 vao KG."""
import sys, io, json, copy, re
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

LIGHT = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly-light/web/data')
HEAVY = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly/web/data')

NEW_CONCEPTS = [
    {'id': 'CONCEPT_DASS_Y', 'type': 'Concept', 'label': 'DASS-Y (Depression Anxiety Stress Scales for Youth, Szabó 2010)'},
    {'id': 'CONCEPT_LOGISTIC_VS_SEM', 'type': 'Concept', 'label': 'Hồi quy logistic vs SEM — so sánh phương pháp phân tích'},
    {'id': 'CONCEPT_DICHOTOMIZATION_BIAS', 'type': 'Concept', 'label': 'Dichotomization bias (mất thông tin liên tục khi dùng ngưỡng cắt)'},
    {'id': 'CONCEPT_SAMPLING_DESIGN', 'type': 'Concept', 'label': 'Thiết kế lấy mẫu (PPS, ngẫu nhiên đa bậc, thuận tiện)'},
]

NEW_PAPERS = [
    {'id': 'PAPER_VN020_TranHoVinhLoc_2024', 'type': 'Paper',
     'label': 'VN020 Trần Hồ Vĩnh Lộc et al. (2024) DASS-Y + ESSA TPHCM'},
    {'id': 'PAPER_VN029_Truc_2025', 'type': 'Paper',
     'label': 'VN029 Trúc Thanh Thái et al. (2025) Q1 Soc Psychiatry n=2.631'},
    {'id': 'PAPER_Szabo_2010_DASS_Y', 'type': 'Paper',
     'label': 'Szabó (2010) DASS-Y phiên bản VTN'},
]

NEW_DOCS = [
    {'id': 'DOC_PHAN_BIEN_VN020_ALHT', 'type': 'Doc',
     'label': '01_Bao-cao/Phan_bien_VN020_TranHoVinhLoc_2024_ALHT.docx'},
]

NEW_QUESTIONS = [
    {'id': 'QA_31', 'type': 'Question',
     'label': 'Phản biện VN020 Trần Hồ Vĩnh Lộc (2024) về cách đo áp lực học tập — 7 hạn chế phương pháp luận',
     'topic': 'Statistical interpretation', 'date_asked': '2026-05-09'},
]

NEW_EDGES = [
    ('QA_31', 'TOPIC_STATISTICAL_INTERPRETATION', 'BELONGS_TO'),
    ('QA_31', 'CONCEPT_DASS_Y', 'EXPLAINS'),
    ('QA_31', 'CONCEPT_LOGISTIC_VS_SEM', 'EXPLAINS'),
    ('QA_31', 'CONCEPT_DICHOTOMIZATION_BIAS', 'EXPLAINS'),
    ('QA_31', 'CONCEPT_SAMPLING_DESIGN', 'EXPLAINS'),
    ('QA_31', 'CONCEPT_ESSA', 'EXPLAINS'),
    ('QA_31', 'PAPER_VN020_TranHoVinhLoc_2024', 'CITES'),
    ('QA_31', 'PAPER_VN029_Truc_2025', 'CITES'),
    ('QA_31', 'PAPER_Szabo_2010_DASS_Y', 'CITES'),
    ('QA_31', 'PAPER_Sun_2011_ESSA', 'CITES'),
    ('QA_31', 'DOC_PHAN_BIEN_VN020_ALHT', 'ANSWERED_IN'),
    ('PAPER_VN020_TranHoVinhLoc_2024', 'CONCEPT_ESSA', 'RELATED_TO'),
    ('PAPER_VN020_TranHoVinhLoc_2024', 'CONCEPT_DASS_Y', 'RELATED_TO'),
    ('PAPER_Szabo_2010_DASS_Y', 'CONCEPT_DASS_Y', 'RELATED_TO'),
    ('DOC_PHAN_BIEN_VN020_ALHT', 'PAPER_VN020_TranHoVinhLoc_2024', 'CITES'),
    ('DOC_PHAN_BIEN_VN020_ALHT', 'PAPER_VN029_Truc_2025', 'CITES'),
    ('DOC_PHAN_BIEN_VN020_ALHT', 'PAPER_Sun_2011_ESSA', 'CITES'),
    ('DOC_PHAN_BIEN_VN020_ALHT', 'PAPER_QT040', 'CITES'),
]

NEW_RAG_QUESTIONS = [
    {
        'id': 'QA_31',
        'date_asked': '2026-05-09',
        'question': 'Phản biện VN020 Trần Hồ Vĩnh Lộc (2024) về cách đo áp lực học tập — bài VN tốt nhất hiện có nhưng còn hạn chế gì?',
        'short_answer': 'VN020 (n=976 HS THPT TPHCM, DASS-Y + ESSA, ESSA ≥ 59 yếu tố mạnh nhất, lo âu 25,1%) là bài VN TỐT NHẤT về ALHT đến nay (dùng ESSA chuẩn quốc tế). Tuy nhiên có 7 hạn chế: (1) chỉ 2 trường TPHCM — không đại diện; (2) chọn lớp thuận tiện ở bậc lớp; (3) hồi quy logistic + ngưỡng cắt nhị phân — mất thông tin liên tục, không có β chuẩn hóa và R²; (4) DASS-Y chỉ 1 thang lo âu — không phân tách 3 dạng RLLA như RCADS; (5) bỏ sót 5 yếu tố quan trọng (bắt nạt, tự trọng, hỗ trợ XH, NĐT, BPĐP); (6) DASS-Y chưa phổ biến quốc tế — tỷ lệ 25,1% thấp hơn DASS-21 (40-86%) gợi ý ngưỡng quá nghiêm; (7) tạp chí Y học TPHCM không PubMed/Scopus. Đề xuất: mở rộng đa trường, dùng SEM với RCADS, bổ sung 5 yếu tố, lặp lại với DASS-21 để so sánh ngưỡng.',
        'topic': 'Statistical interpretation',
        'concepts': ['DASS-Y', 'Logistic vs SEM', 'Dichotomization bias', 'Sampling design', 'ESSA'],
        'doc_path': '01_Bao-cao/Phan_bien_VN020_TranHoVinhLoc_2024_ALHT.docx',
        'paper_refs': ['VN020', 'VN029'],
    },
]


def tokenize(text):
    text = text.lower()
    counts = {}
    for t in re.findall(r'[\w]+', text, re.UNICODE):
        if len(t) >= 2: counts[t] = counts.get(t, 0) + 1
    return counts


def update(data_dir):
    path = data_dir / 'questions_kg.json'
    with open(path, encoding='utf-8') as f: kg = json.load(f)
    eids = {n['id'] for n in kg['nodes']}
    eedges = {(e['source'],e['target'],e['relation']) for e in kg['edges']}
    na = 0
    for n in NEW_CONCEPTS + NEW_PAPERS + NEW_DOCS + NEW_QUESTIONS:
        if n['id'] not in eids: kg['nodes'].append(n); eids.add(n['id']); na += 1
    ea = 0
    for src,tgt,rel in NEW_EDGES:
        if (src,tgt,rel) in eedges or src not in eids or tgt not in eids: continue
        kg['edges'].append({'source':src,'target':tgt,'relation':rel}); eedges.add((src,tgt,rel)); ea += 1
    from collections import Counter
    kg['meta']['n_nodes']=len(kg['nodes']); kg['meta']['n_edges']=len(kg['edges'])
    kg['meta']['node_types']=dict(Counter(n['type'] for n in kg['nodes']))
    kg['meta']['edge_relations']=dict(Counter(e['relation'] for e in kg['edges']))
    kg['meta']['updated']='2026-05-09'
    with open(path,'w',encoding='utf-8') as f: json.dump(kg,f,ensure_ascii=False,indent=2)
    nn = len(kg['nodes']); ne = len(kg['edges'])
    print(f'  KG: +{na} nodes, +{ea} edges; total {nn}/{ne}')

    path = data_dir / 'rag_questions_index.json'
    with open(path, encoding='utf-8') as f: rqi = json.load(f)
    eids = {e['id'] for e in rqi['entries']}
    added = 0
    for nq in NEW_RAG_QUESTIONS:
        if nq['id'] in eids: continue
        e = copy.deepcopy(nq)
        e['tokens'] = tokenize(e['question']+' '+e['short_answer'])
        rqi['entries'].append(e); added += 1
    rqi['meta']['n_entries']=len(rqi['entries']); rqi['meta']['updated']='2026-05-09'
    with open(path,'w',encoding='utf-8') as f: json.dump(rqi,f,ensure_ascii=False,indent=2)
    nr = len(rqi['entries'])
    print(f'  rag_questions: +{added}; total {nr}')


for label, dir_ in [('LIGHT',LIGHT),('HEAVY',HEAVY)]:
    print(f'\n==== {label} ====')
    if not dir_.exists(): print('missing'); continue
    update(dir_)
print('\nDone.')
