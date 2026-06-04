"""Update KG + RAG cho doc phan bien VN014 thieu ALHT."""
import sys, io, json, copy, re
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

LIGHT = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly-light/web/data')
HEAVY = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly/web/data')

NEW_CONCEPTS = [
    {'id': 'CONCEPT_OMITTED_VARIABLE_BIAS', 'type': 'Concept', 'label': 'Omitted variable bias (thiên lệch do bỏ biến)'},
    {'id': 'CONCEPT_R_SQUARED_INTERPRETATION', 'type': 'Concept', 'label': 'Diễn giải R² + effect size'},
    {'id': 'CONCEPT_DASS21', 'type': 'Concept', 'label': 'DASS-21 (Lovibond 1995)'},
]

NEW_DOCS = [
    {'id': 'DOC_PHAN_BIEN_VN014_THIEU_ALHT', 'type': 'Doc',
     'label': '01_Bao-cao/Phan_bien_VN014_thieu_yeu_to_ap_luc_hoc_tap.docx'},
]

NEW_QUESTIONS = [
    {'id': 'QA_30', 'type': 'Question',
     'label': 'Phản biện VN014 H.T. Học không nghiên cứu áp lực học tập — hậu quả về R² và diễn giải',
     'topic': 'Statistical interpretation', 'date_asked': '2026-05-09'},
]

NEW_EDGES = [
    ('QA_30', 'TOPIC_STATISTICAL_INTERPRETATION', 'BELONGS_TO'),
    ('QA_30', 'CONCEPT_OMITTED_VARIABLE_BIAS', 'EXPLAINS'),
    ('QA_30', 'CONCEPT_R_SQUARED_INTERPRETATION', 'EXPLAINS'),
    ('QA_30', 'CONCEPT_DASS21', 'EXPLAINS'),
    ('QA_30', 'CONCEPT_ESSA', 'EXPLAINS'),
    ('QA_30', 'PAPER_VN014_HoangTrungHoc_2025', 'CITES'),
    ('QA_30', 'PAPER_QT067', 'CITES'),
    ('QA_30', 'PAPER_Sun_2011_ESSA', 'CITES'),
    ('QA_30', 'DOC_PHAN_BIEN_VN014_THIEU_ALHT', 'ANSWERED_IN'),
    ('PAPER_VN014_HoangTrungHoc_2025', 'CONCEPT_OMITTED_VARIABLE_BIAS', 'RELATED_TO'),
]

NEW_RAG_QUESTIONS = [
    {
        'id': 'QA_30',
        'date_asked': '2026-05-09',
        'question': 'Phản biện VN014 H.T. Học (2025) — không nghiên cứu áp lực học tập? Hậu quả?',
        'short_answer': 'XÁC NHẬN — VN014 KHÔNG nghiên cứu áp lực học tập (verified Tom-tat đoạn [7]). VN014 chỉ đo 6 yếu tố hành vi: thời gian ngủ (β=-0,149), thời gian điện thoại (β=0,176), thể thao (β=-0,087), quan hệ cha mẹ-con (β=0,272), hình thức học tập COVID, giãn cách xã hội + 3 nhân khẩu (giới, khu vực, khối). Đây là HẠN CHẾ NGHIÊM TRỌNG vì: (1) ALHT là yếu tố nguy cơ HÀNG ĐẦU theo Pascoe 2020 khung 6 trục; (2) chương 3 luận án CTH có β ALHT=0,510 (lớn nhất trong YTNC riêng lẻ); (3) R² VN014 chỉ 0,190 (medium) trong khi chương 3 tích hợp R²=0,598 (3,1 lần lớn hơn) — phần lớn do VN014 thiếu ALHT + tự trọng + bắt nạt; (4) diễn giải "quan hệ cha mẹ-con là yếu tố mạnh nhất" có thể bị PHÓNG ĐẠI vì không kiểm soát ALHT. Khuyến nghị: chỉ trích VN014 cho yếu tố đã đo (giấc ngủ, điện thoại, thể thao); KHÔNG trích cho ALHT; có thể trích làm phản biện về thiết kế.',
        'topic': 'Statistical interpretation',
        'concepts': ['Omitted variable bias', 'R²', 'ESSA', 'DASS-21'],
        'doc_path': '01_Bao-cao/Phan_bien_VN014_thieu_yeu_to_ap_luc_hoc_tap.docx',
        'paper_refs': ['VN014', 'QT067'],
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
    for n in NEW_CONCEPTS + NEW_DOCS + NEW_QUESTIONS:
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
