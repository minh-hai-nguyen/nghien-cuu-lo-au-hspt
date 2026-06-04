"""Update KG + RAG cho doc danh gia ban luan + cach quy ve %."""
import sys, io, json, copy, re
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

LIGHT = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly-light/web/data')
HEAVY = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly/web/data')

NEW_CONCEPTS = [
    {'id': 'CONCEPT_PERCENT_CONVERSION', 'type': 'Concept', 'label': 'Quy đổi điểm số sang tỷ lệ %'},
    {'id': 'CONCEPT_RCADS_CUTOFF', 'type': 'Concept', 'label': 'RCADS cutoff (T-score 65 borderline / 70 clinical)'},
    {'id': 'CONCEPT_Z_SCORE', 'type': 'Concept', 'label': 'Z-score (chuẩn hóa thống kê)'},
    {'id': 'CONCEPT_NORMAL_DISTRIBUTION', 'type': 'Concept', 'label': 'Phân phối chuẩn (normal distribution)'},
    {'id': 'CONCEPT_SKEWED_DISTRIBUTION', 'type': 'Concept', 'label': 'Phân phối lệch (skewed)'},
    {'id': 'CONCEPT_PREVALENCE', 'type': 'Concept', 'label': 'Prevalence (tỷ lệ hiện mắc)'},
    {'id': 'CONCEPT_CONTINUOUS_VS_CATEGORICAL', 'type': 'Concept', 'label': 'Thang đo liên tục vs phân loại'},
]

NEW_DOCS = [
    {'id': 'DOC_DANH_GIA_BANLUAN_QUY_PHAN_TRAM', 'type': 'Doc',
     'label': '01_Bao-cao/Danh_gia_banluan_3_6_va_cach_quy_ve_phan_tram.docx'},
]

NEW_QUESTIONS = [
    {'id': 'QA_27', 'type': 'Question',
     'label': 'Đánh giá phần Bàn luận 3.6 của các bạn trợ lý + cách quy điểm số RCADS sang tỷ lệ %',
     'topic': 'Statistical interpretation', 'date_asked': '2026-05-09'},
]

NEW_EDGES = [
    ('QA_27', 'TOPIC_STATISTICAL_INTERPRETATION', 'BELONGS_TO'),
    ('QA_27', 'CONCEPT_PERCENT_CONVERSION', 'EXPLAINS'),
    ('QA_27', 'CONCEPT_RCADS_CUTOFF', 'EXPLAINS'),
    ('QA_27', 'CONCEPT_Z_SCORE', 'EXPLAINS'),
    ('QA_27', 'CONCEPT_NORMAL_DISTRIBUTION', 'EXPLAINS'),
    ('QA_27', 'CONCEPT_PREVALENCE', 'EXPLAINS'),
    ('QA_27', 'CONCEPT_CONTINUOUS_VS_CATEGORICAL', 'EXPLAINS'),
    ('QA_27', 'CONCEPT_RCADS', 'EXPLAINS'),
    ('QA_27', 'PAPER_VN002', 'CITES'),
    ('QA_27', 'PAPER_VN014_HoangTrungHoc_2025', 'CITES'),
    ('QA_27', 'DOC_DANH_GIA_BANLUAN_QUY_PHAN_TRAM', 'ANSWERED_IN'),
]

NEW_RAG_QUESTIONS = [
    {
        'id': 'QA_27',
        'date_asked': '2026-05-09',
        'question': 'Đánh giá phần Bàn luận 3.6 của các bạn trợ lý + tại sao số liệu RCADS không quy ra %, cách quy về %?',
        'short_answer': 'Đánh giá: phần bàn luận có cấu trúc tốt (cite Ngô Anh Vinh 2022, Lê T.P. Thanh 2024, Steare 2023, Carver 1997, Li 2025, ...) nhưng có 6 thiếu sót: (1) năm Nguyễn Thị Vân SAI = 2018 → đúng 2020; (2) BỎ QUA phát hiện đặc biệt BNHĐ → RLLACL β=0,376 mạnh nhất; (3) BỎ QUA cảnh báo fit indices KÉM của BPĐP (RMSEA 0,080-0,204); (4) KHÔNG nêu pattern ba tầng chênh lệch giới (RLLAC NS); (5) thiếu bảng so sánh cường độ; (6) KHÔNG quy ra %. Tại sao không có %: RCADS là thang đo MỨC ĐỘ liên tục, không phải chẩn đoán nhị phân — cần ngưỡng cắt cụ thể để quy ra %. Cách quy về %: (a) ngưỡng RCADS chuẩn quốc tế (T-score 65 borderline; 70 clinical); (b) z-score giả định normal — RLLATQ ~33,85%; RLLAXH ~25,98%; RLLAC ~5,01% (ngưỡng 65); (c) liên hệ tác giả xin data raw. Cảnh báo: dữ liệu RLLA thường lệch phải, ước lượng z-score có thể cao hơn thực tế.',
        'topic': 'Statistical interpretation',
        'concepts': ['RCADS cutoff', 'Z-score', 'Prevalence', 'Normal distribution', 'Continuous vs categorical'],
        'doc_path': '01_Bao-cao/Danh_gia_banluan_3_6_va_cach_quy_ve_phan_tram.docx',
        'paper_refs': ['VN002', 'VN014'],
    },
]


def tokenize(text):
    text = text.lower()
    counts = {}
    for t in re.findall(r'[\w]+', text, re.UNICODE):
        if len(t) >= 2:
            counts[t] = counts.get(t, 0) + 1
    return counts


def update(data_dir):
    # KG
    path = data_dir / 'questions_kg.json'
    with open(path, encoding='utf-8') as f: kg = json.load(f)
    eids = {n['id'] for n in kg['nodes']}
    eedges = {(e['source'], e['target'], e['relation']) for e in kg['edges']}
    na = 0
    for n in NEW_CONCEPTS + NEW_DOCS + NEW_QUESTIONS:
        if n['id'] not in eids: kg['nodes'].append(n); eids.add(n['id']); na += 1
    ea = 0
    for src, tgt, rel in NEW_EDGES:
        if (src,tgt,rel) in eedges or src not in eids or tgt not in eids: continue
        kg['edges'].append({'source':src,'target':tgt,'relation':rel}); eedges.add((src,tgt,rel)); ea += 1
    from collections import Counter
    kg['meta']['n_nodes']=len(kg['nodes']); kg['meta']['n_edges']=len(kg['edges'])
    kg['meta']['node_types']=dict(Counter(n['type'] for n in kg['nodes']))
    kg['meta']['edge_relations']=dict(Counter(e['relation'] for e in kg['edges']))
    kg['meta']['updated']='2026-05-09'
    with open(path,'w',encoding='utf-8') as f: json.dump(kg,f,ensure_ascii=False,indent=2)
    print(f'  KG: +{na} nodes, +{ea} edges; total {len(kg["nodes"])}/{len(kg["edges"])}')

    # RAG
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
    print(f'  rag_questions: +{added}; total {len(rqi["entries"])}')


for label, dir_ in [('LIGHT',LIGHT), ('HEAVY',HEAVY)]:
    print(f'\n==== {label} ====')
    if not dir_.exists(): print('missing'); continue
    update(dir_)
print('\nDone.')
