"""Update KG + RAG cho doc phan bien VN004 yeu to hoc tap."""
import sys, io, json, copy, re
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

LIGHT = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly-light/web/data')
HEAVY = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly/web/data')

NEW_CONCEPTS = [
    {'id': 'CONCEPT_ESSA', 'type': 'Concept', 'label': 'ESSA (Educational Stress Scale for Adolescents, Sun 2011)'},
    {'id': 'CONCEPT_CTAS', 'type': 'Concept', 'label': 'CTAS (Cognitive Test Anxiety Scale, Cassady & Johnson 2002)'},
    {'id': 'CONCEPT_TAI', 'type': 'Concept', 'label': 'TAI (Test Anxiety Inventory, Spielberger 1980)'},
    {'id': 'CONCEPT_TEST_ANXIETY', 'type': 'Concept', 'label': 'Test anxiety (lo âu thi cử)'},
    {'id': 'CONCEPT_ACADEMIC_PRESSURE_DIMENSIONS', 'type': 'Concept', 'label': 'Sáu chiều con của áp lực học tập'},
]

NEW_PAPERS = [
    {'id': 'PAPER_Sun_2011_ESSA', 'type': 'Paper', 'label': 'Sun et al. (2011) ESSA J Psychoeduc Assess'},
    {'id': 'PAPER_Cassady_2002_CTAS', 'type': 'Paper', 'label': 'Cassady & Johnson (2002) CTAS Contemp Educ Psychol'},
    {'id': 'PAPER_Spielberger_1980_TAI', 'type': 'Paper', 'label': 'Spielberger (1980) TAI manual'},
    {'id': 'PAPER_Putwain_2008', 'type': 'Paper', 'label': 'Putwain (2008) The Psychologist examination stress'},
]

NEW_DOCS = [
    {'id': 'DOC_PHAN_BIEN_VN004_HOC_TAP', 'type': 'Doc',
     'label': '01_Bao-cao/Phan_bien_VN004_yeu_to_hoc_tap.docx'},
]

NEW_QUESTIONS = [
    {'id': 'QA_28', 'type': 'Question',
     'label': 'Phản biện VN004 yếu tố HỌC TẬP — khái niệm rộng quá, không phân chiều, theo style đã làm với VN014',
     'topic': 'Statistical interpretation', 'date_asked': '2026-05-09'},
]

NEW_EDGES = [
    ('QA_28', 'TOPIC_STATISTICAL_INTERPRETATION', 'BELONGS_TO'),
    ('QA_28', 'CONCEPT_ESSA', 'EXPLAINS'),
    ('QA_28', 'CONCEPT_CTAS', 'EXPLAINS'),
    ('QA_28', 'CONCEPT_TAI', 'EXPLAINS'),
    ('QA_28', 'CONCEPT_TEST_ANXIETY', 'EXPLAINS'),
    ('QA_28', 'CONCEPT_ACADEMIC_PRESSURE_DIMENSIONS', 'EXPLAINS'),
    ('QA_28', 'PAPER_VN004_NguyenThiVan_2020', 'CITES'),
    ('QA_28', 'PAPER_Sun_2011_ESSA', 'CITES'),
    ('QA_28', 'PAPER_Cassady_2002_CTAS', 'CITES'),
    ('QA_28', 'PAPER_Spielberger_1980_TAI', 'CITES'),
    ('QA_28', 'PAPER_Putwain_2008', 'CITES'),
    ('QA_28', 'PAPER_QT067', 'CITES'),
    ('QA_28', 'DOC_PHAN_BIEN_VN004_HOC_TAP', 'ANSWERED_IN'),
    # Paper-Concept
    ('PAPER_Sun_2011_ESSA', 'CONCEPT_ESSA', 'RELATED_TO'),
    ('PAPER_Cassady_2002_CTAS', 'CONCEPT_CTAS', 'RELATED_TO'),
    ('PAPER_Spielberger_1980_TAI', 'CONCEPT_TAI', 'RELATED_TO'),
    ('PAPER_VN004_NguyenThiVan_2020', 'CONCEPT_ACADEMIC_PRESSURE_DIMENSIONS', 'RELATED_TO'),
    # Doc-Paper
    ('DOC_PHAN_BIEN_VN004_HOC_TAP', 'PAPER_VN004_NguyenThiVan_2020', 'CITES'),
    ('DOC_PHAN_BIEN_VN004_HOC_TAP', 'PAPER_Sun_2011_ESSA', 'CITES'),
    ('DOC_PHAN_BIEN_VN004_HOC_TAP', 'PAPER_Cassady_2002_CTAS', 'CITES'),
    ('DOC_PHAN_BIEN_VN004_HOC_TAP', 'PAPER_Spielberger_1980_TAI', 'CITES'),
    ('DOC_PHAN_BIEN_VN004_HOC_TAP', 'PAPER_Putwain_2008', 'CITES'),
    ('DOC_PHAN_BIEN_VN004_HOC_TAP', 'PAPER_QT067', 'CITES'),
]

NEW_RAG_QUESTIONS = [
    {
        'id': 'QA_28',
        'date_asked': '2026-05-09',
        'question': 'Phản biện yếu tố HỌC TẬP của VN004 Nguyễn Thị Vân (2020) — khái niệm có rộng quá không, theo style đã làm với VN014?',
        'short_answer': 'CÓ — khái niệm "yếu tố học tập" trong VN004 RỘNG QUÁ. Tác giả đo r=0,37 với lo âu bằng MỘT biến tổng — không phân biệt 6 chiều con (áp lực thi / khối lượng bài tập / lo lắng điểm / kỳ vọng tự thân / kỳ vọng cha mẹ / định hướng tương lai). Top 3 biểu hiện (áp lực thi ĐH 56,7%; định hướng nghề 51,5%; kỳ vọng cha mẹ 48,9%) chỉ là TẦN SUẤT, không phải PREDICTOR. Trái với VN004, Trần Thảo Vi (2024) trong J Rural Medicine dùng ESSA 16 mục 5 chiều (Sun 2011); chương 3 luận án dùng ESSA rút gọn 4 mục — tốt hơn VN004 nhưng vẫn yếu hơn ESSA full. Bốn hậu quả: (1) không phân biệt aspect mạnh nhất; (2) tần suất ≠ predictor; (3) không so sánh quốc tế; (4) không can thiệp đích. Đề xuất: dùng ESSA full 16 mục hoặc CTAS (Cassady 2002) hoặc TAI (Spielberger 1980); phân tích RIÊNG từng chiều với SEM.',
        'topic': 'Statistical interpretation',
        'concepts': ['ESSA', 'CTAS', 'TAI', 'Test anxiety', 'Academic pressure dimensions'],
        'doc_path': '01_Bao-cao/Phan_bien_VN004_yeu_to_hoc_tap.docx',
        'paper_refs': ['VN004', 'QT067'],
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
    print(f'  KG: +{na} nodes, +{ea} edges; total {len(kg["nodes"])}/{len(kg["edges"])}')

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


for label, dir_ in [('LIGHT',LIGHT),('HEAVY',HEAVY)]:
    print(f'\n==== {label} ====')
    if not dir_.exists(): print('missing'); continue
    update(dir_)
print('\nDone.')
