"""Update KG + RAG cho doc khung tap huan v2 — bo sung 4 paper missing."""
import sys, io, json, copy, re
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

LIGHT = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly-light/web/data')
HEAVY = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly/web/data')

NEW_TOPICS = [
    {'id': 'TOPIC_INTERVENTION_DESIGN', 'type': 'Topic', 'label': 'Thiết kế can thiệp phòng ngừa lo âu học đường'},
]

NEW_CONCEPTS = [
    {'id': 'CONCEPT_DMHI', 'type': 'Concept', 'label': 'DMHI (Digital Mental Health Intervention)'},
    {'id': 'CONCEPT_iCBT', 'type': 'Concept', 'label': 'iCBT (internet-delivered CBT)'},
    {'id': 'CONCEPT_MOBILE_CBT', 'type': 'Concept', 'label': 'Mobile CBT'},
    {'id': 'CONCEPT_CLEAR_FEAR_APP', 'type': 'Concept', 'label': 'Clear Fear app — 6 modules CBT'},
    {'id': 'CONCEPT_EACP', 'type': 'Concept', 'label': 'EACP (Early Adolescent Coping Power)'},
    {'id': 'CONCEPT_PSYCHOEDUCATION', 'type': 'Concept', 'label': 'Psychoeducation (tâm lý giáo dục)'},
    {'id': 'CONCEPT_EXPOSURE_LADDER', 'type': 'Concept', 'label': 'Exposure ladder (thang phơi nhiễm)'},
    {'id': 'CONCEPT_COGNITIVE_CHALLENGES', 'type': 'Concept', 'label': 'Cognitive challenges (thử thách suy nghĩ)'},
    {'id': 'CONCEPT_BEHAVIORAL_ACTIVATION', 'type': 'Concept', 'label': 'Behavioral activation (kích hoạt hành vi)'},
    {'id': 'CONCEPT_RELAXATION_TOOLS', 'type': 'Concept', 'label': 'Relaxation tools (thở 4-7-8, grounding 5-4-3-2-1)'},
    {'id': 'CONCEPT_GUIDED_VS_UNGUIDED_DMHI', 'type': 'Concept', 'label': 'Guided vs unguided DMHI'},
    {'id': 'CONCEPT_SAD_SPECIFIC_APP', 'type': 'Concept', 'label': 'SAD-specific app (Walder 2025 g=0,878)'},
    {'id': 'CONCEPT_WHOLE_SCHOOL_APPROACH', 'type': 'Concept', 'label': 'Whole-school approach'},
]

NEW_PAPERS = [
    {'id': 'PAPER_QT050_Qiaochu_2025', 'type': 'Paper', 'label': 'QT050 Qiaochu & Wang (2025) Mobile CBT systematic review'},
    {'id': 'PAPER_QT062_Samele_2025', 'type': 'Paper', 'label': 'QT062 Samele et al. (2025) Clear Fear app evaluation'},
    {'id': 'PAPER_QT065_Lochman_2025', 'type': 'Paper', 'label': 'QT065 Lochman et al. (2025) EACP RCT 40 trường'},
    {'id': 'PAPER_VN003_Pham_2024', 'type': 'Paper', 'label': 'VN003 Pham et al. (2024) Hue social support'},
]

NEW_DOCS = [
    {'id': 'DOC_KHUNG_TAP_HUAN_V2', 'type': 'Doc',
     'label': '01_Bao-cao/bang-so-lieu-binh-luan/00_Khung tập huấn-v2.docx'},
]

NEW_QUESTIONS = [
    {'id': 'QA_29', 'type': 'Question',
     'label': 'Sửa Khung tập huấn 8 chủ đề theo CBT 8-12 buổi + DMHI/iCBT + whole-school + số liệu chương 3',
     'topic': 'Thiết kế can thiệp phòng ngừa lo âu học đường', 'date_asked': '2026-05-09'},
]

NEW_EDGES = [
    ('QA_29', 'TOPIC_INTERVENTION_DESIGN', 'BELONGS_TO'),
    ('QA_29', 'CONCEPT_DMHI', 'EXPLAINS'),
    ('QA_29', 'CONCEPT_iCBT', 'EXPLAINS'),
    ('QA_29', 'CONCEPT_CLEAR_FEAR_APP', 'EXPLAINS'),
    ('QA_29', 'CONCEPT_EACP', 'EXPLAINS'),
    ('QA_29', 'CONCEPT_WHOLE_SCHOOL_APPROACH', 'EXPLAINS'),
    ('QA_29', 'CONCEPT_GUIDED_VS_UNGUIDED_DMHI', 'EXPLAINS'),
    ('QA_29', 'PAPER_QT040', 'CITES'),
    ('QA_29', 'PAPER_QT050_Qiaochu_2025', 'CITES'),
    ('QA_29', 'PAPER_QT062_Samele_2025', 'CITES'),
    ('QA_29', 'PAPER_QT065_Lochman_2025', 'CITES'),
    ('QA_29', 'PAPER_QT066', 'CITES'),
    ('QA_29', 'PAPER_QT033_SchmidtPersson_2024', 'CITES'),
    ('QA_29', 'PAPER_VN014_HoangTrungHoc_2025', 'CITES'),
    ('QA_29', 'PAPER_VN004_NguyenThiVan_2020', 'CITES'),
    ('QA_29', 'PAPER_VN003_Pham_2024', 'CITES'),
    ('QA_29', 'DOC_KHUNG_TAP_HUAN_V2', 'ANSWERED_IN'),
    # Concept-Topic
    ('CONCEPT_DMHI', 'TOPIC_INTERVENTION_DESIGN', 'BELONGS_TO'),
    ('CONCEPT_iCBT', 'TOPIC_INTERVENTION_DESIGN', 'BELONGS_TO'),
    ('CONCEPT_MOBILE_CBT', 'TOPIC_INTERVENTION_DESIGN', 'BELONGS_TO'),
    ('CONCEPT_EACP', 'TOPIC_INTERVENTION_DESIGN', 'BELONGS_TO'),
    ('CONCEPT_WHOLE_SCHOOL_APPROACH', 'TOPIC_INTERVENTION_DESIGN', 'BELONGS_TO'),
    # Paper-Concept
    ('PAPER_QT040', 'CONCEPT_DMHI', 'RELATED_TO'),
    ('PAPER_QT040', 'CONCEPT_iCBT', 'RELATED_TO'),
    ('PAPER_QT040', 'CONCEPT_SAD_SPECIFIC_APP', 'RELATED_TO'),
    ('PAPER_QT050_Qiaochu_2025', 'CONCEPT_MOBILE_CBT', 'RELATED_TO'),
    ('PAPER_QT062_Samele_2025', 'CONCEPT_CLEAR_FEAR_APP', 'RELATED_TO'),
    ('PAPER_QT062_Samele_2025', 'CONCEPT_EXPOSURE_LADDER', 'RELATED_TO'),
    ('PAPER_QT065_Lochman_2025', 'CONCEPT_EACP', 'RELATED_TO'),
    ('PAPER_QT065_Lochman_2025', 'CONCEPT_WHOLE_SCHOOL_APPROACH', 'RELATED_TO'),
    # Doc-Paper
    ('DOC_KHUNG_TAP_HUAN_V2', 'PAPER_QT040', 'CITES'),
    ('DOC_KHUNG_TAP_HUAN_V2', 'PAPER_QT050_Qiaochu_2025', 'CITES'),
    ('DOC_KHUNG_TAP_HUAN_V2', 'PAPER_QT062_Samele_2025', 'CITES'),
    ('DOC_KHUNG_TAP_HUAN_V2', 'PAPER_QT065_Lochman_2025', 'CITES'),
    ('DOC_KHUNG_TAP_HUAN_V2', 'PAPER_QT066', 'CITES'),
]

NEW_RAG_QUESTIONS = [
    {
        'id': 'QA_29',
        'date_asked': '2026-05-09',
        'question': 'Sửa Khung tập huấn 8 chủ đề theo CBT 8-12 buổi + DMHI/iCBT + whole-school approach + số liệu chương 3',
        'short_answer': 'Khung v2 (48 KB) bổ sung phần XANH dựa trên: (1) Walder 2025 (QT040) DMHI MA 21 RCT g=0,508 chung / g=0,878 SAD-specific / g=0,825 guided; (2) Qiaochu 2025 (QT050) Mobile CBT 9 RCT — yếu cho lo âu (chỉ 2/6 sig); (3) Samele 2025 (QT062) Clear Fear app 6 modules (psychoeducation, symptom tracking, cognitive challenges, relaxation, behavioral activation, exposure ladder); (4) Lochman 2025 (QT065) EACP cluster RCT 40 trường — 25 buổi HS + 16 buổi cha mẹ; (5) Cai 2025 school-based resilience SMD=0,17; (6) Murphy 2024 (QT066) peer support; (7) số liệu chương 3 (β ALHT 0,510/0,490; β TTr -0,455/-0,415; β BPĐP 0,749 fit kém; F gender 44,484/45,984/0,246; ước lượng % ngưỡng 65: TQ 33,85%, XH 25,98%, CL 5,01%). Cấu trúc 8 chủ đề × 19 buổi (8 LT + 11 TH) phù hợp 8-12 buổi CBT châu Á.',
        'topic': 'Thiết kế can thiệp phòng ngừa lo âu học đường',
        'concepts': ['DMHI', 'iCBT', 'Mobile CBT', 'Clear Fear app', 'EACP', 'Whole-school approach'],
        'doc_path': '01_Bao-cao/bang-so-lieu-binh-luan/00_Khung tập huấn-v2.docx',
        'paper_refs': ['QT040', 'QT050', 'QT062', 'QT065', 'QT066', 'QT033', 'VN014', 'VN004', 'VN003'],
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
    for n in NEW_TOPICS + NEW_CONCEPTS + NEW_PAPERS + NEW_DOCS + NEW_QUESTIONS:
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
