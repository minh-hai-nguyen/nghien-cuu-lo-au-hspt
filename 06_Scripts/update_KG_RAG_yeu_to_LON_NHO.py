"""Update KG + RAG voi 5 doc moi (yeu to LON + 4 yeu to NHO)."""
import sys, io, json, copy, re
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

LIGHT = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly-light/web/data')
HEAVY = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly/web/data')

NEW_TOPICS = [
    {'id': 'TOPIC_RISK_PROTECTIVE_MODEL', 'type': 'Topic', 'label': 'Risk-protective factor model'},
    {'id': 'TOPIC_COPING', 'type': 'Topic', 'label': 'Coping strategies'},
]

NEW_CONCEPTS = [
    # YTNC + YTBV
    {'id': 'CONCEPT_YTNC_TONG', 'type': 'Concept', 'label': 'Yếu tố nguy cơ tổng (YTNC)'},
    {'id': 'CONCEPT_YTBV_TONG', 'type': 'Concept', 'label': 'Yếu tố bảo vệ tổng (YTBV)'},
    # GBTH
    {'id': 'CONCEPT_GBTH_SCHOOL_CONNECTEDNESS', 'type': 'Concept', 'label': 'GBTH (school connectedness)'},
    {'id': 'CONCEPT_PSSM', 'type': 'Concept', 'label': 'PSSM (Psychological Sense of School Membership, Goodenow 1993)'},
    {'id': 'CONCEPT_TEACHER_STUDENT_RELATION', 'type': 'Concept', 'label': 'Quan hệ giáo viên-học sinh'},
    # HTCM, HTBB, MSPSS
    {'id': 'CONCEPT_HTCM_PARENT_SUPPORT', 'type': 'Concept', 'label': 'HTCM (parent support)'},
    {'id': 'CONCEPT_HTBB_PEER_SUPPORT', 'type': 'Concept', 'label': 'HTBB (peer support)'},
    {'id': 'CONCEPT_MSPSS', 'type': 'Concept', 'label': 'MSPSS (Multidimensional Scale Perceived Social Support, Zimet 1988)'},
    # BPĐP
    {'id': 'CONCEPT_BPDP_COPING', 'type': 'Concept', 'label': 'BPĐP (biện pháp đối phó)'},
    {'id': 'CONCEPT_BRIEF_COPE', 'type': 'Concept', 'label': 'Brief-COPE (Carver 1997, 14 nhân tố)'},
    {'id': 'CONCEPT_MALADAPTIVE_COPING_ESCALATION', 'type': 'Concept', 'label': 'Maladaptive coping escalation (Compas 2017)'},
    {'id': 'CONCEPT_AVOIDANCE_COPING', 'type': 'Concept', 'label': 'Avoidance coping (tránh né)'},
    {'id': 'CONCEPT_PROBLEM_FOCUSED_COPING', 'type': 'Concept', 'label': 'Problem-focused coping'},
    {'id': 'CONCEPT_SUPPORT_SEEKING', 'type': 'Concept', 'label': 'Support-seeking coping'},
]

NEW_PAPERS = [
    {'id': 'PAPER_Goodenow_1993_PSSM', 'type': 'Paper', 'label': 'Goodenow (1993) PSSM scale'},
    {'id': 'PAPER_Zimet_1988_MSPSS', 'type': 'Paper', 'label': 'Zimet et al. (1988) MSPSS'},
    {'id': 'PAPER_Carver_1997_Brief_COPE', 'type': 'Paper', 'label': 'Carver (1997) Brief-COPE'},
]

NEW_DOCS = [
    {'id': 'DOC_DG_YTNC_YTBV_TICH_HOP', 'type': 'Doc', 'label': '01_Bao-cao/Dien_giai_yeu_to_LON_YTNC_YTBV_mo_hinh_tich_hop.docx'},
    {'id': 'DOC_DG_GBTH', 'type': 'Doc', 'label': '01_Bao-cao/Dien_giai_yeu_to_GBTH_gan_bo_truong_hoc.docx'},
    {'id': 'DOC_DG_HTCM', 'type': 'Doc', 'label': '01_Bao-cao/Dien_giai_yeu_to_HTCM_ho_tro_cha_me.docx'},
    {'id': 'DOC_DG_HTBB', 'type': 'Doc', 'label': '01_Bao-cao/Dien_giai_yeu_to_HTBB_ho_tro_ban_be.docx'},
    {'id': 'DOC_DG_BPDP', 'type': 'Doc', 'label': '01_Bao-cao/Dien_giai_yeu_to_BPDP_bien_phap_doi_pho.docx'},
]

NEW_QUESTIONS = [
    {'id': 'QA_25', 'type': 'Question',
     'label': 'Diễn giải yếu tố LỚN — YTNC + YTBV + Mô hình tích hợp (chương 3)',
     'topic': 'Risk-protective factor model', 'date_asked': '2026-05-08'},
    {'id': 'QA_26', 'type': 'Question',
     'label': 'Diễn giải các yếu tố NHỎ còn lại: GBTH, HTCM, HTBB, BPĐP',
     'topic': 'Risk-protective factor model', 'date_asked': '2026-05-08'},
]

NEW_EDGES = [
    # Topic links
    ('QA_25', 'TOPIC_RISK_PROTECTIVE_MODEL', 'BELONGS_TO'),
    ('QA_26', 'TOPIC_RISK_PROTECTIVE_MODEL', 'BELONGS_TO'),
    # QA_25 — yếu tố lớn
    ('QA_25', 'CONCEPT_YTNC_TONG', 'EXPLAINS'),
    ('QA_25', 'CONCEPT_YTBV_TONG', 'EXPLAINS'),
    ('QA_25', 'DOC_DG_YTNC_YTBV_TICH_HOP', 'ANSWERED_IN'),
    # QA_26 — 4 yếu tố nhỏ
    ('QA_26', 'CONCEPT_GBTH_SCHOOL_CONNECTEDNESS', 'EXPLAINS'),
    ('QA_26', 'CONCEPT_HTCM_PARENT_SUPPORT', 'EXPLAINS'),
    ('QA_26', 'CONCEPT_HTBB_PEER_SUPPORT', 'EXPLAINS'),
    ('QA_26', 'CONCEPT_BPDP_COPING', 'EXPLAINS'),
    ('QA_26', 'CONCEPT_PSSM', 'EXPLAINS'),
    ('QA_26', 'CONCEPT_MSPSS', 'EXPLAINS'),
    ('QA_26', 'CONCEPT_BRIEF_COPE', 'EXPLAINS'),
    ('QA_26', 'CONCEPT_MALADAPTIVE_COPING_ESCALATION', 'EXPLAINS'),
    ('QA_26', 'DOC_DG_GBTH', 'ANSWERED_IN'),
    ('QA_26', 'DOC_DG_HTCM', 'ANSWERED_IN'),
    ('QA_26', 'DOC_DG_HTBB', 'ANSWERED_IN'),
    ('QA_26', 'DOC_DG_BPDP', 'ANSWERED_IN'),
    # Concept-Topic
    ('CONCEPT_YTNC_TONG', 'TOPIC_RISK_PROTECTIVE_MODEL', 'BELONGS_TO'),
    ('CONCEPT_YTBV_TONG', 'TOPIC_RISK_PROTECTIVE_MODEL', 'BELONGS_TO'),
    ('CONCEPT_BPDP_COPING', 'TOPIC_COPING', 'BELONGS_TO'),
    ('CONCEPT_BRIEF_COPE', 'TOPIC_COPING', 'BELONGS_TO'),
    ('CONCEPT_AVOIDANCE_COPING', 'TOPIC_COPING', 'BELONGS_TO'),
    ('CONCEPT_PROBLEM_FOCUSED_COPING', 'TOPIC_COPING', 'BELONGS_TO'),
    ('CONCEPT_SUPPORT_SEEKING', 'TOPIC_COPING', 'BELONGS_TO'),
    # Paper-Concept
    ('PAPER_Goodenow_1993_PSSM', 'CONCEPT_PSSM', 'RELATED_TO'),
    ('PAPER_Goodenow_1993_PSSM', 'CONCEPT_GBTH_SCHOOL_CONNECTEDNESS', 'RELATED_TO'),
    ('PAPER_Zimet_1988_MSPSS', 'CONCEPT_MSPSS', 'RELATED_TO'),
    ('PAPER_Zimet_1988_MSPSS', 'CONCEPT_HTCM_PARENT_SUPPORT', 'RELATED_TO'),
    ('PAPER_Zimet_1988_MSPSS', 'CONCEPT_HTBB_PEER_SUPPORT', 'RELATED_TO'),
    ('PAPER_Carver_1997_Brief_COPE', 'CONCEPT_BRIEF_COPE', 'RELATED_TO'),
    ('PAPER_Compas_2017_Psych_Bulletin', 'CONCEPT_MALADAPTIVE_COPING_ESCALATION', 'RELATED_TO'),
    # Doc-Paper
    ('DOC_DG_GBTH', 'PAPER_Goodenow_1993_PSSM', 'CITES'),
    ('DOC_DG_HTCM', 'PAPER_Zimet_1988_MSPSS', 'CITES'),
    ('DOC_DG_HTCM', 'PAPER_VN003', 'CITES'),
    ('DOC_DG_HTBB', 'PAPER_Zimet_1988_MSPSS', 'CITES'),
    ('DOC_DG_HTBB', 'PAPER_QT027_Fassi_2025', 'CITES'),
    ('DOC_DG_HTBB', 'PAPER_QT066', 'CITES'),
    ('DOC_DG_BPDP', 'PAPER_Carver_1997_Brief_COPE', 'CITES'),
    ('DOC_DG_BPDP', 'PAPER_Compas_2017_Psych_Bulletin', 'CITES'),
]

NEW_RAG_QUESTIONS = [
    {
        'id': 'QA_25',
        'date_asked': '2026-05-08',
        'question': 'Diễn giải yếu tố LỚN — YTNC tổng + YTBV tổng + Mô hình tích hợp (Bảng 3.37–3.42 chương 3)',
        'short_answer': 'YTNC riêng (Bảng 3.39-3.40): β=0,747; R²=0,558; CFI=0,994; RMSEA=0,039 (XUẤT SẮC). YTBV riêng (Bảng 3.41-3.42): β=−0,352; R²=0,124; CFI=0,982; RMSEA=0,070 (chấp nhận). Mô hình tích hợp YTNC+YTBV (Bảng 3.37-3.38): β YTNC=0,669; β YTBV=−0,220; R²=0,598; CFI=0,937; RMSEA=0,077 (acceptable). YTNC giải thích 55,8% RLLA; YTBV chỉ 12,4%. |β YTBV|/|β YTNC| ≈ 0,47 — bảo vệ 47% nguy cơ. Khi tích hợp, β YTNC giảm 10%, β YTBV giảm 37% — gợi ý YTBV không hoàn toàn độc lập với YTNC.',
        'topic': 'Risk-protective factor model',
        'concepts': ['YTNC tổng', 'YTBV tổng', 'SEM', 'R²', 'Risk-protective model'],
        'doc_path': '01_Bao-cao/Dien_giai_yeu_to_LON_YTNC_YTBV_mo_hinh_tich_hop.docx',
        'paper_refs': [],
    },
    {
        'id': 'QA_26',
        'date_asked': '2026-05-08',
        'question': 'Diễn giải các yếu tố NHỎ: GBTH (gắn bó trường học), HTCM (hỗ trợ cha mẹ), HTBB (hỗ trợ bạn bè), BPĐP (biện pháp đối phó)',
        'short_answer': 'GBTH (PSSM): β → RLLAXH=−0,187 mạnh nhất; β → RLLATQ=−0,108; β → RLLACL=0,014 (NS). R²=2,4%. ĐTB tổng=52,60; điểm yếu là quan hệ GV-HS (PSSM.5=47,36). HTCM (MSPSS): β → RLLAXH=−0,273 mạnh nhất; β → RLLATQ=−0,172; β → RLLACL=0,000 (NS — KHÔNG tác động!). R²=5,5%. ĐTB tổng=57,65 cao nhất YTBV. HTBB (MSPSS): GẦN NHƯ KHÔNG có tác động — tất cả β NS hoặc < 0,10. R²=0,2%. Phát hiện CRITICAL trái y văn quốc tế. BPĐP (Brief-COPE): β CỰC LỚN → RLLATQ=0,749; → RLLAXH=0,670; → RLLACL=0,132. R² lớn (44,9-56,1%) NHƯNG fit KÉM (RMSEA 0,080-0,204; CFI 0,865-0,911). Maladaptive coping escalation (Compas 2017). KHÔNG báo cáo β=0,749 làm phát hiện chính.',
        'topic': 'Risk-protective factor model',
        'concepts': ['GBTH', 'HTCM', 'HTBB', 'BPĐP', 'PSSM', 'MSPSS', 'Brief-COPE', 'Maladaptive coping escalation'],
        'doc_path': '01_Bao-cao/Dien_giai_yeu_to_GBTH_gan_bo_truong_hoc.docx; .../HTCM...; .../HTBB...; .../BPDP...',
        'paper_refs': ['VN003', 'QT027', 'QT066'],
    },
]


def tokenize(text):
    text = text.lower()
    tokens = re.findall(r'[\w]+', text, re.UNICODE)
    counts = {}
    for t in tokens:
        if len(t) >= 2:
            counts[t] = counts.get(t, 0) + 1
    return counts


def update_kg(data_dir):
    path = data_dir / 'questions_kg.json'
    with open(path, encoding='utf-8') as f: kg = json.load(f)
    eids = {n['id'] for n in kg['nodes']}
    eedges = {(e['source'], e['target'], e['relation']) for e in kg['edges']}
    na = 0
    for n in NEW_TOPICS + NEW_CONCEPTS + NEW_PAPERS + NEW_DOCS + NEW_QUESTIONS:
        if n['id'] not in eids:
            kg['nodes'].append(n); eids.add(n['id']); na += 1
    ea = 0
    for src, tgt, rel in NEW_EDGES:
        if (src, tgt, rel) in eedges or src not in eids or tgt not in eids: continue
        kg['edges'].append({'source': src, 'target': tgt, 'relation': rel})
        eedges.add((src, tgt, rel)); ea += 1
    from collections import Counter
    kg['meta']['n_nodes'] = len(kg['nodes'])
    kg['meta']['n_edges'] = len(kg['edges'])
    kg['meta']['node_types'] = dict(Counter(n['type'] for n in kg['nodes']))
    kg['meta']['edge_relations'] = dict(Counter(e['relation'] for e in kg['edges']))
    kg['meta']['updated'] = '2026-05-08'
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(kg, f, ensure_ascii=False, indent=2)
    print(f'  KG: +{na} nodes, +{ea} edges; total {len(kg["nodes"])} / {len(kg["edges"])}')


def update_rag(data_dir):
    path = data_dir / 'rag_questions_index.json'
    with open(path, encoding='utf-8') as f: rqi = json.load(f)
    eids = {e['id'] for e in rqi['entries']}
    added = 0
    for nq in NEW_RAG_QUESTIONS:
        if nq['id'] in eids: continue
        e = copy.deepcopy(nq)
        e['tokens'] = tokenize(e['question'] + ' ' + e['short_answer'])
        rqi['entries'].append(e); added += 1
    rqi['meta']['n_entries'] = len(rqi['entries'])
    rqi['meta']['updated'] = '2026-05-08'
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(rqi, f, ensure_ascii=False, indent=2)
    print(f'  rag_questions: +{added}; total {len(rqi["entries"])}')


def main():
    for label, dir_ in [('LIGHT', LIGHT), ('HEAVY', HEAVY)]:
        print(f'\n==== {label} ====')
        if not dir_.exists(): print(f'  ⚠ missing'); continue
        update_kg(dir_); update_rag(dir_)
    print('\nDone.')


if __name__ == '__main__': main()
