"""Cap nhat KG + RAG cho 2 doc moi: H4 gender + BNHD.

THEM VAO questions_kg.json:
- 2 Topic: GENDER_PSYCHOPATHOLOGY, SCHOOL_BULLYING
- 9 Concept: GENDER_DIFFERENCES, SEX_RATIO, HPG_AXIS, RUMINATION,
  SCHOOL_BULLYING_CONCEPT, OBVQ, SCHOOL_REFUSAL, SEPARATION_ANXIETY, SAD_GAD
- 4 Paper: QT009 (missing!), Salk_2017, McLean_2011, Olweus_1996
- 2 Doc: DOC_H4_GENDER, DOC_BNHD
- 2 Question: QA_23 (H4), QA_24 (BNHD)

THEM VAO rag_questions_index.json: QA_23, QA_24
"""
import sys, io, json, copy, re
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

LIGHT = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly-light/web/data')
HEAVY = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly/web/data')

NEW_TOPICS = [
    {'id': 'TOPIC_GENDER_PSYCHOPATHOLOGY', 'type': 'Topic', 'label': 'Gender differences in psychopathology'},
    {'id': 'TOPIC_SCHOOL_BULLYING', 'type': 'Topic', 'label': 'School bullying'},
]

NEW_CONCEPTS = [
    {'id': 'CONCEPT_GENDER_DIFFERENCES', 'type': 'Concept', 'label': 'Gender differences in anxiety'},
    {'id': 'CONCEPT_SEX_RATIO', 'type': 'Concept', 'label': 'Sex ratio (female:male)'},
    {'id': 'CONCEPT_HPG_AXIS', 'type': 'Concept', 'label': 'HPG axis (hypothalamic-pituitary-gonadal)'},
    {'id': 'CONCEPT_RUMINATION', 'type': 'Concept', 'label': 'Rumination (Nolen-Hoeksema)'},
    {'id': 'CONCEPT_SCHOOL_BULLYING', 'type': 'Concept', 'label': 'School bullying / peer victimization'},
    {'id': 'CONCEPT_OBVQ', 'type': 'Concept', 'label': 'OBVQ (Olweus Bully/Victim Questionnaire)'},
    {'id': 'CONCEPT_SCHOOL_REFUSAL', 'type': 'Concept', 'label': 'School refusal'},
    {'id': 'CONCEPT_SEPARATION_ANXIETY', 'type': 'Concept', 'label': 'Separation Anxiety Disorder (RLLAC)'},
    {'id': 'CONCEPT_SOCIAL_ANXIETY_DISORDER', 'type': 'Concept', 'label': 'Social Anxiety Disorder (SAD/RLLAXH)'},
    {'id': 'CONCEPT_GENERALIZED_ANXIETY_DISORDER', 'type': 'Concept', 'label': 'Generalized Anxiety Disorder (GAD/RLLATQ)'},
    {'id': 'CONCEPT_OBPP', 'type': 'Concept', 'label': 'Olweus Bullying Prevention Program (OBPP)'},
    {'id': 'CONCEPT_RCADS', 'type': 'Concept', 'label': 'RCADS (Revised Children Anxiety Depression Scale)'},
]

NEW_PAPERS = [
    {'id': 'PAPER_QT009_Qiu_2022', 'type': 'Paper', 'label': 'QT009 Qiu (2022) Frontiers Public Health'},
    {'id': 'PAPER_Salk_2017_Psych_Bulletin', 'type': 'Paper', 'label': 'Salk, Hyde & Abramson (2017) Psych Bulletin meta-analysis 1.7M'},
    {'id': 'PAPER_McLean_2011_J_Psychiatr_Res', 'type': 'Paper', 'label': 'McLean et al. (2011) J Psychiatr Res — gender anxiety adults'},
    {'id': 'PAPER_Olweus_1996_OBVQ', 'type': 'Paper', 'label': 'Olweus (1996) OBVQ-R bullying questionnaire'},
    {'id': 'PAPER_NolenHoeksema_2012_Annu_Rev', 'type': 'Paper', 'label': 'Nolen-Hoeksema (2012) Annu Rev Clin Psychol — emotion regulation gender'},
]

NEW_DOCS = [
    {'id': 'DOC_H4_GENDER_LOAI_RLLA', 'type': 'Doc', 'label': '01_Bao-cao/Gia_thuyet_khac_biet_gioi_tinh_theo_loai_RLLA.docx'},
    {'id': 'DOC_BNHD_BAT_NAT_HOC_DUONG', 'type': 'Doc', 'label': '01_Bao-cao/Dien_giai_yeu_to_bat_nat_hoc_duong.docx'},
    {'id': 'DOC_AUDIT_5_VONG_H4_BNHD', 'type': 'Doc', 'label': '01_Bao-cao/AUDIT_5_vong_H4_gender_va_BNHD.docx'},
]

NEW_QUESTIONS = [
    {'id': 'QA_23', 'type': 'Question',
     'label': 'Phát biểu giả thuyết về sự khác biệt giới tính theo loại RLLA: nữ > nam ở RLLATQ + RLLAXH; KHÔNG khác biệt ở RLLAC',
     'topic': 'Gender differences in psychopathology', 'date_asked': '2026-05-08'},
    {'id': 'QA_24', 'type': 'Question',
     'label': 'Diễn giải yếu tố bắt nạt học đường (BNHĐ) — tác động đến RLLA HS THCS theo chương 3 luận án',
     'topic': 'School bullying', 'date_asked': '2026-05-08'},
]

NEW_EDGES = [
    # Topic links
    ('QA_23', 'TOPIC_GENDER_PSYCHOPATHOLOGY', 'BELONGS_TO'),
    ('QA_24', 'TOPIC_SCHOOL_BULLYING', 'BELONGS_TO'),
    # QA_23 H4 gender
    ('QA_23', 'CONCEPT_GENDER_DIFFERENCES', 'EXPLAINS'),
    ('QA_23', 'CONCEPT_SEX_RATIO', 'EXPLAINS'),
    ('QA_23', 'CONCEPT_HPG_AXIS', 'EXPLAINS'),
    ('QA_23', 'CONCEPT_RUMINATION', 'EXPLAINS'),
    ('QA_23', 'CONCEPT_SOCIAL_ANXIETY_DISORDER', 'EXPLAINS'),
    ('QA_23', 'CONCEPT_GENERALIZED_ANXIETY_DISORDER', 'EXPLAINS'),
    ('QA_23', 'CONCEPT_SEPARATION_ANXIETY', 'EXPLAINS'),
    ('QA_23', 'PAPER_Salk_2017_Psych_Bulletin', 'CITES'),
    ('QA_23', 'PAPER_McLean_2011_J_Psychiatr_Res', 'CITES'),
    ('QA_23', 'PAPER_NolenHoeksema_2012_Annu_Rev', 'CITES'),
    ('QA_23', 'PAPER_QT007_Chen_2023', 'CITES'),
    ('QA_23', 'PAPER_QT008', 'CITES'),
    ('QA_23', 'PAPER_QT009_Qiu_2022', 'CITES'),
    ('QA_23', 'DOC_H4_GENDER_LOAI_RLLA', 'ANSWERED_IN'),
    # QA_24 BNHĐ
    ('QA_24', 'CONCEPT_SCHOOL_BULLYING', 'EXPLAINS'),
    ('QA_24', 'CONCEPT_OBVQ', 'EXPLAINS'),
    ('QA_24', 'CONCEPT_SCHOOL_REFUSAL', 'EXPLAINS'),
    ('QA_24', 'CONCEPT_OBPP', 'EXPLAINS'),
    ('QA_24', 'PAPER_Olweus_1996_OBVQ', 'CITES'),
    ('QA_24', 'PAPER_QT007_Chen_2023', 'CITES'),
    ('QA_24', 'PAPER_QT067', 'CITES'),
    ('QA_24', 'DOC_BNHD_BAT_NAT_HOC_DUONG', 'ANSWERED_IN'),
    # Concept-Topic
    ('CONCEPT_GENDER_DIFFERENCES', 'TOPIC_GENDER_PSYCHOPATHOLOGY', 'BELONGS_TO'),
    ('CONCEPT_SEX_RATIO', 'TOPIC_GENDER_PSYCHOPATHOLOGY', 'BELONGS_TO'),
    ('CONCEPT_HPG_AXIS', 'TOPIC_GENDER_PSYCHOPATHOLOGY', 'BELONGS_TO'),
    ('CONCEPT_RUMINATION', 'TOPIC_GENDER_PSYCHOPATHOLOGY', 'BELONGS_TO'),
    ('CONCEPT_SCHOOL_BULLYING', 'TOPIC_SCHOOL_BULLYING', 'BELONGS_TO'),
    ('CONCEPT_OBVQ', 'TOPIC_SCHOOL_BULLYING', 'BELONGS_TO'),
    ('CONCEPT_SCHOOL_REFUSAL', 'TOPIC_SCHOOL_BULLYING', 'BELONGS_TO'),
    ('CONCEPT_OBPP', 'TOPIC_SCHOOL_BULLYING', 'BELONGS_TO'),
    # Paper-Concept
    ('PAPER_QT009_Qiu_2022', 'CONCEPT_RESILIENCE', 'RELATED_TO'),
    ('PAPER_QT009_Qiu_2022', 'CONCEPT_PARENTING_STYLE', 'RELATED_TO'),
    ('PAPER_QT009_Qiu_2022', 'CONCEPT_GENDER_DIFFERENCES', 'RELATED_TO'),
    ('PAPER_Salk_2017_Psych_Bulletin', 'CONCEPT_GENDER_DIFFERENCES', 'RELATED_TO'),
    ('PAPER_McLean_2011_J_Psychiatr_Res', 'CONCEPT_GENDER_DIFFERENCES', 'RELATED_TO'),
    ('PAPER_McLean_2011_J_Psychiatr_Res', 'CONCEPT_SOCIAL_ANXIETY_DISORDER', 'RELATED_TO'),
    ('PAPER_NolenHoeksema_2012_Annu_Rev', 'CONCEPT_RUMINATION', 'RELATED_TO'),
    ('PAPER_Olweus_1996_OBVQ', 'CONCEPT_OBVQ', 'RELATED_TO'),
    ('PAPER_QT007_Chen_2023', 'CONCEPT_SCHOOL_BULLYING', 'RELATED_TO'),
    # RCADS link
    ('PAPER_VN004_NguyenThiVan_2020', 'CONCEPT_GENERALIZED_ANXIETY_DISORDER', 'RELATED_TO'),
    # Doc-Paper
    ('DOC_H4_GENDER_LOAI_RLLA', 'PAPER_Salk_2017_Psych_Bulletin', 'CITES'),
    ('DOC_H4_GENDER_LOAI_RLLA', 'PAPER_McLean_2011_J_Psychiatr_Res', 'CITES'),
    ('DOC_H4_GENDER_LOAI_RLLA', 'PAPER_QT007_Chen_2023', 'CITES'),
    ('DOC_H4_GENDER_LOAI_RLLA', 'PAPER_QT008', 'CITES'),
    ('DOC_H4_GENDER_LOAI_RLLA', 'PAPER_QT009_Qiu_2022', 'CITES'),
    ('DOC_BNHD_BAT_NAT_HOC_DUONG', 'PAPER_Olweus_1996_OBVQ', 'CITES'),
    ('DOC_BNHD_BAT_NAT_HOC_DUONG', 'PAPER_QT007_Chen_2023', 'CITES'),
    ('DOC_BNHD_BAT_NAT_HOC_DUONG', 'PAPER_QT067', 'CITES'),
    ('DOC_AUDIT_5_VONG_H4_BNHD', 'DOC_H4_GENDER_LOAI_RLLA', 'RELATED_TO'),
    ('DOC_AUDIT_5_VONG_H4_BNHD', 'DOC_BNHD_BAT_NAT_HOC_DUONG', 'RELATED_TO'),
]

NEW_RAG_QUESTIONS = [
    {
        'id': 'QA_23',
        'date_asked': '2026-05-08',
        'question': 'Phát biểu và luận chứng giả thuyết: Sự khác biệt giới tính khác nhau theo loại rối loạn lo âu (nữ > nam ở RLLATQ + RLLAXH; KHÔNG khác biệt ở RLLAC)',
        'short_answer': 'Giả thuyết H4: chênh lệch giới phụ thuộc vào LOẠI RLLA. Bằng chứng VN trực tiếp từ chương 3 luận án (n=1.352, Bảng 3.20): F (giới × RLLATQ)=44,484; F (giới × RLLAXH)=45,984; F (giới × RLLAC)=0,246 (NS). Đặc biệt, F RLLAXH > F RLLATQ — chênh lệch giới ở lo âu xã hội RỌ RỆT NHẤT, TRÁI với McLean 2011 ở người trưởng thành Mỹ (SAD KHÔNG khác biệt giới). Khả năng giải thích: tuổi vị thành niên + văn hóa Á + phương pháp đo. Phù hợp với Salk 2017 (N=1.7M, OR=1,95; đỉnh 13-15 OR=3,02), Chen 2023 (n=63.205 nữ 58% vs 48%), Wen 2020 (n=900 OR nam=0,262), Qiu 2022 (n=2.079 nữ 17,5% > nam 11,1%).',
        'topic': 'Gender differences in psychopathology',
        'concepts': ['Gender differences', 'Sex ratio', 'HPG axis', 'Rumination', 'Social Anxiety Disorder', 'Separation Anxiety'],
        'doc_path': '01_Bao-cao/Gia_thuyet_khac_biet_gioi_tinh_theo_loai_RLLA.docx',
        'paper_refs': ['QT007', 'QT008', 'QT009'],
    },
    {
        'id': 'QA_24',
        'date_asked': '2026-05-08',
        'question': 'Diễn giải yếu tố bắt nạt học đường (BNHĐ) — tác động đến rối loạn lo âu ở học sinh THCS theo chương 3 luận án',
        'short_answer': 'BNHĐ là yếu tố nguy cơ ĐỘC LẬP cho cả 3 dạng RLLA. Chương 3 (Bảng 3.27-3.28, n=1.352): β BNHĐ → RLLATQ=0,215; → RLLACL=0,376 (MẠNH NHẤT); → RLLAXH=0,253 (đều p<0,001). Phát hiện ĐẶC BIỆT: β BNHĐ → RLLACL CAO NHẤT — TRÁI với pattern ALHT/TTr (mạnh nhất RLLATQ). 3 cơ chế: (1) gắn bó-an toàn ở trường; (2) school refusal; (3) đặc thù lứa tuổi THCS. Bằng chứng quốc tế: Olweus 1996 OBVQ-R 42 mục; meta-analysis nạn nhân lo âu OR=4,72; tỷ lệ nạn nhân toàn cầu 25%; Chen 2023 (QT007 n=63.205): 4 loại bắt nạt OR 1,25-1,97. Mô hình BNHĐ-RLLA (2 factors) đạt CFI=0,999, RMSEA=0,030 — XUẤT SẮC.',
        'topic': 'School bullying',
        'concepts': ['School bullying', 'OBVQ', 'School refusal', 'OBPP', 'Separation Anxiety'],
        'doc_path': '01_Bao-cao/Dien_giai_yeu_to_bat_nat_hoc_duong.docx',
        'paper_refs': ['QT007', 'QT067'],
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
    with open(path, encoding='utf-8') as f:
        kg = json.load(f)

    existing_ids = {n['id'] for n in kg['nodes']}
    existing_edges = {(e['source'], e['target'], e['relation']) for e in kg['edges']}

    new_added = 0
    for n in NEW_TOPICS + NEW_CONCEPTS + NEW_PAPERS + NEW_DOCS + NEW_QUESTIONS:
        if n['id'] not in existing_ids:
            kg['nodes'].append(n)
            existing_ids.add(n['id'])
            new_added += 1

    edges_added = 0
    for src, tgt, rel in NEW_EDGES:
        key = (src, tgt, rel)
        if key in existing_edges or src not in existing_ids or tgt not in existing_ids:
            continue
        kg['edges'].append({'source': src, 'target': tgt, 'relation': rel})
        existing_edges.add(key)
        edges_added += 1

    from collections import Counter
    type_counts = Counter(n['type'] for n in kg['nodes'])
    rel_counts = Counter(e['relation'] for e in kg['edges'])
    kg['meta']['n_nodes'] = len(kg['nodes'])
    kg['meta']['n_edges'] = len(kg['edges'])
    kg['meta']['node_types'] = dict(type_counts)
    kg['meta']['edge_relations'] = dict(rel_counts)
    kg['meta']['updated'] = '2026-05-08'

    with open(path, 'w', encoding='utf-8') as f:
        json.dump(kg, f, ensure_ascii=False, indent=2)

    print(f'  KG: +{new_added} nodes, +{edges_added} edges; total {len(kg["nodes"])} nodes, {len(kg["edges"])} edges')


def update_rag_questions(data_dir):
    path = data_dir / 'rag_questions_index.json'
    with open(path, encoding='utf-8') as f:
        rqi = json.load(f)

    existing_ids = {e['id'] for e in rqi['entries']}
    added = 0
    for nq in NEW_RAG_QUESTIONS:
        if nq['id'] in existing_ids: continue
        entry = copy.deepcopy(nq)
        entry['tokens'] = tokenize(entry['question'] + ' ' + entry['short_answer'])
        rqi['entries'].append(entry)
        added += 1

    rqi['meta']['n_entries'] = len(rqi['entries'])
    rqi['meta']['updated'] = '2026-05-08'

    with open(path, 'w', encoding='utf-8') as f:
        json.dump(rqi, f, ensure_ascii=False, indent=2)

    print(f'  rag_questions: +{added} entries; total {len(rqi["entries"])}')


def main():
    for label, dir_ in [('LIGHT', LIGHT), ('HEAVY', HEAVY)]:
        print(f'\n==== {label} ====')
        if not dir_.exists():
            print(f'  ⚠ {dir_} missing'); continue
        update_kg(dir_)
        update_rag_questions(dir_)
    print('\nDone.')


if __name__ == '__main__':
    main()
