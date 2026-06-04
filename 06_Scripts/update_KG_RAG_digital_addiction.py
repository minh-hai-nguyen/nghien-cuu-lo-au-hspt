"""Cập nhật KG + RAG questions index với chủ đề mới Digital Addiction.

THÊM VÀO questions_kg.json:
- 1 Topic: TOPIC_DIGITAL_ADDICTION
- 7 Concept: PROBLEMATIC_SMARTPHONE_USE, SMARTPHONE_ADDICTION, SOCIAL_MEDIA_ADDICTION,
  SCREEN_TIME, FOMO, INTERNET_GAMING_DISORDER, SAS_SV
- 6 Paper: VN004, VN014 (mới), QT022, QT027, QT033, QT041 (QT007 đã có)
- 6 Doc (cho các doc mới phiên 07-08/05/2026)
- 2 Question: QA_21 (VN004 85,4% không khớp), QA_22 (giả thuyết nghiện điện thoại)
- Edges tương ứng

THÊM VÀO rag_questions_index.json:
- QA_21, QA_22

CẬP NHẬT papers_metadata.json:
- Topics cho 7 paper chứa Digital Addiction tag

MIRROR sang heavy version.
"""
import sys, io, json, copy, re
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

LIGHT = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly-light/web/data')
HEAVY = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly/web/data')

# =============== KG UPDATES ===============
NEW_TOPIC = {'id': 'TOPIC_DIGITAL_ADDICTION', 'type': 'Topic', 'label': 'Digital addiction'}

NEW_CONCEPTS = [
    {'id': 'CONCEPT_PROBLEMATIC_SMARTPHONE_USE', 'type': 'Concept', 'label': 'Problematic smartphone use (PSU)'},
    {'id': 'CONCEPT_SMARTPHONE_ADDICTION', 'type': 'Concept', 'label': 'Smartphone addiction'},
    {'id': 'CONCEPT_SOCIAL_MEDIA_ADDICTION', 'type': 'Concept', 'label': 'Social media addiction'},
    {'id': 'CONCEPT_SCREEN_TIME', 'type': 'Concept', 'label': 'Screen time'},
    {'id': 'CONCEPT_FOMO', 'type': 'Concept', 'label': 'FOMO (Fear of Missing Out)'},
    {'id': 'CONCEPT_INTERNET_GAMING_DISORDER', 'type': 'Concept', 'label': 'Internet Gaming Disorder (IGD)'},
    {'id': 'CONCEPT_SAS_SV', 'type': 'Concept', 'label': 'SAS-SV (Smartphone Addiction Scale)'},
    {'id': 'CONCEPT_BSMAS', 'type': 'Concept', 'label': 'BSMAS (Bergen Social Media Addiction Scale)'},
    {'id': 'CONCEPT_R_SQUARED', 'type': 'Concept', 'label': 'r² (coefficient of determination)'},
    {'id': 'CONCEPT_PARENT_CHILD_RELATIONSHIP', 'type': 'Concept', 'label': 'Parent-child relationship'},
    {'id': 'CONCEPT_PARENTING_STYLE', 'type': 'Concept', 'label': 'Parenting style'},
    {'id': 'CONCEPT_COPING', 'type': 'Concept', 'label': 'Coping (ứng phó)'},
    {'id': 'CONCEPT_RESILIENCE', 'type': 'Concept', 'label': 'Resilience'},
    {'id': 'CONCEPT_SELF_ESTEEM', 'type': 'Concept', 'label': 'Self-esteem (lòng tự trọng)'},
]

NEW_PAPERS = [
    {'id': 'PAPER_VN004_NguyenThiVan_2020', 'type': 'Paper', 'label': 'VN004 Nguyễn Thị Vân (2020)'},
    {'id': 'PAPER_VN014_HoangTrungHoc_2025', 'type': 'Paper', 'label': 'VN014 Hoàng Trung Học (2025)'},
    {'id': 'PAPER_QT022_Li_2025', 'type': 'Paper', 'label': 'QT022 Li (2025) BJCP'},
    {'id': 'PAPER_QT027_Fassi_2025', 'type': 'Paper', 'label': 'QT027 Fassi (2025) Nature Hum Behav'},
    {'id': 'PAPER_QT033_SchmidtPersson_2024', 'type': 'Paper', 'label': 'QT033 Schmidt-Persson (2024) JAMA Network Open'},
    {'id': 'PAPER_QT041_Zheng_2025', 'type': 'Paper', 'label': 'QT041 Zheng & Peng (2025) PRBM'},
    # External (not in canonical but referenced)
    {'id': 'PAPER_Sohn_2019_BMC_Psychiatry', 'type': 'Paper', 'label': 'Sohn et al. (2019) BMC Psychiatry meta-analysis'},
    {'id': 'PAPER_Compas_2017_Psych_Bulletin', 'type': 'Paper', 'label': 'Compas et al. (2017) Psychological Bulletin meta-analysis'},
]

NEW_DOCS = [
    {'id': 'DOC_VN004_85_4_KHONG_KHOP', 'type': 'Doc', 'label': '01_Bao-cao/VN004_85_4_khong_khop_giai_thich.docx'},
    {'id': 'DOC_GT_NGHIEN_DIEN_THOAI', 'type': 'Doc', 'label': '01_Bao-cao/Gia_thuyet_nghien_dien_thoai_yeu_to_nguy_co_RLLA.docx'},
    {'id': 'DOC_PHAN_BIEN_VN014_QUAN_HE_CHA_ME', 'type': 'Doc', 'label': '01_Bao-cao/Phan_bien_VN014_quan_he_cha_me_rong_qua.docx'},
    {'id': 'DOC_UNG_PHO_YEU_TO_BAO_VE', 'type': 'Doc', 'label': '01_Bao-cao/Ung_pho_yeu_to_bao_ve_bang_chung_va_ham_y.docx'},
    {'id': 'DOC_PHAN_BIEN_VN004_BAN_THAN_HS', 'type': 'Doc', 'label': '01_Bao-cao/Phan_bien_VN004_NguyenThiVan_yeu_to_ban_than_HS.docx'},
    {'id': 'DOC_TONG_QUAN_SELF_ESTEEM', 'type': 'Doc', 'label': '01_Bao-cao/Tong_quan_self_esteem_va_lo_au_6_NC_QT.docx'},
]

NEW_QUESTIONS = [
    {'id': 'QA_21', 'type': 'Question',
     'label': 'Trong VN004, "85,4% giải thích" công bố không khớp với r²=0,1764 từ r=0,42 — phép toán giải thích',
     'topic': 'Statistical interpretation', 'date_asked': '2026-05-08'},
    {'id': 'QA_22', 'type': 'Question',
     'label': 'Phát biểu và luận chứng giả thuyết: Nghiện điện thoại là yếu tố nguy cơ với rối loạn lo âu của HS',
     'topic': 'Digital addiction', 'date_asked': '2026-05-08'},
]

NEW_EDGES = [
    # Topic
    ('QA_22', 'TOPIC_DIGITAL_ADDICTION', 'BELONGS_TO'),
    ('QA_21', 'TOPIC_STATISTICAL_INTERPRETATION', 'BELONGS_TO'),
    # QA_21 — VN004 85,4%
    ('QA_21', 'CONCEPT_R_SQUARED', 'EXPLAINS'),
    ('QA_21', 'PAPER_VN004_NguyenThiVan_2020', 'CITES'),
    ('QA_21', 'DOC_VN004_85_4_KHONG_KHOP', 'ANSWERED_IN'),
    # QA_22 — Giả thuyết nghiện điện thoại
    ('QA_22', 'CONCEPT_PROBLEMATIC_SMARTPHONE_USE', 'EXPLAINS'),
    ('QA_22', 'CONCEPT_SMARTPHONE_ADDICTION', 'EXPLAINS'),
    ('QA_22', 'CONCEPT_SOCIAL_MEDIA_ADDICTION', 'EXPLAINS'),
    ('QA_22', 'CONCEPT_SCREEN_TIME', 'EXPLAINS'),
    ('QA_22', 'CONCEPT_FOMO', 'EXPLAINS'),
    ('QA_22', 'CONCEPT_INTERNET_GAMING_DISORDER', 'EXPLAINS'),
    ('QA_22', 'CONCEPT_SAS_SV', 'EXPLAINS'),
    ('QA_22', 'CONCEPT_BSMAS', 'EXPLAINS'),
    ('QA_22', 'PAPER_Sohn_2019_BMC_Psychiatry', 'CITES'),
    ('QA_22', 'PAPER_QT007_Chen_2023', 'CITES'),
    ('QA_22', 'PAPER_QT022_Li_2025', 'CITES'),
    ('QA_22', 'PAPER_QT027_Fassi_2025', 'CITES'),
    ('QA_22', 'PAPER_QT033_SchmidtPersson_2024', 'CITES'),
    ('QA_22', 'PAPER_QT041_Zheng_2025', 'CITES'),
    ('QA_22', 'PAPER_VN014_HoangTrungHoc_2025', 'CITES'),
    ('QA_22', 'DOC_GT_NGHIEN_DIEN_THOAI', 'ANSWERED_IN'),
    # Concept-Topic links
    ('CONCEPT_PROBLEMATIC_SMARTPHONE_USE', 'TOPIC_DIGITAL_ADDICTION', 'BELONGS_TO'),
    ('CONCEPT_SMARTPHONE_ADDICTION', 'TOPIC_DIGITAL_ADDICTION', 'BELONGS_TO'),
    ('CONCEPT_SOCIAL_MEDIA_ADDICTION', 'TOPIC_DIGITAL_ADDICTION', 'BELONGS_TO'),
    ('CONCEPT_SCREEN_TIME', 'TOPIC_DIGITAL_ADDICTION', 'BELONGS_TO'),
    ('CONCEPT_FOMO', 'TOPIC_DIGITAL_ADDICTION', 'BELONGS_TO'),
    ('CONCEPT_INTERNET_GAMING_DISORDER', 'TOPIC_DIGITAL_ADDICTION', 'BELONGS_TO'),
    # Paper-Concept links (from verified Tom-tat data)
    ('PAPER_QT007_Chen_2023', 'CONCEPT_INTERNET_GAMING_DISORDER', 'RELATED_TO'),
    ('PAPER_QT022_Li_2025', 'CONCEPT_SCREEN_TIME', 'RELATED_TO'),
    ('PAPER_QT027_Fassi_2025', 'CONCEPT_SOCIAL_MEDIA_ADDICTION', 'RELATED_TO'),
    ('PAPER_QT033_SchmidtPersson_2024', 'CONCEPT_SCREEN_TIME', 'RELATED_TO'),
    ('PAPER_QT041_Zheng_2025', 'CONCEPT_SOCIAL_MEDIA_ADDICTION', 'RELATED_TO'),
    ('PAPER_VN014_HoangTrungHoc_2025', 'CONCEPT_SCREEN_TIME', 'RELATED_TO'),
    ('PAPER_Sohn_2019_BMC_Psychiatry', 'CONCEPT_PROBLEMATIC_SMARTPHONE_USE', 'RELATED_TO'),
    # Doc-Paper links
    ('DOC_VN004_85_4_KHONG_KHOP', 'PAPER_VN004_NguyenThiVan_2020', 'CITES'),
    ('DOC_GT_NGHIEN_DIEN_THOAI', 'PAPER_QT007_Chen_2023', 'CITES'),
    ('DOC_GT_NGHIEN_DIEN_THOAI', 'PAPER_QT022_Li_2025', 'CITES'),
    ('DOC_GT_NGHIEN_DIEN_THOAI', 'PAPER_QT027_Fassi_2025', 'CITES'),
    ('DOC_GT_NGHIEN_DIEN_THOAI', 'PAPER_QT033_SchmidtPersson_2024', 'CITES'),
    ('DOC_GT_NGHIEN_DIEN_THOAI', 'PAPER_QT041_Zheng_2025', 'CITES'),
    ('DOC_GT_NGHIEN_DIEN_THOAI', 'PAPER_VN014_HoangTrungHoc_2025', 'CITES'),
    ('DOC_GT_NGHIEN_DIEN_THOAI', 'PAPER_Sohn_2019_BMC_Psychiatry', 'CITES'),
    # Other doc links from earlier session 07-08/05/2026
    ('DOC_PHAN_BIEN_VN014_QUAN_HE_CHA_ME', 'PAPER_VN014_HoangTrungHoc_2025', 'CITES'),
    ('DOC_PHAN_BIEN_VN014_QUAN_HE_CHA_ME', 'CONCEPT_PARENT_CHILD_RELATIONSHIP', 'EXPLAINS'),
    ('DOC_PHAN_BIEN_VN014_QUAN_HE_CHA_ME', 'CONCEPT_PARENTING_STYLE', 'EXPLAINS'),
    ('DOC_UNG_PHO_YEU_TO_BAO_VE', 'PAPER_Compas_2017_Psych_Bulletin', 'CITES'),
    ('DOC_UNG_PHO_YEU_TO_BAO_VE', 'CONCEPT_COPING', 'EXPLAINS'),
    ('DOC_UNG_PHO_YEU_TO_BAO_VE', 'CONCEPT_RESILIENCE', 'EXPLAINS'),
    ('DOC_PHAN_BIEN_VN004_BAN_THAN_HS', 'PAPER_VN004_NguyenThiVan_2020', 'CITES'),
    ('DOC_PHAN_BIEN_VN004_BAN_THAN_HS', 'CONCEPT_SELF_ESTEEM', 'EXPLAINS'),
    ('DOC_TONG_QUAN_SELF_ESTEEM', 'CONCEPT_SELF_ESTEEM', 'EXPLAINS'),
]

# =============== rag_questions_index ENTRIES ===============
NEW_RAG_QUESTIONS = [
    {
        'id': 'QA_21',
        'date_asked': '2026-05-08',
        'question': 'Trong VN004 Nguyễn Thị Vân, "85,4% giải thích" công bố không khớp với r² = 0,1764 từ r = 0,42 — vì sao?',
        'short_answer': 'r = 0,42 → r² = 0,42² = 0,1764 = 17,64% là tỷ lệ phương sai Y giải thích bởi X qua tương quan tuyến tính bivariate. Tác giả VN004 công bố "85,4%" — chênh lệch 67,76 điểm phần trăm (4,84 lần). Để có 85,4% cần r = √0,854 ≈ 0,924 (gần gấp đôi 0,42 đã công bố). Bốn khả năng giải thích: (1) R² hồi quy đa biến — toán hợp lệ nhưng tác giả thiếu báo cáo β, F-test, VIF; (2) tỷ lệ % HS có biểu hiện — nhầm thuật ngữ; (3) tỷ lệ định tính từ phỏng vấn — không thống kê; (4) sai sót đánh máy. Khuyến nghị KHÔNG trích con số 85,4% vào báo cáo CTH.',
        'topic': 'Statistical interpretation',
        'concepts': ['r²', 'Coefficient of determination', 'R² multiple regression', 'STROBE'],
        'doc_path': '01_Bao-cao/VN004_85_4_khong_khop_giai_thich.docx',
        'paper_refs': ['VN004'],
    },
    {
        'id': 'QA_22',
        'date_asked': '2026-05-08',
        'question': 'Phát biểu và luận chứng giả thuyết: Nghiện điện thoại là yếu tố nguy cơ đối với rối loạn lo âu của học sinh',
        'short_answer': 'Giả thuyết: học sinh THCS có mức độ nghiện điện thoại (đo bằng SAS-SV hoặc BSMAS) cao hơn sẽ có tỷ lệ RLLA cao hơn — sau khi kiểm soát biến nhân khẩu và yếu tố tâm lý xã hội đồng biến. Bằng chứng: Sohn 2019 BMC Psychiatry meta-analysis 41 NC N=41.871 — OR lo âu = 3,05 (KTC 95% 2,64–3,53; I²=0%); Chen 2023 BMC Psychiatry n=63.205 — game disorder OR = 5,00; Zheng & Peng 2025 PRBM — SMAS r=0,415, β=0,153, mediation 63,13% qua self-efficacy; Fassi 2025 Nature Hum Behav n=3.340 với chẩn đoán lâm sàng; Schmidt-Persson 2024 JAMA Network Open RCT thí điểm n=89 gia đình Cohen d=0,53; Hoàng Trung Học 2025 VN n=8.473 — β=0,176. Cảnh báo: Li 2025 BJCP dọc 12 tháng — screen time T1 KHÔNG dự báo lo âu T2 (p=0,443) → quan hệ HAI CHIỀU. Phân biệt TIME ≠ ADDICTION. VN014 đo time, đề tài có thể lấp khoảng trống bằng đo addiction.',
        'topic': 'Digital addiction',
        'concepts': ['Problematic smartphone use (PSU)', 'Smartphone addiction', 'Social media addiction', 'Screen time', 'FOMO', 'Internet Gaming Disorder', 'SAS-SV', 'BSMAS'],
        'doc_path': '01_Bao-cao/Gia_thuyet_nghien_dien_thoai_yeu_to_nguy_co_RLLA.docx',
        'paper_refs': ['VN014', 'QT007', 'QT022', 'QT027', 'QT033', 'QT041'],
    },
]

# =============== papers_metadata UPDATES ===============
TOPIC_UPDATES = {
    'VN014': ['Mental Health (general)', 'Digital addiction', 'Parent-child relationship'],
    'QT007': ['Mental Health (general)', 'Internet Gaming Disorder', 'Sleep'],
    'QT022': ['Digital addiction', 'Screen time', 'Longitudinal design'],
    'QT027': ['Digital addiction', 'Social media', 'Mental Health (general)'],
    'QT033': ['Digital addiction', 'Screen time', 'RCT'],
    'QT041': ['Digital addiction', 'Social media addiction', 'Mediation analysis'],
}


def tokenize(text):
    """Simple tokenizer for the rag_questions_index style."""
    text = text.lower()
    tokens = re.findall(r'[\w]+', text, re.UNICODE)
    counts = {}
    for t in tokens:
        if len(t) >= 2:
            counts[t] = counts.get(t, 0) + 1
    return counts


def update_kg(data_dir):
    """Update questions_kg.json — append new nodes/edges, update meta."""
    path = data_dir / 'questions_kg.json'
    with open(path, encoding='utf-8') as f:
        kg = json.load(f)

    existing_ids = {n['id'] for n in kg['nodes']}
    existing_edges = {(e['source'], e['target'], e['relation']) for e in kg['edges']}

    # Append new nodes (skip duplicates)
    new_added = 0
    for n in [NEW_TOPIC] + NEW_CONCEPTS + NEW_PAPERS + NEW_DOCS + NEW_QUESTIONS:
        if n['id'] not in existing_ids:
            kg['nodes'].append(n)
            existing_ids.add(n['id'])
            new_added += 1

    # Append new edges (skip duplicates, skip if source/target not in nodes)
    edges_added = 0
    edges_skipped = 0
    for src, tgt, rel in NEW_EDGES:
        key = (src, tgt, rel)
        if key in existing_edges:
            continue
        if src not in existing_ids:
            print(f'  ⚠ skip edge — source {src} not found')
            edges_skipped += 1
            continue
        if tgt not in existing_ids:
            print(f'  ⚠ skip edge — target {tgt} not found')
            edges_skipped += 1
            continue
        kg['edges'].append({'source': src, 'target': tgt, 'relation': rel})
        existing_edges.add(key)
        edges_added += 1

    # Update meta
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

    print(f'  KG: +{new_added} nodes, +{edges_added} edges (skipped {edges_skipped})')
    print(f'  KG total: {len(kg["nodes"])} nodes, {len(kg["edges"])} edges')


def update_rag_questions(data_dir):
    """Append QA_21 + QA_22 to rag_questions_index.json with tokens."""
    path = data_dir / 'rag_questions_index.json'
    with open(path, encoding='utf-8') as f:
        rqi = json.load(f)

    existing_ids = {e['id'] for e in rqi['entries']}
    added = 0
    for nq in NEW_RAG_QUESTIONS:
        if nq['id'] in existing_ids:
            continue
        entry = copy.deepcopy(nq)
        # Build tokens for TF-IDF
        text = entry['question'] + ' ' + entry['short_answer']
        entry['tokens'] = tokenize(text)
        rqi['entries'].append(entry)
        added += 1

    rqi['meta']['n_entries'] = len(rqi['entries'])
    rqi['meta']['updated'] = '2026-05-08'

    with open(path, 'w', encoding='utf-8') as f:
        json.dump(rqi, f, ensure_ascii=False, indent=2)

    print(f'  rag_questions_index: +{added} entries (total {len(rqi["entries"])})')


def update_papers_metadata(data_dir):
    """Update topics for 6 papers."""
    path = data_dir / 'papers_metadata.json'
    with open(path, encoding='utf-8') as f:
        pm = json.load(f)

    entries = pm.get('entries', pm if isinstance(pm, list) else [])
    updated = 0
    for e in entries:
        if isinstance(e, dict) and e.get('id') in TOPIC_UPDATES:
            e['topics'] = TOPIC_UPDATES[e['id']]
            updated += 1

    if 'meta' in pm:
        pm['meta']['updated'] = '2026-05-08'

    with open(path, 'w', encoding='utf-8') as f:
        json.dump(pm, f, ensure_ascii=False, indent=2)

    print(f'  papers_metadata: updated topics for {updated} papers')


def main():
    for label, dir_ in [('LIGHT', LIGHT), ('HEAVY', HEAVY)]:
        print(f'\n==== {label} ({dir_}) ====')
        if not dir_.exists():
            print(f'  ⚠ {dir_} does not exist — skip')
            continue
        update_kg(dir_)
        update_rag_questions(dir_)
        update_papers_metadata(dir_)

    print('\nDone.')


if __name__ == '__main__':
    main()
