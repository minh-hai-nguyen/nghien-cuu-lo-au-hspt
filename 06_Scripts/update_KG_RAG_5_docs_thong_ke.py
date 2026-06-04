"""Update KG + RAG cho 5 doc tra loi cau hoi thong ke gan day:
- F_test_ANOVA_gioi (QA_32)
- Chenh_lech_gioi_LAXH_vs_LATQ (QA_33)
- R_squared_YTBV (QA_34)
- TTr_ngang_ALHT_85_89 (QA_35)
- Giai_phap_quy_ra_phan_tram (QA_36)
"""
import sys, io, json, copy, re
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

LIGHT = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly-light/web/data')
HEAVY = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/tro-ly-nghien-cuu-tam-ly/web/data')

NEW_CONCEPTS = [
    {'id': 'CONCEPT_F_TEST', 'type': 'Concept', 'label': 'F-test (Fisher\'s test) trong ANOVA'},
    {'id': 'CONCEPT_ANOVA', 'type': 'Concept', 'label': 'ANOVA (Analysis of Variance)'},
    {'id': 'CONCEPT_COHEN_D_RATIO', 'type': 'Concept', 'label': 'Tỷ số Cohen d giữa các nhóm/biến'},
    {'id': 'CONCEPT_BETA_RATIO', 'type': 'Concept', 'label': 'Tỷ số |β| để so sánh cường độ tác động'},
    {'id': 'CONCEPT_RAW_VS_STANDARDIZED', 'type': 'Concept', 'label': 'Chênh lệch raw vs effect size chuẩn hóa'},
    {'id': 'CONCEPT_VARIANCE_EXPLAINED', 'type': 'Concept', 'label': 'Variance explained (phương sai được giải thích)'},
    {'id': 'CONCEPT_GUARDIAN_BARRIER_RATIO', 'type': 'Concept', 'label': 'Tỷ số yếu tố bảo vệ / nguy cơ — đánh giá sức mạnh tương đối'},
    {'id': 'CONCEPT_T_SCORE_CUTOFF', 'type': 'Concept', 'label': 'T-score cutoff (60/65/70/80) cho phân loại lâm sàng'},
    {'id': 'CONCEPT_GIA_DUC_HOI_DONG_VN', 'type': 'Concept', 'label': 'Yêu cầu của Hội đồng VN về tỷ lệ % trong báo cáo luận án'},
]

NEW_DOCS = [
    {'id': 'DOC_F_TEST_ANOVA_GIOI', 'type': 'Doc',
     'label': '01_Bao-cao/F_test_ANOVA_gioi_la_gi.docx'},
    {'id': 'DOC_CHENH_LECH_GIOI_LAXH_LATQ', 'type': 'Doc',
     'label': '01_Bao-cao/Chenh_lech_gioi_LAXH_vs_LATQ_2_goc_nhin.docx'},
    {'id': 'DOC_R_SQUARED_YTBV', 'type': 'Doc',
     'label': '01_Bao-cao/R_squared_YTBV_la_gi.docx'},
    {'id': 'DOC_TTR_NGANG_ALHT_85_89', 'type': 'Doc',
     'label': '01_Bao-cao/Tu_trong_ngang_ALHT_85_89_phan_tram_giai_thich.docx'},
    {'id': 'DOC_GIAI_PHAP_QUY_PHAN_TRAM', 'type': 'Doc',
     'label': '01_Bao-cao/Giai_phap_quy_ra_phan_tram_de_HoiDong_VN_OK.docx'},
]

NEW_QUESTIONS = [
    {'id': 'QA_32', 'type': 'Question',
     'label': 'F-test ANOVA giới là gì? Áp dụng vào Bảng 3.20',
     'topic': 'Statistical interpretation', 'date_asked': '2026-05-09'},
    {'id': 'QA_33', 'type': 'Question',
     'label': 'Chênh lệch giới ở RLLAXH so với RLLATQ — hai góc nhìn (raw vs Cohen d)',
     'topic': 'Statistical interpretation', 'date_asked': '2026-05-09'},
    {'id': 'QA_34', 'type': 'Question',
     'label': 'R² YTBV gọi là gì? Hệ số xác định = 12,4% phương sai RLLA giải thích bởi YTBV',
     'topic': 'Statistical interpretation', 'date_asked': '2026-05-09'},
    {'id': 'QA_35', 'type': 'Question',
     'label': 'Lòng tự trọng có cường độ tác động ngang ALHT 85-89% nghĩa là gì?',
     'topic': 'Statistical interpretation', 'date_asked': '2026-05-09'},
    {'id': 'QA_36', 'type': 'Question',
     'label': 'Giải pháp quy ra tỷ lệ % để Hội đồng VN OK — 7 phương án',
     'topic': 'Statistical interpretation', 'date_asked': '2026-05-09'},
]

NEW_EDGES = [
    # QA_32
    ('QA_32', 'TOPIC_STATISTICAL_INTERPRETATION', 'BELONGS_TO'),
    ('QA_32', 'CONCEPT_F_TEST', 'EXPLAINS'),
    ('QA_32', 'CONCEPT_ANOVA', 'EXPLAINS'),
    ('QA_32', 'CONCEPT_COHEN_S_D', 'EXPLAINS'),
    ('QA_32', 'DOC_F_TEST_ANOVA_GIOI', 'ANSWERED_IN'),
    # QA_33
    ('QA_33', 'TOPIC_STATISTICAL_INTERPRETATION', 'BELONGS_TO'),
    ('QA_33', 'TOPIC_GENDER_PSYCHOPATHOLOGY', 'BELONGS_TO'),
    ('QA_33', 'CONCEPT_RAW_VS_STANDARDIZED', 'EXPLAINS'),
    ('QA_33', 'CONCEPT_COHEN_S_D', 'EXPLAINS'),
    ('QA_33', 'CONCEPT_F_TEST', 'EXPLAINS'),
    ('QA_33', 'DOC_CHENH_LECH_GIOI_LAXH_LATQ', 'ANSWERED_IN'),
    # QA_34
    ('QA_34', 'TOPIC_STATISTICAL_INTERPRETATION', 'BELONGS_TO'),
    ('QA_34', 'CONCEPT_R_SQUARED', 'EXPLAINS'),
    ('QA_34', 'CONCEPT_VARIANCE_EXPLAINED', 'EXPLAINS'),
    ('QA_34', 'CONCEPT_R_SQUARED_INTERPRETATION', 'EXPLAINS'),
    ('QA_34', 'CONCEPT_YTBV_TONG', 'EXPLAINS'),
    ('QA_34', 'DOC_R_SQUARED_YTBV', 'ANSWERED_IN'),
    # QA_35
    ('QA_35', 'TOPIC_STATISTICAL_INTERPRETATION', 'BELONGS_TO'),
    ('QA_35', 'CONCEPT_BETA_RATIO', 'EXPLAINS'),
    ('QA_35', 'CONCEPT_GUARDIAN_BARRIER_RATIO', 'EXPLAINS'),
    ('QA_35', 'CONCEPT_SELF_ESTEEM', 'EXPLAINS'),
    ('QA_35', 'DOC_TTR_NGANG_ALHT_85_89', 'ANSWERED_IN'),
    # QA_36
    ('QA_36', 'TOPIC_STATISTICAL_INTERPRETATION', 'BELONGS_TO'),
    ('QA_36', 'CONCEPT_PERCENT_CONVERSION', 'EXPLAINS'),
    ('QA_36', 'CONCEPT_RCADS_CUTOFF', 'EXPLAINS'),
    ('QA_36', 'CONCEPT_T_SCORE_CUTOFF', 'EXPLAINS'),
    ('QA_36', 'CONCEPT_GIA_DUC_HOI_DONG_VN', 'EXPLAINS'),
    ('QA_36', 'CONCEPT_Z_SCORE', 'EXPLAINS'),
    ('QA_36', 'PAPER_VN002', 'CITES'),
    ('QA_36', 'PAPER_VN014_HoangTrungHoc_2025', 'CITES'),
    ('QA_36', 'PAPER_VN020_TranHoVinhLoc_2024', 'CITES'),
    ('QA_36', 'DOC_GIAI_PHAP_QUY_PHAN_TRAM', 'ANSWERED_IN'),
    # Concept-Concept
    ('CONCEPT_F_TEST', 'CONCEPT_ANOVA', 'RELATED_TO'),
    ('CONCEPT_BETA_RATIO', 'CONCEPT_GUARDIAN_BARRIER_RATIO', 'RELATED_TO'),
    ('CONCEPT_RAW_VS_STANDARDIZED', 'CONCEPT_COHEN_S_D', 'RELATED_TO'),
]

NEW_RAG_QUESTIONS = [
    {
        'id': 'QA_32',
        'date_asked': '2026-05-09',
        'question': 'F-test ANOVA giới là gì?',
        'short_answer': 'F-test ANOVA giới = kiểm định thống kê trả lời "Mức độ lo âu CÓ khác biệt có ý nghĩa thống kê giữa NAM và NỮ không?". Công thức F = (Phương sai GIỮA nhóm) / (Phương sai TRONG nhóm). F lớn + p<0,05 = có khác biệt; F nhỏ + p>0,05 = không. Bảng 3.20: F (giới × RLLATQ)=44,484 (p<0,001) CÓ khác biệt; F (giới × RLLAC)=0,246 (p=0,620) KHÔNG khác biệt; F (giới × RLLAXH)=45,984 (p<0,001) CÓ khác biệt. CẢNH BÁO: F-test KHÁC effect size — F LỚN không nghĩa cường độ LỚN. Cohen d RLLATQ=0,365 và RLLAXH=0,370 GẦN BẰNG NHAU dù F khác biệt. Báo cáo CẢ HAI: F-test (có khác biệt không) + Cohen d (cường độ).',
        'topic': 'Statistical interpretation',
        'concepts': ['F-test', 'ANOVA', 'Cohen d', 'Effect size'],
        'doc_path': '01_Bao-cao/F_test_ANOVA_gioi_la_gi.docx',
        'paper_refs': [],
    },
    {
        'id': 'QA_33',
        'date_asked': '2026-05-09',
        'question': 'Ở LAXH, nữ có mức độ khác biệt cao hơn nam so với LATQ?',
        'short_answer': 'CÓ ĐÚNG theo MỘT GÓC NHÌN. Theo CHÊNH LỆCH ĐTB RAW: LAXH chênh nữ-nam = 9,54 điểm > LATQ = 8,04 điểm (+18,7%, +1,5 điểm). Theo Cohen d (chuẩn hóa SD): LAXH = 0,370 vs LATQ = 0,365 — gần BẰNG NHAU (chênh chỉ 1,5%). Lý do khác biệt: SD pooled LAXH (25,77) > LATQ (22,04). Khi chuẩn hóa, chênh raw bị "co lại". Diễn giải: LAXH có biến thiên cá nhân lớn hơn — một số HS rất sợ tình huống xã hội, một số bình thản. Báo cáo CẢ HAI để Hội đồng VN không hiểu sai. Phát biểu đúng: "Trên thang điểm thực tế chênh ở LAXH lớn hơn LATQ 18,7%, nhưng Cohen d gần bằng nhau".',
        'topic': 'Statistical interpretation',
        'concepts': ['Raw vs standardized', 'Cohen d', 'F-test', 'Gender differences'],
        'doc_path': '01_Bao-cao/Chenh_lech_gioi_LAXH_vs_LATQ_2_goc_nhin.docx',
        'paper_refs': [],
    },
    {
        'id': 'QA_34',
        'date_asked': '2026-05-09',
        'question': 'R² YTBV gọi là gì?',
        'short_answer': 'R² gọi là HỆ SỐ XÁC ĐỊNH (Coefficient of Determination). R² YTBV = TỶ LỆ PHƯƠNG SAI rối loạn lo âu được giải thích bởi yếu tố bảo vệ tổng. Chương 3 luận án: R² YTBV = 0,124 → "YTBV tổng giải thích 12,4% phương sai RLLA; còn 87,6% do yếu tố khác". Cohen 1988: R²<0,02 nhỏ; 0,02-0,13 small; 0,13-0,26 medium; ≥0,26 large. R²=0,124 nằm RANH GIỚI giữa small và medium. So sánh: R² YTNC = 0,558 (rất lớn); R² tích hợp YTNC+YTBV = 0,598. Tỷ số YTBV/YTNC = 0,222 — YTBV chỉ bằng 22% YTNC, gợi ý ưu tiên giảm yếu tố nguy cơ trong can thiệp.',
        'topic': 'Statistical interpretation',
        'concepts': ['R²', 'Variance explained', 'Cohen 1988 thresholds', 'YTBV', 'YTNC'],
        'doc_path': '01_Bao-cao/R_squared_YTBV_la_gi.docx',
        'paper_refs': [],
    },
    {
        'id': 'QA_35',
        'date_asked': '2026-05-09',
        'question': 'Lòng tự trọng có cường độ tác động tương đương 85-89% với áp lực học tập nghĩa là gì?',
        'short_answer': 'Tỷ số |β TTr| / |β ALHT|: RLLATQ = 0,455/0,510 = 0,892 (~89%); RLLAXH = 0,415/0,490 = 0,847 (~85%). Nghĩa là: khi ALHT tăng 1 SD làm lo âu tăng 1,0 SD, thì TTr tăng 1 SD làm lo âu GIẢM 0,85-0,89 SD. Tự trọng có "lực bảo vệ" ngang 85-89% "lực gây lo âu" của ALHT. Dấu β khác (DƯƠNG vs ÂM) chỉ chỉ HƯỚNG, |β| chỉ CƯỜNG ĐỘ. Hàm ý can thiệp: có thể GIẢM lo âu BẰNG (a) giảm ALHT hoặc (b) tăng TTr — hiệu quả TƯƠNG ĐƯƠNG. Tăng TTr KHẢ THI HƠN qua tập huấn. Lưu ý: tỷ số 85-89% CHỈ áp dụng cho RLLATQ + RLLAXH; với RLLACL chỉ ~34% (β TTr=−0,087 vs β ALHT=0,253).',
        'topic': 'Statistical interpretation',
        'concepts': ['Beta ratio', 'Standardized beta', 'Self-esteem', 'Academic pressure', 'Guardian-barrier ratio'],
        'doc_path': '01_Bao-cao/Tu_trong_ngang_ALHT_85_89_phan_tram_giai_thich.docx',
        'paper_refs': [],
    },
    {
        'id': 'QA_36',
        'date_asked': '2026-05-09',
        'question': 'Giải pháp quy ra tỷ lệ % để Hội đồng VN OK với báo cáo luận án',
        'short_answer': 'Bảy phương án theo thứ tự ưu tiên: (1) Liên hệ tác giả luận án xin DATA RAW + đếm trực tiếp HS vượt ngưỡng — CHÍNH XÁC NHẤT; (2) Áp dụng cutoff Chorpita 2000 (T-score 60/65/70/80) — chuẩn quốc tế; (3) Phân loại 5 mức độ (bình thường/borderline/mild/moderate/severe) như VN006; (4) Z-score giả định normal — tính ngay; (5) So sánh với meta-analysis quốc tế cùng RCADS; (6) Trình bày SONG HÀNH ĐTB liên tục + % phân loại — toàn diện; (7) Bổ sung mục 3.2.X "Tỷ lệ HS theo mức độ". Ước lượng z-score: RLLATQ 42,5%/33,9%/26,0%/13,6% (T≥60/65/70/80); RLLAXH 32,6%/26,0%/20,1%/11,0%; RLLAC 7,5%/5,0%/3,2%/1,2%. Tỷ lệ NẰM TRONG dải VN (V-NAMHS 18,45%; HT Học 25,4%; VN029 50,3%) — phù hợp y văn.',
        'topic': 'Statistical interpretation',
        'concepts': ['Percent conversion', 'RCADS cutoff', 'T-score cutoff', 'Z-score', 'Hội đồng VN'],
        'doc_path': '01_Bao-cao/Giai_phap_quy_ra_phan_tram_de_HoiDong_VN_OK.docx',
        'paper_refs': ['VN002', 'VN014', 'VN020', 'VN006', 'VN029'],
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
