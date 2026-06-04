"""Append 5 new Q&A entries (27/04 - 02/05) to rag_questions_index + questions_kg.

QA_16: I² heterogeneity (27/04, thầy tự viết doc)
QA_17: 4 OR parenting + resilience (02/05)
QA_18: OR vs RR — Wen 2020 'nguy cơ' diễn đạt (02/05)
QA_19: Help-seeking VN treatment gap 91,6% (02/05)
QA_20: VTN definition (02/05)
"""
import sys, io, json, re, math
from pathlib import Path
from datetime import datetime
from collections import Counter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = Path('c:/Users/HLC/OneDrive/read_books/Lo-au')
QA_LIGHT = ROOT / 'tro-ly-nghien-cuu-tam-ly-light/web/data/rag_questions_index.json'
KG_LIGHT = ROOT / 'tro-ly-nghien-cuu-tam-ly-light/web/data/questions_kg.json'
QA_HEAVY = ROOT / 'tro-ly-nghien-cuu-tam-ly/web/data/rag_questions_index.json'
KG_HEAVY = ROOT / 'tro-ly-nghien-cuu-tam-ly/web/data/questions_kg.json'

STOPWORDS = set('''a an and or but in on at to from for of with by is are was were be been being
have has had do does did the this that these those it its they them their we us our you your
he she his her i me my là và hoặc nhưng trong tại đến từ cho của với bởi là các một những này đó
nó họ chúng ta chúng tôi tôi bạn anh ấy cô ấy có được bị đã đang sẽ cũng không chỉ rất đều như
ai ở khi nào tới sau trước giữa ngoài trên dưới bên cùng qua vào ra lên xuống đây đó kia'''.split())

def _tokenize(text):
    if not text: return []
    text = text.lower()
    tokens = re.findall(r'\b[\wÀ-ỹđĐ]+\b', text, flags=re.UNICODE)
    return [t for t in tokens if t not in STOPWORDS and len(t) > 1]

def tf_dict(text):
    tokens = _tokenize(text)
    return dict(Counter(tokens))

# ---------------------------------------------------------
# 5 NEW Q&A entries
# ---------------------------------------------------------
NEW_QA = [
    {
        'id': 'QA_16',
        'date_asked': '2026-04-27',
        'question': 'I² (I-squared / I bình phương) là gì? Khác Cochran Q và τ² thế nào? Ngưỡng diễn giải Cochrane?',
        'short_answer': 'I² là thước đo dị tính (heterogeneity) trong meta-analysis, biểu thị % phương sai giữa các nghiên cứu DO khác biệt thực giữa nghiên cứu (không phải sai số ngẫu nhiên). Công thức: I² = max(0, (Q − df) / Q × 100%), với Q = Cochran Q-statistic, df = số NC − 1. Ngưỡng Cochrane: 0–25% thấp; 25–50% trung bình; 50–75% đáng kể; 75–100% rất cao. Khác τ² (tau-squared, đo lượng PHƯƠNG SAI tuyệt đối giữa NC trong random-effects model). KHÔNG nhầm với chi-square test (test 2 biến phân loại). Khi I² cao → dùng random-effects + sub-group/meta-regression.',
        'topic': 'Heterogeneity',
        'concepts': ['I²', 'Cochran Q', 'τ² (tau-squared)', 'Heterogeneity', 'Random-effects', 'Fixed-effects', 'Cochrane', 'PRISMA 2020'],
        'doc_path': '01_Bao-cao/I_binh_phuong_giai_thich_cho_thay_27042026.docx',
        'paper_refs': ['QT049', 'QT040'],
    },
    {
        'id': 'QA_17',
        'date_asked': '2026-05-02',
        'question': 'Bình luận mức độ 4 OR (parenting tiêu cực/tích cực + resilience thấp ↔ lo âu/trầm cảm). Cao hay thấp?',
        'short_answer': 'OR=2,01 (parenting tiêu cực ↔ lo âu) → YẾU-TRUNG BÌNH. OR=0,32 (parenting tích cực ↔ lo âu, protective; 1/0,32≈3,13) → TRUNG BÌNH. OR=6,74 (resilience thấp ↔ trầm cảm) → RẤT MẠNH (very strong, > 6). OR=2,80 (resilience thấp ↔ lo âu) → TRUNG BÌNH. Phát hiện nổi bật: resilience tác động lên trầm cảm gấp >2 lần lên lo âu; protective effect của nuôi dạy tích cực mạnh hơn risk effect của nuôi dạy tiêu cực — nên ưu tiên xây dựng nuôi dạy tích cực. Khung: <1,5 yếu / 1,5–2,5 yếu-TB / 2,5–4 TB / 4–6 mạnh / >6 rất mạnh.',
        'topic': 'Statistical interpretation',
        'concepts': ['Odds Ratio', 'OR magnitude', 'Parenting style', 'Resilience', 'Protective factor', 'Risk factor', 'Effect size'],
        'doc_path': '01_Bao-cao/OR_muc_do_binh_luan_parenting_resilience.docx',
        'paper_refs': ['QT008', 'QT067'],
    },
    {
        'id': 'QA_18',
        'date_asked': '2026-05-02',
        'question': 'Diễn đạt OR=11,579 (Wen 2020) là "nguy cơ cao gấp 11,5 lần" có đúng không?',
        'short_answer': 'ĐÚNG về ý nhưng SAI về thuật ngữ thống kê. OR (Odds Ratio) ≠ RR (Risk Ratio). Khi outcome KHÔNG hiếm (lo âu nặng HS ~5–15% theo VNAMHS), OR phóng đại RR — với OR=11,5 prevalence~10%, RR thật ước ~4–6. Cách diễn đạt chuẩn: "có ODDS (khả năng tương đối) gấp khoảng 11,5 lần" thay vì "nguy cơ". OR=11,579 thuộc nhóm CỰC MẠNH (>6) — phát hiện đáng nghi cần kiểm tra KTC 95% rộng và cỡ mẫu. Tham khảo: Zhang & Yu 1998, Davies et al. 1998 BMJ.',
        'topic': 'Statistical interpretation',
        'concepts': ['Odds Ratio', 'Risk Ratio', 'OR vs RR', 'Outcome prevalence', 'Statistical magnitude', 'Wen 2020'],
        'doc_path': '01_Bao-cao/OR_Wen_2020_dien_dat_nguy_co_vs_odds.docx',
        'paper_refs': ['QT008'],
    },
    {
        'id': 'QA_19',
        'date_asked': '2026-05-02',
        'question': 'Tỷ lệ HS có vấn đề SKTT chủ động tìm gặp bác sĩ tâm lý ở trường là bao nhiêu? Có tài liệu nào ước tính?',
        'short_answer': 'CÓ — VN002 VNAMHS 2022 (đại diện QG VN, n=5.996). Số liệu cốt lõi: 21,7% VTN VN có vấn đề SKTT 12 tháng qua; chỉ 8,4% trong nhóm này dùng dịch vụ → treatment gap 91,6%. Trong nhóm dùng dịch vụ: 5,5% đến nhân viên trường, 1,4% đến chuyên gia tâm thần. Tính trên TỔNG HS có vấn đề SKTT: chỉ ~0,46% chủ động tìm gặp nhân viên trường (8,4% × 5,5%); ~0,12% gặp chuyên gia tâm thần. Bổ sung: VN022 UNICEF 2022 (5 tỉnh, stigma + rào cản); QT066 Murphy 2024 peer support; tham chiếu QT lớn: Gulliver 2010, Radez 2021.',
        'topic': 'Help-seeking',
        'concepts': ['Help-seeking', 'Treatment gap', 'School counselor', 'School staff', 'Mental health service', 'V-NAMHS 2022', 'Service utilization'],
        'doc_path': '01_Bao-cao/Ty_le_HS_tim_gap_bac_si_tam_ly_o_truong.docx',
        'paper_refs': ['VN002', 'VN022', 'QT066', 'QT008'],
    },
    {
        'id': 'QA_20',
        'date_asked': '2026-05-02',
        'question': 'VTN là gì?',
        'short_answer': 'VTN = Vị Thành Niên (adolescent / teenager). Theo WHO: 10–19 tuổi. VNAMHS 2022 khảo sát nhóm hẹp hơn 10–17. Phân biệt: TTN = thanh thiếu niên (đến ~24); TN = thanh niên (16–30 theo Luật Thanh niên VN); HS = học sinh; THCS lớp 6–9 (~11–15); THPT lớp 10–12 (~15–18); SV = sinh viên (~18–22); SKTT = sức khỏe tâm thần. Đối tượng đề tài lo âu HS THCS/THPT (~11–18) hoàn toàn nằm trong nhóm VTN.',
        'topic': 'Statistical terminology',
        'concepts': ['VTN', 'Adolescent', 'WHO age groups', 'Vietnamese terminology'],
        'doc_path': 'chat only',
        'paper_refs': ['VN002', 'VN022'],
    },
]

def main():
    # ---------- Load existing ----------
    with open(QA_LIGHT, 'r', encoding='utf-8') as f:
        qa_data = json.load(f)
    with open(KG_LIGHT, 'r', encoding='utf-8') as f:
        kg_data = json.load(f)

    existing_ids = {e['id'] for e in qa_data['entries']}

    # ---------- Append 5 new Q&A ----------
    for nq in NEW_QA:
        if nq['id'] in existing_ids:
            print(f'SKIP {nq["id"]} (already exists)')
            continue
        # Build TF tokens for full Q+A+concepts text
        full_text = nq['question'] + ' ' + nq['short_answer'] + ' ' + ' '.join(nq.get('concepts', []))
        tokens = tf_dict(full_text)
        nq['tokens'] = tokens
        qa_data['entries'].append(nq)

    # ---------- Recompute IDF ----------
    n_docs = len(qa_data['entries'])
    df = Counter()
    for e in qa_data['entries']:
        for tok in e.get('tokens', {}):
            df[tok] += 1
    idf = {}
    for tok, d in df.items():
        idf[tok] = round(math.log((n_docs + 1) / (d + 1)) + 1, 4)
    qa_data['idf'] = idf

    qa_data['meta']['n_entries'] = n_docs
    qa_data['meta']['total_unique_tokens'] = len(idf)
    qa_data['meta']['updated'] = datetime.now().strftime('%Y-%m-%d')

    # ---------- KG: add nodes + edges ----------
    existing_node_ids = {n['id'] for n in kg_data['nodes']}

    def concept_id(label):
        return 'CONCEPT_' + re.sub(r'[^A-Z0-9]+', '_', label.upper()).strip('_')

    new_topics = set()
    for nq in NEW_QA:
        # Question node
        qnode_id = nq['id']
        if qnode_id not in existing_node_ids:
            kg_data['nodes'].append({
                'id': qnode_id,
                'type': 'Question',
                'label': nq['question'][:80],
                'topic': nq['topic'],
                'date_asked': nq['date_asked'],
            })
            existing_node_ids.add(qnode_id)

        # Topic node
        topic_id = 'TOPIC_' + re.sub(r'[^A-Z0-9]+', '_', nq['topic'].upper()).strip('_')
        if topic_id not in existing_node_ids:
            kg_data['nodes'].append({
                'id': topic_id,
                'type': 'Topic',
                'label': nq['topic'],
            })
            existing_node_ids.add(topic_id)
            new_topics.add(nq['topic'])
        kg_data['edges'].append({'source': qnode_id, 'target': topic_id, 'relation': 'BELONGS_TO'})

        # Concept nodes + edges
        for c in nq.get('concepts', []):
            cid = concept_id(c)
            if cid not in existing_node_ids:
                kg_data['nodes'].append({'id': cid, 'type': 'Concept', 'label': c})
                existing_node_ids.add(cid)
            kg_data['edges'].append({'source': qnode_id, 'target': cid, 'relation': 'EXPLAINS'})

        # Paper nodes + edges
        for p in nq.get('paper_refs', []):
            pid = 'PAPER_' + p
            if pid not in existing_node_ids:
                kg_data['nodes'].append({'id': pid, 'type': 'Paper', 'label': p})
                existing_node_ids.add(pid)
            kg_data['edges'].append({'source': qnode_id, 'target': pid, 'relation': 'CITES'})

        # Doc node + edge (if doc_path is a real file)
        if nq.get('doc_path') and nq['doc_path'] != 'chat only':
            doc_label = nq['doc_path'].split('/')[-1]
            did = 'DOC_' + re.sub(r'[^A-Z0-9]+', '_', doc_label.upper()).strip('_')[:40]
            if did not in existing_node_ids:
                kg_data['nodes'].append({'id': did, 'type': 'Doc', 'label': doc_label})
                existing_node_ids.add(did)
            kg_data['edges'].append({'source': qnode_id, 'target': did, 'relation': 'ANSWERED_IN'})

    # Update KG meta
    type_ct = Counter(n['type'] for n in kg_data['nodes'])
    rel_ct = Counter(e['relation'] for e in kg_data['edges'])
    kg_data['meta']['n_nodes'] = len(kg_data['nodes'])
    kg_data['meta']['n_edges'] = len(kg_data['edges'])
    kg_data['meta']['node_types'] = dict(type_ct)
    kg_data['meta']['edge_relations'] = dict(rel_ct)
    kg_data['meta']['updated'] = datetime.now().strftime('%Y-%m-%d')

    # ---------- Save (light + heavy) ----------
    for path, data in [(QA_LIGHT, qa_data), (QA_HEAVY, qa_data)]:
        if path.parent.exists():
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f'Saved QA: {path}')
    for path, data in [(KG_LIGHT, kg_data), (KG_HEAVY, kg_data)]:
        if path.parent.exists():
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f'Saved KG: {path}')

    print()
    print('=== Summary ===')
    print(f'QA entries: {n_docs} (+5)')
    print(f'KG nodes: {len(kg_data["nodes"])} ({type_ct})')
    print(f'KG edges: {len(kg_data["edges"])} ({rel_ct})')
    print(f'IDF tokens: {len(idf)}')

if __name__ == '__main__':
    main()
