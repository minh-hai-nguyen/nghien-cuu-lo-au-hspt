# -*- coding: utf-8 -*-
"""
Glossary v3 — PEDAGOGICAL improvements.

Áp dụng nguyên tắc sư phạm:
1. SCAFFOLDING — thuật ngữ đơn giản trước phức tạp, tránh jargon circular
2. ANALOGY — phép ẩn dụ cho mọi khái niệm khó
3. COMMON MISCONCEPTIONS — nêu rõ lỗi phổ biến
4. LEARNING PATH — order để học
5. "WHEN TO USE A vs B" — phân biệt rõ các thuật ngữ dễ nhầm
6. CONCRETE → ABSTRACT — ví dụ đời thường trước công thức
"""
import os, sys, io, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
OUT_DIR = os.path.join(os.path.dirname(__file__), 'glossary_data')
WEB_DATA = os.path.join(ROOT, 'tro-ly-nghien-cuu-tam-ly', 'web', 'data')

# Load existing enhanced v2
with open(os.path.join(OUT_DIR, 'glossary_enhanced.json'), encoding='utf-8') as f:
    v2 = json.load(f)

ABBREV = v2['abbreviations']
DETAILED = dict(v2['detailed_terms'])
ORGS = dict(v2['organizations'])

# ============================================================
# PEDAGOGICAL ENHANCEMENTS — add fields to existing terms
# ============================================================
# Mỗi term sẽ có các field mới (nếu relevant):
# - analogy: phép ẩn dụ
# - misconception: sai lầm phổ biến
# - compare_to: so với thuật ngữ gần
# - exercise: bài tập tự kiểm tra

PEDAGOGY_UPDATES = {
    'GAD-7': {
        'analogy': 'Như "cân nhanh" chỉ số lo âu — 7 câu hỏi bấm xong trong 2 phút, cho điểm 0–21. '
                   'Không chẩn đoán nhưng giúp quyết định có cần khám sâu hơn không.',
        'misconception': 'SAI: "GAD-7 ≥ 10 = chẩn đoán GAD". ĐÚNG: GAD-7 chỉ SÀNG LỌC, không chẩn đoán. '
                         'Người đạt ngưỡng cần được phỏng vấn lâm sàng (DISC-5, MINI) để chẩn đoán chính thức.',
        'compare_to': 'GAD-7 (7 items, lo âu) ≠ PHQ-9 (9 items, trầm cảm). Hai thang thường đi cặp đôi '
                      'nhưng KHÔNG GỘP CHUNG điểm (xem memory feedback_gad7_phq9_separation).',
        'exercise': 'Thầy đọc: "HS có GAD-7 = 12". Câu hỏi: (a) mức độ lo âu? (b) có nên giới thiệu khám không? '
                    '→ (a) Moderate (10–14); (b) Nên giới thiệu.',
    },
    'SMD': {
        'analogy': 'Như "chia cho SD để đo bằng đơn vị chung" — giống như chuyển °F sang °C để so nhiệt độ '
                   '2 nước. SMD = hiệu ứng đo bằng "đơn vị SD" thay vì "điểm thang gốc".',
        'misconception': 'SAI: "SMD lớn = có ý nghĩa thống kê". ĐÚNG: SMD lớn = có Ý NGHĨA LÂM SÀNG. '
                         'Có ý nghĩa thống kê (CI không chứa 0) là chuyện khác. SMD = –0,05 có thể p < 0,001 '
                         'nhưng về lâm sàng là TRIVIAL.',
        'compare_to': 'SMD (chuẩn hoá) vs MD (thô): dùng SMD khi các NC đo bằng thang khác nhau '
                      '(GAD-7 vs SCARED vs STAI-C). Dùng MD khi TẤT CẢ đo cùng thang (hiếm gặp).',
        'exercise': 'Thầy đọc: "CBT giảm lo âu SMD = –0,30 (95% CI: –0,42; –0,18)". (a) độ lớn? (b) khoảng có '
                    'chứa 0 không? (c) ý nghĩa? → (a) Small (0,2–0,5); (b) Không; (c) Có ý nghĩa thống kê '
                    'và nhỏ về lâm sàng.',
    },
    'CI': {
        'analogy': 'Như "thả lưới đánh cá nhiều lần" — 95 % các lưới sẽ bắt được con cá lớn nhất, nhưng '
                   'lưới cụ thể trong tay thầy có thể bắt hoặc không (KHÔNG biết). Tuyên bố về PHƯƠNG PHÁP, '
                   'không phải về khoảng cụ thể.',
        'misconception': 'SAI (rất phổ biến): "Có 95 % xác suất giá trị thật trong CI này". ĐÚNG: Đó là định '
                         'nghĩa của CrI Bayesian, không phải CI frequentist. CI nói về procedure "nếu lặp lại", '
                         'không về một khoảng cụ thể. Tuy nhiên trong THỰC HÀNH, đa số đều hiểu loose như CrI.',
        'compare_to': 'CI (tần suất) vs CrI (Bayesian): cùng giá trị số trong 90% trường hợp, nhưng triết '
                      'học khác. Xem doc "Giai_thich_KTC_vs_CrI_cho_thay_v3".',
        'exercise': 'Thầy đọc: "SMD = –0,30, 95% CI: –0,42; –0,18". Câu: Đọc theo Bayesian cho dễ hiểu: '
                    '"Có ~95% khả năng SMD thực nằm giữa –0,42 và –0,18" → tuy về mặt triết học sai nhưng '
                    'thực hành OK.',
    },
    'CrI': {
        'analogy': 'Như "đoán vị trí cá trong ao" — có 95% khả năng con cá nằm trong khu vực tôi khoanh. '
                   'Tuyên bố trực tiếp về xác suất.',
        'misconception': 'SAI: "CrI luôn đúng hơn CI". ĐÚNG: Khi prior weakly informative (mặc định) và '
                         'mẫu đủ lớn, CrI VÀ CI có giá trị số gần như trùng nhau. CrI chỉ khác biệt '
                         'có ý nghĩa khi prior mạnh hoặc mẫu rất nhỏ.',
        'compare_to': 'Xem CI. Chỉ dùng CrI khi đọc Bayesian MA/NMA.',
        'exercise': 'Thầy đọc QT049 Zhang 2026: "SMD = –0,19 (95% CrI: –0,22; –0,17)". Câu: "Với 95% xác '
                    'suất, hiệu ứng thực nằm giữa –0,22 và –0,17" → diễn giải đúng. Khoảng KHÔNG chứa 0 → '
                    'có ý nghĩa thống kê. |SMD| < 0,2 → trivial về lâm sàng.',
    },
    "Cohen's d": {
        'analogy': 'Như "đo độ khác biệt 2 người, đơn vị là SD" — d = 1 nghĩa "người A cao hơn người B '
                   '1 độ lệch chuẩn" — tương đương ~84 % người B thấp hơn người A.',
        'misconception': 'SAI: "d = 0,5 luôn là medium". ĐÚNG: Cohen cutoffs là heuristic GUIDELINE, '
                         'không phải rule cứng. Trong 1 bối cảnh cụ thể (ví dụ can thiệp trường học '
                         'universal prevention), d = 0,2 có thể đã là "đáng giá đầu tư" do cost thấp.',
        'compare_to': "Cohen's d (gốc) vs Hedges' g (hiệu chỉnh small-sample) — khi n ≥ 50, d ≈ g.",
        'exercise': 'Ví dụ d = 0,11 (VN030 Happy House universal CBT). Câu: (a) độ lớn theo Cohen? '
                    '(b) ý nghĩa cho triển khai universal? → (a) Dưới trivial; (b) Không nên universal, '
                    'chuyển targeted.',
    },
    'MA': {
        'analogy': 'Như "hỏi 100 người đã ăn phở + bún, tính điểm trung bình của từng loại → kết luận '
                   'phở ngon hơn bao nhiêu". Pool ý kiến nhiều người thành 1 kết luận chung.',
        'misconception': 'SAI: "MA càng nhiều NC pool càng tốt". ĐÚNG: Quality > quantity. Nếu pool 10 NC có '
                         'heterogeneity cao (I² > 75%) → pooled estimate có thể MISLEADING. Thà pool 3 NC '
                         'đồng nhất còn hơn 30 NC khác biệt.',
        'compare_to': 'MA (2 can thiệp) vs NMA (≥ 3 can thiệp cùng lúc). Xem doc MA vs NMA.',
        'exercise': 'QT049 Zhang: 31 RCT CBT học đường → SMD –0,19. Câu: Câu hỏi MA trả lời là gì? '
                    '→ "Trung bình qua 31 NC, CBT học đường có giảm lo âu thực sự không?" (Trả lời: có '
                    'nhưng nhỏ).',
    },
    'NMA': {
        'analogy': 'Như "xếp hạng 9 quán phở Hà Nội" — không phải cả 9 đã được so trực tiếp, nhưng qua '
                   'mạng lưới reviews (quán 1 vs 2, 2 vs 3, 1 vs 5...), ta có thể xếp hạng cả 9.',
        'misconception': 'SAI: "NMA không cần RCT trực tiếp". ĐÚNG: NMA cần ÍT NHẤT một số RCT trực tiếp '
                         'để "khởi động" mạng. Nếu A chưa từng so với B ai hết (A là "node cô lập") → '
                         'không ước lượng được effect.',
        'compare_to': 'NMA yêu cầu thêm giả định TRANSITIVITY (bằng chứng gián tiếp đáng tin). Nếu các '
                      'NC về A vs B khác bối cảnh so với NC về B vs C (ví dụ khác độ tuổi), NMA có thể '
                      'sai lệch.',
        'exercise': 'QT039 Xian 2024: iCBT SUCRA 71,2% cho SAD. Câu: NMA thêm thông tin gì so với MA? '
                    '→ Xếp hạng (MA chỉ cho biết "A vs đối chứng", NMA cho biết "A so với TẤT CẢ").',
    },
    'DISC-5': {
        'analogy': 'Như "khám sức khoẻ đầy đủ" so với "cân + đo huyết áp" (sàng lọc) — DISC-5 là khám đầy đủ '
                   'theo tiêu chí DSM-5, cho chẩn đoán chính thức. SDQ/GAD-7 chỉ là cân nhanh.',
        'misconception': 'SAI: "Thang đo cao = chẩn đoán". ĐÚNG: Chỉ DISC-5/MINI/CIDI/K-SADS cho chẩn '
                         'đoán. SDQ/GAD-7/DASS chỉ sàng lọc — cao mới là CẦN khám sâu.',
        'compare_to': 'DISC-5 (structured, lay interviewer) vs K-SADS (semi-structured, clinician). '
                      'DISC-5 rẻ hơn, phù hợp NC dịch tễ lớn (như V-NAMHS).',
        'exercise': 'V-NAMHS dùng DISC-5 → 3,3% chẩn đoán disorder. VN001 dùng GAD-7 → 40,6% ≥ 5. Tại sao '
                    'khác biệt? → DISC-5 yêu cầu full DSM criteria; GAD-7 chỉ sàng lọc (nhiều false '
                    'positive).',
    },
    'CBT': {
        'analogy': 'Như "huấn luyện viên tâm lý" — giúp khách hàng nhận ra suy nghĩ không hợp lý (cognitive) '
                   'và thay đổi hành vi (behavioral). Kết hợp 2 mặt như chân + tay trong một cơ thể.',
        'misconception': 'SAI: "CBT chỉ là talk therapy". ĐÚNG: CBT có 4 thành phần rõ: psychoeducation, '
                         'cognitive restructuring, behavioural experiments, exposure. Có bài tập về nhà '
                         '— không chỉ "nói chuyện".',
        'compare_to': 'CBT (traditional) vs ACT (third wave, thêm acceptance+mindfulness) vs DBT (thêm '
                      'dialectics). CA-CBT = CBT + adapt văn hoá.',
        'exercise': 'Walkup 2008 CAMS: CBT + Sertraline > CBT alone > Sertraline alone > Placebo cho '
                    'anxiety 7-17 tuổi. Câu: Kết hợp CBT + thuốc có nên là first-line không? → Có, cho '
                    'moderate-severe anxiety theo guidelines NICE.',
    },
    'iCBT': {
        'analogy': 'Như "học online" so với "lớp trực tiếp" — iCBT là CBT qua internet. Có thể tự học '
                   '(unguided) hoặc có coach (guided). Scale up nhanh, cost thấp, phù hợp LMIC.',
        'misconception': 'SAI: "iCBT kém hơn face-to-face CBT". ĐÚNG: Meta-analyses cho thấy guided iCBT '
                         'hiệu quả TƯƠNG ĐƯƠNG face-to-face cho anxiety/depression nhẹ-vừa. Unguided '
                         'hiệu quả hơi thấp hơn.',
        'compare_to': 'iCBT (internet-based) vs Mobile CBT (app smartphone) — mobile là subset của '
                      'digital. Guided > Unguided (Walder 2025: g = 0,878 guided vs 0,3-0,4 unguided).',
        'exercise': 'QT045 Matsumoto Japan: iCBT unguided 8 modules → OR đáp ứng 4,97. Vì sao unguided '
                    'vẫn hiệu quả ở đây? → Vì subthreshold SAD (đơn giản hơn full SAD) + tự-motivated '
                    'đại học sinh Nhật Bản.',
    },
    'Universal vs Targeted prevention': {
        'analogy': 'Như "tiêm vaccine cho MỌI trẻ em" (universal) vs "tiêm vaccine cho TRẺ có phơi nhiễm" '
                   '(targeted). Cả 2 đều hợp lý tuỳ bệnh, cost-benefit.',
        'misconception': 'SAI: "Universal luôn tốt hơn vì cover nhiều người". ĐÚNG: Universal có diluting '
                         'effect — cần cost per case reduction CAO hơn targeted. Với CBT học đường '
                         '(Zhang 2026 SMD –0,19), universal kém cost-effective.',
        'compare_to': 'Universal (all) > Selective (nhóm rủi ro cao) > Indicated/Targeted (có triệu '
                      'chứng) > Treatment (đã chẩn đoán).',
        'exercise': 'Happy House VN030: universal, d = 0,11. BESST UK (QT042): targeted self-referral. '
                    'Câu: Nếu triển khai tại VN, chọn cái nào? → BESST-style targeted cho HS có GAD-7 ≥ 5 '
                    '(hiệu quả lớn hơn, cost thấp hơn).',
    },
    'Subthreshold symptoms': {
        'analogy': 'Như "tiền bệnh" — chưa đủ tiêu chí chẩn đoán nhưng đã có triệu chứng. Ví dụ: "tiền '
                   'đái tháo đường" (HbA1c 5,7-6,4) chưa đủ tiêu chí DM nhưng cần can thiệp.',
        'misconception': 'SAI: "Subthreshold không quan trọng, chỉ full disorder mới cần điều trị". ĐÚNG: '
                         'Subthreshold có 30-50% nguy cơ progression thành full disorder trong 2-3 năm. '
                         'Can thiệp sớm ở subthreshold cost-effective hơn điều trị full disorder.',
        'compare_to': 'Full threshold (đủ DSM criteria + impairment) > Subthreshold (≥50% triệu chứng) > '
                      'No threshold (không có triệu chứng đáng kể).',
        'exercise': 'V-NAMHS: 3,3% full disorder vs 21,7% any MH problem (subthreshold+). Tỷ lệ 7:1. '
                    'Câu: Ý nghĩa cho đề cương can thiệp VN? → Ưu tiên targeted subthreshold (pool '
                    'lớn 21,7%) thay vì chỉ treatment 3,3%.',
    },
}

# Apply pedagogy updates to existing terms
for key, updates in PEDAGOGY_UPDATES.items():
    if key in DETAILED:
        DETAILED[key].update(updates)
    else:
        # New term — shouldn't happen but safeguard
        DETAILED[key] = updates

# ============================================================
# ADD NEW DETAILED TERMS (theo góc nhìn sư phạm)
# ============================================================
NEW_DETAILED = {
    'I²': {
        'category': 'Tham số meta-analysis',
        'name_vn': 'Chỉ Số Heterogeneity I²',
        'explanation_vn': 'Đo % variance trong MA là do heterogeneity thực sự (khác biệt giữa NC) chứ '
                          'không phải do chance. 0% = tất cả NC đồng nhất; 100% = cực kỳ khác biệt.',
        'citation': 'Higgins JPT, Thompson SG (2002). Quantifying heterogeneity in a meta-analysis. '
                    'Statistics in Medicine, 21(11):1539–1558.',
        'wikipedia': 'https://en.wikipedia.org/wiki/I-squared',
        'example': 'QT029 Li NMA: I² = 45% → moderate heterogeneity (acceptable).',
        'threshold': '0-25% low; 25-75% moderate; >75% high. Nếu > 75% → pooled estimate có thể misleading.',
        'analogy': 'Như "độ phân tán của đám đông ý kiến" — I² 0% = tất cả ý kiến trùng; 100% = mỗi người '
                   'một phách. Khi I² cao → pool chung có nguy cơ đánh mất nuance.',
        'misconception': 'SAI: "I² > 50% thì MA không giá trị". ĐÚNG: I² cao KHÔNG phải là lý do loại MA, '
                         'mà là tín hiệu cần: (a) subgroup analysis; (b) meta-regression; (c) random-effects model.',
        'papers_in_csdl': ['Mọi MA/NMA trong CSDL'],
    },
    'Random-effects vs Fixed-effects': {
        'category': 'Mô hình meta-analysis',
        'name_vn': 'Mô Hình Hiệu Ứng Ngẫu Nhiên vs Cố Định',
        'explanation_vn': 'Fixed-effects: giả định TẤT CẢ NC đo cùng một hiệu ứng thực duy nhất. '
                          'Random-effects: mỗi NC có hiệu ứng thực riêng, rút từ phân phối chung.',
        'citation': 'Borenstein M et al. (2010). A basic introduction to fixed-effect and random-effects '
                    'models for meta-analysis. Research Synthesis Methods, 1(2):97–111.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Meta-analysis#Random_effects_model',
        'analogy': 'Fixed-effects: như "đo chiều cao 1 người nhiều lần — có 1 giá trị thật". Random-effects: '
                   'như "đo chiều cao của NHIỀU người — mỗi người có giá trị thật khác nhau, nhưng đều '
                   'rút từ 1 quần thể chung".',
        'misconception': 'SAI: "Random-effects luôn đúng hơn". ĐÚNG: Random-effects là mặc định trong '
                         'psychology MA vì heterogeneity cao, nhưng nếu I² ≈ 0 thì fixed-effects đủ và '
                         'hiệu quả hơn.',
        'compare_to': 'Fixed: CI hẹp hơn, dễ significant nhưng chỉ valid nếu I² thấp. Random: CI rộng hơn, '
                      'thận trọng hơn, chuẩn APA 7 khuyến nghị.',
        'papers_in_csdl': ['Mọi MA/NMA'],
    },
    'p-value': {
        'category': 'Test statistics',
        'name_vn': 'Giá Trị p',
        'explanation_vn': 'Xác suất quan sát được data (hoặc extreme hơn) NẾU null hypothesis (H0 = không '
                          'có hiệu ứng) ĐÚNG. p < 0,05 = nếu H0 đúng, dữ liệu này hiếm gặp (< 5%) → '
                          'loại H0.',
        'citation': 'Fisher RA (1925). Statistical Methods for Research Workers. Oliver & Boyd.',
        'wikipedia': 'https://en.wikipedia.org/wiki/P-value',
        'example': 'QT038 De Silva Sri Lanka: β = −0,096, p = 0,038 → có ý nghĩa thống kê.',
        'threshold': 'p < 0,05 = có ý nghĩa; p < 0,01 = highly significant; p < 0,001 = very highly',
        'analogy': 'Như "xác suất tung đồng xu ra mặt X 10 lần liên tiếp NẾU đồng xu công bằng". Xác suất '
                   'cực thấp (< 5%) → nghi ngờ đồng xu không công bằng (loại H0).',
        'misconception': 'SAI (RẤT phổ biến): "p = 0,03 nghĩa là 3% khả năng H0 đúng". ĐÚNG: p là xác '
                         'suất observe data GIVEN H0 đúng (p(Data|H0)), không phải xác suất H0 đúng '
                         'GIVEN data (p(H0|Data)). Đây là sai lầm nền tảng của inferential statistics.',
        'compare_to': 'p-value (frequentist) vs Bayes Factor (Bayesian). BF cho tỷ lệ bằng chứng 2 '
                      'giả thuyết trực tiếp, trực quan hơn.',
        'exercise': 'QT049 Zhang: SMD = –0,19, p < 0,001. Câu: Có ý nghĩa lâm sàng không? → Không! '
                    'p < 0,001 là ý nghĩa THỐNG KÊ. |SMD| < 0,2 nên trivial LÂM SÀNG.',
        'papers_in_csdl': ['Mọi NC định lượng'],
    },
    'Statistical vs Clinical significance': {
        'category': 'Khái niệm nền tảng',
        'name_vn': 'Ý Nghĩa Thống Kê vs Ý Nghĩa Lâm Sàng',
        'explanation_vn': 'Statistical significance: kết quả KHÔNG do chance (p < 0,05). Clinical '
                          'significance: kết quả ĐỦ LỚN để có ích trong thực hành lâm sàng. HAI VIỆC '
                          'KHÁC NHAU.',
        'citation': 'Kirk RE (1996). Practical significance: A concept whose time has come. '
                    'Educational and Psychological Measurement, 56(5):746–759.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Clinical_significance',
        'analogy': 'Như "phát hiện đồng xu nặng hơn 0,001 g so với quy chuẩn" — statistical significant '
                   'nếu n cực lớn, nhưng KHÔNG clinical significant (không ai nhận ra được).',
        'misconception': 'SAI: "p < 0,05 nghĩa là can thiệp nên triển khai". ĐÚNG: Phải có CẢ hai: '
                         'statistical (p < 0,05) + clinical (SMD đủ lớn, NNT hợp lý).',
        'exercise': 'Zhang 2026 Bayesian MA: SMD = –0,19 (95% CrI: –0,22; –0,17). Có ý nghĩa thống kê? '
                    'Có ý nghĩa lâm sàng? → Thống kê CÓ (CrI không chứa 0). Lâm sàng KHÔNG (|SMD| < 0,2 '
                    'trivial).',
        'papers_in_csdl': ['Mọi MA/NMA'],
    },
    'PRISMA flow diagram': {
        'category': 'Chuẩn báo cáo SR',
        'name_vn': 'Sơ Đồ Dòng Chảy PRISMA',
        'explanation_vn': 'Sơ đồ bắt buộc trong SR/MA hiện đại — mô tả quá trình chọn lọc NC: records '
                          'identified → duplicates removed → screened → assessed → included.',
        'citation': 'Page MJ et al. (2021). PRISMA 2020 flow diagram. BMJ, 372:n71.',
        'wikipedia': 'https://en.wikipedia.org/wiki/PRISMA',
        'analogy': 'Như "biểu đồ lọc nước" — nước bẩn ở đầu → qua N bộ lọc → nước sạch ở cuối. Mỗi '
                   'bước lọc có số lượng bị loại và lý do.',
        'example': 'QT049 Zhang 2026: Records identified 2.156 → screened 1.842 → assessed 187 → '
                   'included 31 RCT.',
        'papers_in_csdl': ['Mọi SR/MA trong CSDL'],
    },
    'Transitivity (NMA assumption)': {
        'category': 'Giả định NMA',
        'name_vn': 'Giả Định Transitivity',
        'explanation_vn': 'Trong NMA, giả định các NC về A vs B có thể kết hợp với NC về B vs C để suy '
                          'ra A vs C — chỉ đúng nếu A vs B và B vs C được thực hiện ở BỐI CẢNH tương tự '
                          '(cùng độ tuổi, severity, setting).',
        'citation': 'Salanti G (2012). Indirect and mixed-treatment comparison, network, or '
                    'multiple-treatments meta-analysis. Research Synthesis Methods, 3(2):80–97.',
        'analogy': 'Như "hai người chưa gặp nhau nhưng cùng quen Bob". Ta có thể suy đoán họ giống Bob '
                   'ở vài điểm, nhưng chỉ nếu 2 người gặp Bob ở cùng HOÀN CẢNH (công sở vs đám cưới).',
        'misconception': 'SAI: "Transitivity tự động đúng". ĐÚNG: Phải kiểm tra trước khi làm NMA. Nếu '
                         'vi phạm → incoherence → kết quả không đáng tin.',
        'papers_in_csdl': ['QT029', 'QT039'],
    },
    'Publication bias': {
        'category': 'Sai lệch nghiên cứu',
        'name_vn': 'Sai Lệch Xuất Bản',
        'explanation_vn': 'Hiện tượng NC có kết quả DƯƠNG TÍNH dễ được publish hơn NC âm tính. Dẫn đến '
                          'MA overestimate effect size (vì thiếu NC âm tính trong pool).',
        'citation': 'Rothstein HR, Sutton AJ, Borenstein M (2005). Publication Bias in Meta-Analysis: '
                    'Prevention, Assessment and Adjustments. Wiley.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Publication_bias',
        'analogy': 'Như "tin tức chỉ đưa tin tai nạn, không đưa tin ngày bình thường" — mắt ta dễ thấy '
                   'sự kiện đáng nhớ hơn sự kiện "không xảy ra gì".',
        'misconception': 'SAI: "MA đủ tốt thì không có publication bias". ĐÚNG: Publication bias hiện diện '
                         'gần như trong MỌI MA. Cần đánh giá bằng funnel plot + Egger test.',
        'example': 'QT048 Chen 2025: "publication bias ở mức thấp qua toàn bộ các NC" — đã đánh giá '
                   'và công bố.',
        'compare_to': 'Funnel plot: asymmetric → có thể có publication bias. Egger test: p < 0,1 → '
                      'flag bias.',
        'papers_in_csdl': ['QT048', 'mọi SR/MA chất lượng'],
    },
    'Blinding (masking)': {
        'category': 'Phương pháp giảm bias',
        'name_vn': 'Mù Hoá (Làm Mù)',
        'explanation_vn': 'Giữ BN và/hoặc người đánh giá KHÔNG BIẾT ai thuộc nhóm can thiệp vs đối chứng. '
                          'Giảm bias kỳ vọng + placebo effect.',
        'citation': 'Schulz KF, Grimes DA (2002). Blinding in randomised trials. Lancet, 359(9307):'
                    '696–700.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Blinded_experiment',
        'analogy': 'Như "thi không tên, không danh" — người chấm không biết bài của ai → không bias '
                   'theo cảm tình.',
        'misconception': 'SAI: "Psychology intervention không thể mù hoá được". ĐÚNG: BN không thể mù '
                         'với psychotherapy (biết mình đang nhận CBT), nhưng NGƯỜI ĐÁNH GIÁ outcome có '
                         'thể mù (assessor-blind).',
        'compare_to': 'Single-blind (chỉ BN mù), double-blind (BN + clinician), triple-blind (+ người '
                      'phân tích).',
        'papers_in_csdl': ['RCT trong CSDL'],
    },
    'Allocation concealment': {
        'category': 'Phương pháp giảm bias',
        'name_vn': 'Che Giấu Phân Bổ',
        'explanation_vn': 'Khác biệt với blinding — che giấu KẾ HOẠCH phân bổ (ngẫu nhiên) trước khi '
                          'BN được enroll, để investigator không thể manipulate.',
        'citation': 'Schulz KF, Chalmers I, Hayes RJ, Altman DG (1995). Empirical evidence of bias. '
                    'JAMA, 273(5):408–412.',
        'analogy': 'Như "không được xem phiếu rút thăm trước khi học sinh vào phòng thi" — tránh thiên '
                   'vị trong việc ai được vào đội A, ai vào đội B.',
        'misconception': 'SAI: "Allocation concealment = blinding". ĐÚNG: Khác nhau! Allocation = trước '
                         'enroll; blinding = sau enroll.',
        'papers_in_csdl': ['Cochrane RoB 2 kiểm tra'],
    },
    'Cluster RCT': {
        'category': 'Thiết kế RCT',
        'name_vn': 'Thử Nghiệm Ngẫu Nhiên Theo Cụm',
        'explanation_vn': 'Đơn vị ngẫu nhiên là CỤM (trường, lớp, bệnh viện) chứ không phải cá nhân. Phù '
                          'hợp khi can thiệp tác động cả nhóm (ví dụ: thay đổi curriculum cả lớp).',
        'citation': 'Donner A, Klar N (2000). Design and Analysis of Cluster Randomization Trials in '
                    'Health Research. Arnold.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Cluster_randomised_controlled_trial',
        'analogy': 'Như "bốc thăm NHÓM học, không phải cá nhân" — nếu bốc cá nhân, HS trong cùng lớp '
                   'có thể ảnh hưởng lẫn nhau (contamination).',
        'misconception': 'SAI: "Cluster RCT cho phép n nhỏ". ĐÚNG: Cluster RCT cần nhiều người hơn '
                         'individual RCT do intra-cluster correlation (ICC). Cỡ mẫu cần × DE (Design '
                         'Effect) = 1 + (m-1)×ICC.',
        'example': 'QT038 De Silva Sri Lanka: 36 trường (18 CT + 18 ĐC) × 20 HS/trường = 720. '
                   'Nếu individual RCT cần chỉ ~300.',
        'papers_in_csdl': ['QT038', 'VN030'],
    },
    'Attrition bias': {
        'category': 'Sai lệch NC',
        'name_vn': 'Sai Lệch Do Mất Mẫu',
        'explanation_vn': 'Khi người tham gia drop-out KHÔNG NGẪU NHIÊN (ví dụ: chỉ người cải thiện mới '
                          'ở lại follow-up), kết quả bị lệch.',
        'citation': 'Higgins JPT, Altman DG (2011). Assessing risk of bias in included studies. '
                    'Cochrane Handbook 5.1, Chapter 8.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Selection_bias',
        'analogy': 'Như "lớp học hỏi ai hài lòng — những người bỏ học đã về mất rồi" — kết luận sẽ '
                   'quá tích cực.',
        'misconception': 'SAI: "Drop-out < 20% là OK". ĐÚNG: Còn phụ thuộc lý do drop-out. Nếu drop-out '
                         'có HỆ THỐNG (ví dụ người nặng bỏ) → bias bất kể %.',
        'compare_to': 'ITT (Intention-To-Treat) analysis khắc phục một phần: analyze theo nhóm '
                      'được assign, dù không hoàn thành.',
        'papers_in_csdl': ['VN030 (mất mẫu <3%, ITT ok)'],
    },
    'Sensitivity vs Specificity': {
        'category': 'Psychometric',
        'name_vn': 'Tính Nhạy vs Tính Đặc Hiệu',
        'explanation_vn': 'Sensitivity: tỷ lệ người BỆNH được thang đo phát hiện (true positive rate). '
                          'Specificity: tỷ lệ người KHÔNG BỆNH được loại trừ đúng (true negative rate).',
        'citation': 'Altman DG, Bland JM (1994). Diagnostic tests 1: sensitivity and specificity. BMJ, '
                    '308(6943):1552.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Sensitivity_and_specificity',
        'example': 'GAD-7 ngưỡng 10: sensitivity 89%, specificity 82% — tức 89% người GAD được phát hiện '
                   '+ 82% người không GAD được loại trừ.',
        'analogy': 'Sensitivity = "lưới bắt cá không sót con nào" (bỏ lỡ ít); Specificity = "lưới không '
                   'bắt nhầm rác" (false positive ít). Hai tính chất thường TRADE-OFF — tăng sensitivity '
                   'thì giảm specificity.',
        'misconception': 'SAI: "Sensitivity càng cao càng tốt". ĐÚNG: Phải cân bằng với specificity tuỳ '
                         'bối cảnh. Screening: ưu tiên sensitivity (không bỏ sót). Diagnosis: ưu tiên '
                         'specificity (không false positive).',
        'compare_to': 'PPV (Positive Predictive Value): trong số người test +, bao nhiêu % thực sự bệnh. '
                      'Phụ thuộc prevalence quần thể.',
        'papers_in_csdl': ['VN001 (GAD-7 ngưỡng)'],
    },
    'Reliability vs Validity': {
        'category': 'Psychometric',
        'name_vn': 'Độ Tin Cậy vs Độ Giá Trị',
        'explanation_vn': 'Reliability: thang đo CONSISTENT không (đo lại có kết quả tương tự?). Validity: '
                          'thang đo đo ĐÚNG cái cần đo không?',
        'citation': 'Kline P (2000). The Handbook of Psychological Testing (2nd ed.). Routledge.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Reliability_(statistics)',
        'analogy': 'Reliability = "cân cứ cân là ra số y nhau" (consistent). Validity = "cân đo thực sự '
                   'là cân nặng, không phải chiều cao". Có thể reliable mà không valid (cân lệch nhưng '
                   'luôn lệch cùng kiểu).',
        'misconception': 'SAI: "Cronbach α cao = thang tốt". ĐÚNG: Cronbach α cao chỉ chứng minh '
                         'RELIABILITY, không phải validity. Thang có α = 0,9 vẫn có thể đo SAI construct.',
        'example': 'SDQ-25 tiếng Việt (Weiss 2014): đã validate trong NC nationally representative — '
                   'có cả reliability + validity VN context.',
        'compare_to': 'Types of validity: content, construct (convergent + discriminant), criterion '
                      '(concurrent + predictive).',
        'papers_in_csdl': ['Mọi thang đo tâm lý'],
    },
    'Cronbach alpha': {
        'category': 'Psychometric',
        'name_vn': 'Hệ Số Cronbach Alpha',
        'explanation_vn': 'Đo internal consistency reliability của thang nhiều item — các items có đo '
                          'cùng construct không. 0–1, càng cao càng tốt.',
        'citation': 'Cronbach LJ (1951). Coefficient alpha and the internal structure of tests. '
                    'Psychometrika, 16(3):297–334.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Cronbach%27s_alpha',
        'example': 'VN001 Hoa 2024: GAD-7 Cronbach α = 0,916 trong mẫu HS THPT Hà Nội → độ tin cậy rất cao.',
        'threshold': 'α < 0,6 poor; 0,6–0,7 acceptable; 0,7–0,8 good; 0,8–0,9 very good; > 0,9 excellent',
        'analogy': 'Như "hỏi 7 câu cùng chủ đề, xem câu trả lời có consistent không" — nếu consistent '
                   'thì α cao.',
        'misconception': 'SAI: "α = 0,95 tốt nhất". ĐÚNG: α > 0,95 có thể là REDUNDANCY — các items hỏi '
                         'quá giống nhau, nên cắt bớt.',
        'papers_in_csdl': ['VN001 α=0,916; nhiều NC validate'],
    },
}

# Merge new detailed terms
DETAILED.update(NEW_DETAILED)

# ============================================================
# ADD LEARNING PATH
# ============================================================
LEARNING_PATH = {
    'level_1_foundation': {
        'title': 'Tầng 1 — Nền tảng (đọc trước hết)',
        'description': 'Những khái niệm nền mà MỌI thuật ngữ khác đều dựa vào.',
        'terms': ['M (Mean)', 'SD', 'p-value', 'Statistical vs Clinical significance',
                  'Reliability vs Validity', 'DSM-5'],
    },
    'level_2_effect_size': {
        'title': 'Tầng 2 — Độ lớn hiệu ứng',
        'description': 'Hiểu cách đo lường hiệu ứng của can thiệp.',
        'terms': ["Cohen's d", "Hedges' g", 'SMD', 'MD', 'OR', 'NNT'],
    },
    'level_3_intervals': {
        'title': 'Tầng 3 — Khoảng ước lượng',
        'description': 'Không chắc chắn xung quanh ước lượng.',
        'terms': ['CI', 'CrI', 'Sensitivity vs Specificity', 'Cronbach alpha'],
    },
    'level_4_designs': {
        'title': 'Tầng 4 — Thiết kế NC',
        'description': 'Cách thiết kế để có bằng chứng mạnh.',
        'terms': ['RCT', 'Cluster RCT', 'Blinding (masking)', 'Allocation concealment',
                  'Attrition bias'],
    },
    'level_5_synthesis': {
        'title': 'Tầng 5 — Tổng hợp bằng chứng',
        'description': 'Meta-analysis, network MA, chuẩn báo cáo.',
        'terms': ['MA', 'NMA', 'I²', 'Random-effects vs Fixed-effects',
                  'Transitivity (NMA assumption)', 'Publication bias', 'SUCRA', 'PRISMA',
                  'PRISMA flow diagram', 'GRADE', 'RoB 2'],
    },
    'level_6_concepts': {
        'title': 'Tầng 6 — Khái niệm can thiệp',
        'description': 'Các loại can thiệp và cách lựa chọn.',
        'terms': ['CBT', 'iCBT', 'Universal vs Targeted prevention', 'Subthreshold symptoms',
                  'MHPSS', 'mhGAP'],
    },
    'level_7_diagnoses': {
        'title': 'Tầng 7 — Chẩn đoán cụ thể',
        'description': 'Các rối loạn tâm thần quan trọng.',
        'terms': ['GAD', 'SAD', 'MDD', 'ICD-11'],
    },
}

# ============================================================
# ADD QUICK START GUIDE
# ============================================================
QUICK_START = {
    'common_queries': [
        {
            'question': 'Tôi đọc SMD = –0,19 (95 % CrI: –0,22; –0,17). Ý nghĩa?',
            'answer_steps': [
                '1. SMD = Standardized Mean Difference = hiệu ứng chuẩn hoá',
                '2. CrI = Credible Interval (Bayesian), đọc như "có 95% xác suất giá trị thật trong khoảng"',
                '3. Số âm = can thiệp giảm triệu chứng (so với đối chứng)',
                '4. |SMD| = 0,19 < 0,2 theo Cohen → TRIVIAL / dưới small',
                '5. CrI không chứa 0 → có ý nghĩa thống kê',
                '6. Kết luận: Có ý nghĩa thống kê nhưng TRIVIAL về lâm sàng',
            ],
            'related_terms': ['SMD', 'CrI', "Cohen's d", 'Statistical vs Clinical significance'],
        },
        {
            'question': 'OR = 4,97 nghĩa gì?',
            'answer_steps': [
                '1. OR = Odds Ratio = tỷ số odds',
                '2. OR > 1 = tăng odds; OR < 1 = giảm odds',
                '3. OR = 4,97 ≈ 5 lần — odds của outcome ở nhóm can thiệp CAO 5× so với đối chứng',
                '4. Trong QT045 Matsumoto: OR response = 4,97 cho iCBT vs waitlist → iCBT mạnh',
                '5. Lưu ý: OR khác RR (Risk Ratio). OR phóng đại khi tỷ lệ biến cố cao',
            ],
            'related_terms': ['OR', 'RR', 'QT045'],
        },
        {
            'question': 'NMA với SUCRA 71% là gì?',
            'answer_steps': [
                '1. NMA = Network Meta-Analysis = tổng hợp so sánh NHIỀU can thiệp cùng lúc',
                '2. SUCRA = Surface Under Cumulative Ranking curve, thang 0–100%',
                '3. SUCRA 71% = can thiệp này có 71% xác suất xếp hạng CAO trong nhóm',
                '4. QT039 Xian: iCBT SUCRA 71,2% = iCBT có 71% xác suất là can thiệp tốt nhất cho SAD',
                '5. Ngưỡng: > 70% = xếp hạng cao đáng tin cậy',
            ],
            'related_terms': ['NMA', 'SUCRA', 'Bayesian'],
        },
    ],
}

# ============================================================
# BUILD FINAL
# ============================================================
glossary_v3 = {
    'meta': {
        'created': '2026-04-15',
        'version': 'v3-pedagogical',
        'pedagogical_features': [
            'Scaffolding (learning path 7 levels)',
            'Analogies for 20+ abstract concepts',
            'Common misconceptions called out',
            'When-to-use-A-vs-B comparisons',
            'Exercises for self-check',
            'Quick start Q&A',
        ],
        'total_abbreviations': len(ABBREV),
        'total_detailed_terms': len(DETAILED),
        'total_organizations': len(ORGS),
    },
    'abbreviations': ABBREV,
    'detailed_terms': DETAILED,
    'organizations': ORGS,
    'learning_path': LEARNING_PATH,
    'quick_start': QUICK_START,
}

# Save
out1 = os.path.join(OUT_DIR, 'glossary_v3_pedagogical.json')
with open(out1, 'w', encoding='utf-8') as f:
    json.dump(glossary_v3, f, ensure_ascii=False, indent=2)
print(f'Saved: {out1}')

out2 = os.path.join(WEB_DATA, 'glossary.json')
with open(out2, 'w', encoding='utf-8') as f:
    json.dump(glossary_v3, f, ensure_ascii=False, indent=2)
print(f'Saved: {out2}')

print(f'\nAbbreviations: {len(ABBREV)}')
print(f'Detailed terms (with pedagogy): {len(DETAILED)}')
print(f'Organizations: {len(ORGS)}')
print(f'Learning path levels: {len(LEARNING_PATH)}')
print(f'Quick start Q&A: {len(QUICK_START["common_queries"])}')
print(f'File size: {os.path.getsize(out1):,} bytes')
