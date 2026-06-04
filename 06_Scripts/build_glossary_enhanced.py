# -*- coding: utf-8 -*-
"""
Enhanced glossary v2 — add citations, Wikipedia explanations, examples.

Structure:
- abbreviations: dict[short → full name + 1-line VN]
- detailed_terms: dict[term → rich object {citation, wikipedia_url, vn_explanation, example, papers_in_csdl}]
- organizations: rich org profiles
- authors: integrated from author_kg
"""
import os, sys, io, json
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
OUT_DIR = os.path.join(os.path.dirname(__file__), 'glossary_data')
WEB_DATA = os.path.join(ROOT, 'tro-ly-nghien-cuu-tam-ly', 'web', 'data')
os.makedirs(OUT_DIR, exist_ok=True)

# Load existing (v1) abbreviations
with open(os.path.join(OUT_DIR, 'glossary_full.json'), encoding='utf-8') as f:
    v1 = json.load(f)
ABBREV = v1['abbreviations']
ORGS_V1 = v1['organizations']

# ============================================================
# DETAILED TERMS with citation + example + Wikipedia
# ============================================================
# Each entry: {category, name_vn, explanation_vn, citation, wikipedia, example, threshold, papers_in_csdl}

DETAILED = {
    # === STATISTICAL PARAMETERS ===
    'GAD-7': {
        'category': 'Thang đo (sàng lọc lo âu)',
        'name_vn': 'Thang Sàng Lọc Rối Loạn Lo Âu Lan Toả 7 mục',
        'explanation_vn': '7 câu hỏi về triệu chứng lo âu 2 tuần qua. Mỗi câu 0–3 điểm (từ "không có" đến "gần như mỗi ngày"). Tổng 0–21 điểm. Tính nhạy 89 %, tính đặc hiệu 82 % tại ngưỡng 10.',
        'citation': 'Spitzer RL, Kroenke K, Williams JB, Löwe B (2006). A brief measure for assessing generalized anxiety disorder: the GAD-7. Archives of Internal Medicine, 166(10):1092–1097.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Generalized_Anxiety_Disorder_7',
        'example': 'Trong VN001 (Hoa 2024), GAD-7 ≥ 5 được 40,6 % HS THPT Hà Nội đạt → tỷ lệ lo âu ít nhất mức nhẹ trở lên.',
        'threshold': 'Tổng: 0–4 không/thấp; 5–9 nhẹ; 10–14 vừa; ≥ 15 nặng',
        'papers_in_csdl': ['VN001', 'VN002 (đo phụ huynh)'],
    },
    'PHQ-9': {
        'category': 'Thang đo (sàng lọc trầm cảm)',
        'name_vn': 'Thang Sàng Lọc Trầm Cảm 9 mục',
        'explanation_vn': '9 câu tương ứng 9 tiêu chí DSM-IV cho MDD. Mỗi câu 0–3 điểm (2 tuần qua). Tổng 0–27.',
        'citation': 'Kroenke K, Spitzer RL, Williams JB (2001). The PHQ-9: validity of a brief depression severity measure. Journal of General Internal Medicine, 16(9):606–613.',
        'wikipedia': 'https://en.wikipedia.org/wiki/PHQ-9',
        'example': 'V-NAMHS đo PHQ-9 của phụ huynh để đánh giá SKTT gia đình — yếu tố ảnh hưởng lo âu con.',
        'threshold': '0–4 không; 5–9 nhẹ; 10–14 vừa; 15–19 vừa-nặng; ≥ 20 nặng',
        'papers_in_csdl': ['VN002 (phụ huynh)'],
    },
    'DASS-21': {
        'category': 'Thang đo đa chiều',
        'name_vn': 'Thang Trầm Cảm – Lo Âu – Stress 21 mục',
        'explanation_vn': 'Phiên bản rút gọn của DASS-42. 3 subscale × 7 items: Trầm cảm, Lo âu, Stress. Mỗi subscale 0–42 sau khi nhân 2.',
        'citation': 'Lovibond PF, Lovibond SH (1995). The structure of negative emotional states. Behaviour Research and Therapy, 33(3):335–343.',
        'wikipedia': 'https://en.wikipedia.org/wiki/DASS_(psychology)',
        'example': 'VN029 (Duong 2025) đo DASS-21 ở 2.631 HS THPT TPHCM → lo âu 50,3 %.',
        'threshold': 'Lo âu: 0–7 normal; 8–9 mild; 10–14 moderate; 15–19 severe; ≥ 20 extremely severe',
        'papers_in_csdl': ['VN020', 'VN029'],
    },
    'SDQ-25': {
        'category': 'Thang đo đa chiều (trẻ em)',
        'name_vn': 'Bảng Hỏi Điểm Mạnh và Điểm Yếu 25 mục',
        'explanation_vn': '25 câu chia 5 subscale × 5 items: emotional problems, conduct problems, hyperactivity/inattention, peer problems, prosocial. Self-report (11–17), parent/teacher (4–17).',
        'citation': 'Goodman R (1997). The Strengths and Difficulties Questionnaire: A research note. Journal of Child Psychology and Psychiatry, 38(5):581–586.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Strengths_and_Difficulties_Questionnaire',
        'example': 'VN022 UNICEF dùng SDQ-25 trên 668 HS → 26,1 % rủi ro trung/cao vấn đề SKTT.',
        'threshold': 'Total difficulties (tổng 4 subscale đầu, 0–40): 0–13 close to average; 14–16 slightly raised; 17–19 high; ≥ 20 very high',
        'papers_in_csdl': ['VN022', 'VN027', 'Weiss 2014 (cited trong VN002)'],
    },
    'DISC-5': {
        'category': 'Công cụ chẩn đoán (diagnostic interview)',
        'name_vn': 'Bảng Phỏng Vấn Chẩn Đoán cho Trẻ Em, Phiên bản 5',
        'explanation_vn': 'Phỏng vấn cấu trúc hoàn toàn cho trẻ 6–18. Lay-interviewer có thể thực hiện. Cho chẩn đoán DSM-5 chính thức (khác với sàng lọc SDQ/GAD-7).',
        'citation': 'Bitsko R, Adams HR, Holbrook J, et al. (2019). Diagnostic Interview Schedule for Children, Version 5 (DISC-5): Development and Validation of ADHD and Tic Disorder Modules. JAACAP, 58(10): Supplement. | Shaffer D, Fisher P, Lucas CP, et al. (2000). NIMH DISC-IV. JAACAP, 39(1):28–38.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Diagnostic_Interview_Schedule_for_Children',
        'example': 'V-NAMHS dùng DISC-5 trên 5.996 cặp → 3,3 % VTN đáp ứng tiêu chí DSM-5 cho rối loạn tâm thần.',
        'threshold': 'Chẩn đoán theo DSM-5: đáp ứng đủ criteria = full threshold; đáp ứng ≥ 50 % criteria = subthreshold',
        'papers_in_csdl': ['VN002'],
    },
    'CESD-R': {
        'category': 'Thang đo trầm cảm',
        'name_vn': 'Thang Trầm Cảm CES-D Phiên bản Sửa đổi',
        'explanation_vn': '20 câu về triệu chứng trong tuần qua. Revision 2004 hiện đại hoá với DSM-IV. Tổng 0–80.',
        'citation': 'Eaton WW, Smith C, Ybarra M, Muntaner C, Tien A (2004). Center for Epidemiologic Studies Depression Scale: Review and Revision (CESD and CESD-R). In: Maruish ME (Ed.), The Use of Psychological Testing for Treatment Planning and Outcomes Assessment, 3rd Ed., Mahwah NJ: Erlbaum, pp. 363–377.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Center_for_Epidemiologic_Studies_Depression_Scale',
        'example': 'VN030 Happy House dùng CESD-R ≥ 16 là primary outcome → baseline 25,5 % HS đạt ngưỡng trầm cảm.',
        'threshold': 'Cut-off ≥ 16 = có khả năng trầm cảm lâm sàng',
        'papers_in_csdl': ['VN030'],
    },

    # === EFFECT SIZES ===
    "Cohen's d": {
        'category': 'Tham số effect size',
        'name_vn': "Hệ số Cohen's d (độ lớn hiệu ứng chuẩn hoá)",
        'explanation_vn': 'Đo khác biệt giữa 2 mean chia cho độ lệch chuẩn pooled. Không đơn vị. Dùng so sánh effect xuyên thang đo.',
        'citation': 'Cohen J (1988). Statistical Power Analysis for the Behavioral Sciences (2nd ed.). Routledge.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Effect_size#Cohen\'s_d',
        'example': 'VN030 Happy House: d = 0,11 cho CESD-R (depression) — effect size rất nhỏ (dưới trivial).',
        'threshold': '< 0,2 trivial; 0,2 small; 0,5 medium; 0,8 large; 2,0 huge (Sawilowsky 2009)',
        'formula': 'd = (M₁ – M₂) / SD_pooled',
        'papers_in_csdl': ['VN030', 'QT043 Bress d=1,04', 'nhiều NC quốc tế'],
    },
    "Hedges' g": {
        'category': 'Tham số effect size',
        'name_vn': "Hệ số Hedges' g (SMD hiệu chỉnh mẫu nhỏ)",
        'explanation_vn': 'Tương tự Cohen\'s d nhưng nhân thêm correction factor J(df) để giảm bias khi n nhỏ. CHUẨN COCHRANE cho meta-analysis.',
        'citation': 'Hedges LV, Olkin I (1985). Statistical Methods for Meta-Analysis. Academic Press.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Effect_size#Hedges\'_g',
        'example': 'QT040 Walder 2025: g = 0,878 cho guided DMHI SAD-specific (large effect).',
        'threshold': 'Giống Cohen d: < 0,2 trivial; 0,2 small; 0,5 medium; 0,8 large',
        'formula': 'g = d × J(df), J(df) = 1 − 3/(4·df − 1)',
        'papers_in_csdl': ['QT040', 'chuẩn Cochrane cho mọi MA'],
    },
    'SMD': {
        'category': 'Tham số effect size',
        'name_vn': 'Khác Biệt Trung Bình Chuẩn Hoá',
        'explanation_vn': 'Dùng khi các NC đo bằng thang khác nhau. SMD pool được xuyên scales. Trong Cochrane MA, SMD mặc định là Hedges\' g.',
        'citation': 'Cochrane Handbook v6.4 (2023), Chapter 10.5: The standardized mean difference.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Standardized_mean_difference',
        'example': 'QT049 Zhang 2026: SMD anxiety −0,19 [95 % CrI: −0,22; −0,17] từ 31 RCT với GAD-7, SCARED, STAI-C, RCADS...',
        'threshold': 'Giống Cohen d. |SMD| < 0,2 = trivial về lâm sàng dù có ý nghĩa thống kê',
        'papers_in_csdl': ['QT049', 'QT029', 'QT040'],
    },
    'OR': {
        'category': 'Tham số effect size',
        'name_vn': 'Tỷ Số Odds',
        'explanation_vn': 'Đo sự thay đổi odds của biến cố (có bệnh / không bệnh) giữa 2 nhóm. OR = 1: không khác; > 1: tăng nguy cơ; < 1: bảo vệ.',
        'citation': 'Bland JM, Altman DG (2000). Statistics Notes: The odds ratio. BMJ, 320(7247):1468.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Odds_ratio',
        'example': 'QT045 Matsumoto 2024: OR đáp ứng = 4,97 cho iCBT vs waitlist — tức iCBT làm tăng odds đáp ứng 5× so với chờ.',
        'threshold': 'Chen 2010 cho LMIC health: OR 1,5 small; 2,5 medium; 4,0 large',
        'formula': 'OR = (a·d)/(b·c) từ bảng 2×2',
        'papers_in_csdl': ['QT045', 'QT048', 'nhiều NC epidemiology'],
    },
    'CI': {
        'category': 'Khoảng ước lượng',
        'name_vn': 'Khoảng Tin Cậy (Frequentist)',
        'explanation_vn': '"Nếu lặp lại NC vô hạn lần, 95 % CI tính được sẽ chứa giá trị thật." Tuyên bố về procedure, không về khoảng cụ thể.',
        'citation': 'Neyman J (1937). Outline of a theory of statistical estimation based on the classical theory of probability. Philosophical Transactions Royal Society A, 236(767):333–380.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Confidence_interval',
        'example': 'QT040 Walder 2025: g = 0,878 (95 % CI: 0,62; 1,14) cho DMHI guided SAD.',
        'papers_in_csdl': ['Đa số NC frequentist trong CSDL'],
    },
    'CrI': {
        'category': 'Khoảng ước lượng',
        'name_vn': 'Khoảng Tin Cậy Bayesian',
        'explanation_vn': '"Có 95 % xác suất giá trị thật nằm trong khoảng này" — tuyên bố trực tiếp về probability, dễ diễn giải lâm sàng hơn CI.',
        'citation': 'Kruschke JK (2015). Doing Bayesian Data Analysis (2nd ed.), Chapter 12. Academic Press.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Credible_interval',
        'example': 'QT049 Zhang 2026: SMD = −0,19 (95 % CrI: −0,22; −0,17) từ Bayesian MA 31 RCT.',
        'papers_in_csdl': ['QT029', 'QT039', 'QT049'],
    },
    'NNT': {
        'category': 'Tham số lâm sàng',
        'name_vn': 'Số Người Cần Điều Trị để 1 Người Cải Thiện',
        'explanation_vn': 'Nếu NNT = 10 → cần điều trị 10 người mới có 1 người cải thiện (so với không điều trị). NNT thấp = hiệu quả cao.',
        'citation': 'Laupacis A, Sackett DL, Roberts RS (1988). An assessment of clinically useful measures of the consequences of treatment. NEJM, 318(26):1728–1733.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Number_needed_to_treat',
        'example': 'NNT 5–10 = rất tốt (điều trị cấp tính); NNT 10–30 = acceptable cho prevention.',
        'formula': 'NNT = 1 / (Risk_control − Risk_treatment)',
        'papers_in_csdl': ['Hay dùng trong review Cochrane'],
    },
    'SUCRA': {
        'category': 'Tham số NMA',
        'name_vn': 'Diện Tích Dưới Đường Xếp Hạng Tích Luỹ',
        'explanation_vn': 'Tính từ Bayesian NMA: xác suất can thiệp X được xếp hạng cao. 0 % = chắc chắn xếp hạng cuối; 100 % = chắc chắn xếp hạng đầu.',
        'citation': 'Salanti G, Ades AE, Ioannidis JPA (2011). Graphical methods and numerical summaries for presenting results from multiple-treatment meta-analysis. Journal of Clinical Epidemiology, 64(2):163–171.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Network_meta-analysis',
        'example': 'QT039 Xian 2024: iCBT SUCRA 71,2 % cho SAD — tức iCBT có 71 % xác suất là can thiệp tốt nhất trong 9 liệu pháp.',
        'threshold': '> 70 % = xếp hạng cao đáng tin cậy',
        'papers_in_csdl': ['QT029', 'QT039'],
    },

    # === RESEARCH DESIGNS ===
    'RCT': {
        'category': 'Thiết kế NC',
        'name_vn': 'Thử Nghiệm Ngẫu Nhiên Có Đối Chứng',
        'explanation_vn': 'Phân bổ NGẪU NHIÊN người tham gia vào nhóm can thiệp vs đối chứng → chuẩn vàng cho suy diễn nhân quả. Nếu thiết kế tốt, khác biệt quan sát là DO can thiệp, không phải do bias.',
        'citation': 'Hill AB (1952). The clinical trial. NEJM, 247(4):113–119. | Chalmers I (2001). Comparing like with like. BMJ, 322(7281):278.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Randomized_controlled_trial',
        'example': 'QT038 De Silva 2024 Sri Lanka: cluster RCT 720 HS lớp 9, 36 trường → β = −0,096, p = 0,038 cho CBT do GV dạy.',
        'papers_in_csdl': ['QT038', 'QT043', 'QT045', '30+ RCT trong reviews'],
    },
    'MA': {
        'category': 'Thiết kế NC (tổng hợp)',
        'name_vn': 'Phân Tích Tổng Hợp',
        'explanation_vn': 'Pool hiệu ứng từ nhiều RCT/NC cùng chủ đề → ra 1 effect size tổng hợp với CI/CrI. Cho phép tăng statistical power.',
        'citation': 'Glass GV (1976). Primary, secondary, and meta-analysis of research. Educational Researcher, 5(10):3–8. | Borenstein M et al. (2021). Introduction to Meta-Analysis (2nd ed.). Wiley.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Meta-analysis',
        'example': 'QT049 Zhang 2026: MA 31 RCT CBT học đường → SMD anxiety −0,19 (small effect, nhưng có ý nghĩa thống kê).',
        'papers_in_csdl': ['QT040', 'QT044', 'QT048', 'QT049'],
    },
    'NMA': {
        'category': 'Thiết kế NC (tổng hợp mạng)',
        'name_vn': 'Phân Tích Tổng Hợp Mạng',
        'explanation_vn': 'So sánh NHIỀU can thiệp cùng lúc (≥ 3) qua mạng. Sử dụng cả bằng chứng TRỰC TIẾP (A vs B có RCT) và GIÁN TIẾP (A vs C suy ra qua A-B-C) để ước lượng mọi cặp. Cho phép xếp hạng.',
        'citation': 'Lumley T (2002). Network meta-analysis for indirect treatment comparisons. Statistics in Medicine, 21(16):2313–2324. | Dias S et al. (2018). Network Meta-Analysis for Decision-Making. Wiley.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Network_meta-analysis',
        'example': 'QT039 Xian 2024: NMA 30 RCT × 9 liệu pháp cho SAD VTN → iCBT SUCRA 71,2 % xếp hạng 1.',
        'papers_in_csdl': ['QT029', 'QT039'],
    },

    # === INTERVENTIONS ===
    'CBT': {
        'category': 'Can thiệp tâm lý',
        'name_vn': 'Liệu Pháp Nhận Thức – Hành Vi',
        'explanation_vn': 'Kết hợp kỹ thuật nhận thức (cognitive restructuring, thay đổi niềm tin không thích hợp) và kỹ thuật hành vi (exposure, behavioural activation). Evidence-based hàng đầu cho lo âu và trầm cảm.',
        'citation': 'Beck AT (1976). Cognitive Therapy and the Emotional Disorders. International Universities Press. | Hofmann SG et al. (2012). The Efficacy of Cognitive Behavioral Therapy. Cognitive Therapy and Research, 36(5):427–440.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Cognitive_behavioral_therapy',
        'example': 'Walkup 2008 CAMS (NEJM): CBT + sertraline hiệu quả hơn CBT đơn thuần hoặc sertraline đơn thuần cho anxiety 7–17 tuổi.',
        'papers_in_csdl': ['Cite trong hầu hết reviews (QT028, QT029, QT049)', 'Walkup 2008'],
    },
    'iCBT': {
        'category': 'Can thiệp số (digital)',
        'name_vn': 'CBT qua Internet',
        'explanation_vn': 'CBT được cung cấp qua website/platform. GUIDED (có trị liệu viên hỗ trợ qua chat) hiệu quả hơn UNGUIDED (tự học). Cho phép scale up cost-effective.',
        'citation': 'Andersson G, Titov N (2014). Advantages and limitations of Internet-based interventions for common mental disorders. World Psychiatry, 13(1):4–11.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Internet-based_treatments_for_trauma_survivors',
        'example': 'QT045 Matsumoto 2024 Japan: iCBT unguided 8 modules → OR đáp ứng 4,97 cho SAD subthreshold VTN/SV.',
        'papers_in_csdl': ['QT040', 'QT045', 'QT050'],
    },
    'MHPSS': {
        'category': 'Khung can thiệp',
        'name_vn': 'Hỗ Trợ Sức Khoẻ Tâm Thần và Tâm Lý Xã Hội',
        'explanation_vn': 'Khung 4 tầng của WHO/IASC: (1) Services — điều trị chuyên biệt; (2) Focused supports — nhóm hỗ trợ có trained worker; (3) Community supports — gia đình, peer; (4) Basic services — bảo đảm nhu cầu cơ bản.',
        'citation': 'Inter-Agency Standing Committee (IASC) (2007). Guidelines on Mental Health and Psychosocial Support in Emergency Settings. Geneva: IASC.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Psychosocial_intervention',
        'example': 'QT051 Menon 2025 scoping review 69 NC MHPSS LMIC Đông Á-TBD; VN022 UNICEF.',
        'papers_in_csdl': ['VN022', 'VN030', 'QT051'],
    },
    'mhGAP': {
        'category': 'Chương trình WHO',
        'name_vn': 'Chương Trình Hành Động Lấp Khoảng Trống Sức Khoẻ Tâm Thần',
        'explanation_vn': 'Hướng dẫn WHO cho non-specialist health workers triển khai chăm sóc SKTT cơ bản ở LMIC. Bao gồm depression, anxiety, psychosis, epilepsy, substance use, dementia.',
        'citation': 'World Health Organization (2016). mhGAP Intervention Guide for Mental, Neurological and Substance Use Disorders in Non-Specialized Health Settings — Version 2.0. Geneva: WHO.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Mental_Health_Gap_Action_Programme',
        'example': 'V-NAMHS và các khuyến nghị chính sách VN dẫn chiếu mhGAP để lồng ghép SKTT vào y tế cơ sở.',
        'papers_in_csdl': ['VN002', 'nhiều NC LMIC'],
    },

    # === DIAGNOSES ===
    'DSM-5': {
        'category': 'Hệ thống chẩn đoán',
        'name_vn': 'Sổ Tay Chẩn Đoán và Thống Kê Rối Loạn Tâm Thần, Phiên bản 5',
        'explanation_vn': 'Chuẩn chẩn đoán tâm thần của Hiệp hội Tâm thần Hoa Kỳ. Criteria-based (đáp ứng ≥ N triệu chứng trong X thời gian + impairment). Rà soát 2022 (DSM-5-TR).',
        'citation': 'American Psychiatric Association (2013). Diagnostic and Statistical Manual of Mental Disorders (5th ed.). Washington DC: APA Publishing. | APA (2022). DSM-5-TR.',
        'wikipedia': 'https://en.wikipedia.org/wiki/DSM-5',
        'example': 'DISC-5 dựa trên DSM-5. V-NAMHS dùng DISC-5 → 3,3 % VTN VN đáp ứng tiêu chí DSM-5 cho ít nhất 1 rối loạn.',
        'papers_in_csdl': ['VN002', 'QT028', 'QT045'],
    },
    'ICD-11': {
        'category': 'Hệ thống chẩn đoán',
        'name_vn': 'Phân Loại Quốc Tế Bệnh Tật, Phiên bản 11',
        'explanation_vn': 'Chuẩn WHO, có hiệu lực từ 01/01/2022. Mental disorders ở Chapter 6. Khác DSM-5 ở một số category (gộp gaming disorder, prolonged grief disorder, complex PTSD).',
        'citation': 'World Health Organization (2022). International Classification of Diseases, 11th Revision (ICD-11). Geneva: WHO.',
        'wikipedia': 'https://en.wikipedia.org/wiki/ICD-11',
        'example': 'VN vẫn chủ yếu dùng ICD-10 (F00–F99); ICD-11 đang được triển khai.',
        'papers_in_csdl': ['Tham chiếu trong VN002, QT051'],
    },
    'GAD': {
        'category': 'Chẩn đoán',
        'name_vn': 'Rối Loạn Lo Âu Lan Toả',
        'explanation_vn': 'Lo âu quá mức, khó kiểm soát, về nhiều sự kiện, ≥ 6 tháng. Kèm ≥ 3/6 triệu chứng thể chất (căng thẳng cơ, mệt mỏi, khó tập trung, cáu gắt, rối loạn giấc ngủ, bồn chồn). DSM-5 code F41.1, ICD-11 code 6B00.',
        'citation': 'APA (2013). DSM-5, p. 222–226.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Generalized_anxiety_disorder',
        'example': 'V-NAMHS DISC-5 module GAD → gộp với social phobia thành "anxiety disorders" 2,3 %.',
        'papers_in_csdl': ['VN002', 'cite trong QT028'],
    },
    'SAD': {
        'category': 'Chẩn đoán',
        'name_vn': 'Rối Loạn Lo Âu Xã Hội (Ám ảnh Xã Hội)',
        'explanation_vn': 'Sợ hãi có tính dai dẳng về 1 hoặc nhiều tình huống xã hội mà người đó bị người khác quan sát. Sợ bị đánh giá tiêu cực → né tránh hoặc chịu đựng với distress lớn. Ở VTN: phải trong peer-setting. DSM-5 F40.10, ICD-11 6B04.',
        'citation': 'APA (2013). DSM-5, p. 202–208.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Social_anxiety_disorder',
        'example': 'QT039 Xian NMA 30 RCT về điều trị SAD ở trẻ/VTN (n = 1.547) → iCBT xếp hạng 1 SUCRA 71,2 %.',
        'papers_in_csdl': ['QT039', 'QT040', 'QT043', 'QT045'],
    },
    'MDD': {
        'category': 'Chẩn đoán',
        'name_vn': 'Rối Loạn Trầm Cảm Nặng',
        'explanation_vn': '≥ 5/9 triệu chứng trong 2 tuần: (1) tâm trạng buồn hầu hết mỗi ngày; (2) anhedonia (mất hứng thú); (3) thay đổi cân nặng; (4) rối loạn giấc ngủ; (5) tâm động cơ thể; (6) mệt mỏi; (7) cảm giác vô dụng; (8) khó tập trung; (9) ý nghĩ tự sát. Triệu chứng (1) hoặc (2) BẮT BUỘC. DSM-5 F32.x.',
        'citation': 'APA (2013). DSM-5, p. 160–168.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Major_depressive_disorder',
        'example': 'V-NAMHS: 0,9 % VTN VN đáp ứng DSM-5 cho MDD.',
        'papers_in_csdl': ['VN002', 'VN030'],
    },

    # === REPORTING STANDARDS ===
    'PRISMA': {
        'category': 'Chuẩn báo cáo',
        'name_vn': 'Các Mục Báo Cáo Ưu Tiên cho SR/MA',
        'explanation_vn': '27-item checklist + flow diagram chuẩn 2020 update cho SR/MA. BẮT BUỘC ở các tạp chí Cochrane, BMJ, Lancet.',
        'citation': 'Page MJ, McKenzie JE, Bossuyt PM, et al. (2021). The PRISMA 2020 statement: an updated guideline for reporting systematic reviews. BMJ, 372:n71.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Preferred_Reporting_Items_for_Systematic_Reviews_and_Meta-Analyses',
        'example': 'QT029, QT037, QT039, QT048, QT049, QT051 — tất cả SR/MA trong CSDL đều tuân PRISMA 2020.',
        'papers_in_csdl': ['QT028–QT051 (SR/MA)'],
    },
    'GRADE': {
        'category': 'Chuẩn báo cáo',
        'name_vn': 'Xếp Hạng Đánh Giá Phát Triển và Đánh Giá Khuyến Nghị',
        'explanation_vn': 'Hệ thống đánh giá CHẤT LƯỢNG bằng chứng thành 4 mức: High, Moderate, Low, Very Low. Hạ bậc vì: RoB, inconsistency, indirectness, imprecision, publication bias.',
        'citation': 'Guyatt GH, Oxman AD, Vist GE, et al. (2008). GRADE: an emerging consensus on rating quality of evidence and strength of recommendations. BMJ, 336(7650):924–926.',
        'wikipedia': 'https://en.wikipedia.org/wiki/GRADE_approach',
        'example': 'QT049 Zhang 2026 tự kết luận evidence ở mức low-moderate → khuyến nghị cần high-quality trials trước khi triển khai rộng.',
        'papers_in_csdl': ['Các SR/MA chất lượng cao'],
    },
    'RoB 2': {
        'category': 'Chuẩn báo cáo',
        'name_vn': 'Cochrane Đánh Giá Rủi Ro Sai Lệch v2',
        'explanation_vn': 'Đánh giá 5 domain: (1) randomization; (2) deviations; (3) missing outcome data; (4) measurement; (5) reported result selection. Mỗi domain: Low / Some concerns / High.',
        'citation': 'Sterne JAC, Savović J, Page MJ, et al. (2019). RoB 2: a revised tool for assessing risk of bias in randomised trials. BMJ, 366:l4898.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Cochrane_Collaboration',
        'example': 'QT029 Li 2025 Bayesian NMA đánh giá RoB 2 cho 30 RCT: ~40 % low risk, 40 % some concerns, 20 % high.',
        'papers_in_csdl': ['QT029', 'QT038', 'QT039', 'QT049'],
    },

    # === KEY CONCEPTS ===
    'Universal vs Targeted prevention': {
        'category': 'Khái niệm can thiệp',
        'name_vn': 'Phòng Ngừa Phổ Quát vs Có Mục Tiêu',
        'explanation_vn': 'Universal: cho TẤT CẢ (không chọn lọc theo nguy cơ). Targeted/Indicated: cho người có triệu chứng hoặc nguy cơ cao.',
        'citation': 'Mrazek PJ, Haggerty RJ (Eds.) (1994). Reducing Risks for Mental Disorders: Frontiers for Preventive Intervention Research. National Academy Press. | Gordon RS (1983). An operational classification of disease prevention. Public Health Reports, 98(2):107–109.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Preventive_medicine',
        'example': 'QT049 Zhang 2026: universal CBT school SMD anxiety −0,19 (small effect). QT042 BESST UK: targeted self-referral hiệu quả hơn universal.',
        'papers_in_csdl': ['VN030 (universal)', 'QT042 (targeted)', 'QT049 (universal)'],
    },
    'Subthreshold symptoms': {
        'category': 'Khái niệm lâm sàng',
        'name_vn': 'Triệu Chứng Dưới Ngưỡng',
        'explanation_vn': 'Đáp ứng ≥ 50 % số triệu chứng DSM nhưng CHƯA đủ để chẩn đoán full disorder. Nhóm này VẪN cần can thiệp vì có nguy cơ progressing thành full disorder.',
        'citation': 'Pincus HA, Davis WW, McQueen LE (1999). "Subthreshold" mental disorders. A review and synthesis of studies. British Journal of Psychiatry, 174:288–296.',
        'wikipedia': 'https://en.wikipedia.org/wiki/Subthreshold_conditions',
        'example': 'V-NAMHS: 21,7 % "mental health problem" (subthreshold+) vs 3,3 % "mental disorder" (full DSM-5) — khoảng cách 7× cho thấy tầm quan trọng của subthreshold.',
        'papers_in_csdl': ['VN002', 'QT045 Matsumoto (subthreshold SAD)'],
    },
}

# ============================================================
# EXPANDED ORGANIZATIONS
# ============================================================
ORGS_DETAILED = dict(ORGS_V1)  # Start with v1 orgs

# Add more orgs
ORGS_DETAILED['Cochrane'] = {
    'name_full': 'Cochrane (formerly Cochrane Collaboration)',
    'name_vn': 'Tổ chức Cochrane',
    'country': 'International (HQ London, UK)',
    'role': 'Global nonprofit producing systematic reviews on healthcare interventions. Publishes Cochrane Handbook (v6.4, 2023) — chuẩn SR/MA methodology cao nhất. Ra đời 1993.',
    'citation': 'Higgins JPT et al. (2023). Cochrane Handbook for Systematic Reviews of Interventions v6.4. Cochrane.',
    'website': 'https://www.cochrane.org/',
    'papers_related': ['Chuẩn methodology cho tất cả SR/MA trong CSDL'],
}
ORGS_DETAILED['APA'] = {
    'name_full': 'American Psychiatric Association (APA)',
    'name_vn': 'Hiệp hội Tâm thần Hoa Kỳ',
    'country': 'USA',
    'role': 'Xuất bản DSM-5 (2013), DSM-5-TR (2022). Nonprofit thành lập 1844.',
    'citation': 'APA (2013). DSM-5. Washington DC: APA Publishing.',
    'website': 'https://www.psychiatry.org/',
    'papers_related': ['Mọi bài chẩn đoán DSM trong CSDL'],
}

# ============================================================
# BUILD FINAL JSON
# ============================================================
enhanced = {
    'meta': {
        'created': '2026-04-15',
        'version': 'v2-enhanced',
        'total_abbreviations': len(ABBREV),
        'total_detailed_terms': len(DETAILED),
        'total_organizations': len(ORGS_DETAILED),
    },
    'abbreviations': ABBREV,
    'detailed_terms': DETAILED,
    'organizations': ORGS_DETAILED,
}

# Save to 06_Scripts
out1 = os.path.join(OUT_DIR, 'glossary_enhanced.json')
with open(out1, 'w', encoding='utf-8') as f:
    json.dump(enhanced, f, ensure_ascii=False, indent=2)
print(f'Saved: {out1}')

# Save to web/data (overwrite v1)
out2 = os.path.join(WEB_DATA, 'glossary.json')
with open(out2, 'w', encoding='utf-8') as f:
    json.dump(enhanced, f, ensure_ascii=False, indent=2)
print(f'Saved: {out2}')

print(f'\nAbbreviations: {len(ABBREV)}')
print(f'Detailed terms (with citations): {len(DETAILED)}')
print(f'Organizations: {len(ORGS_DETAILED)}')
print(f'File size: {os.path.getsize(out1):,} bytes')
