# -*- coding: utf-8 -*-
"""Sinh ban nhap Q2 song ngu English-Vietnamese - Frontiers in Psychiatry.
Adapted from Q1 v7 (01/06/2026) per thay MD decision (07/06/2026):
- Q1 (BMC Psychiatry) -> Q2 (Frontiers in Psychiatry)
- Bo mixed-methods, chi giu pure quantitative SEM
- Phu hop guideline Frontiers in Psychiatry
SEM Beta verified vs LA Bang 27, 30, 33, 36, 39, 42, 45.
Combined R^2=0,598 tu Bang 47 = INTEGRATED MODEL ANSWER cho Q1-8 BLOCKING.
07/06/2026."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
OUT = os.path.join(ROOT, 'bai-bao-Q1', 'Draft_Q2_SongNgu_v1_07062026.docx')

d = Document()
for sec in d.sections:
    sec.top_margin = Cm(2.0); sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.0); sec.right_margin = Cm(2.0)
s = d.styles['Normal']
s.font.name = 'Times New Roman'; s.font.size = Pt(12)
s.paragraph_format.line_spacing = 1.5


def H1(en, vi):
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(20); p.paragraph_format.space_after = Pt(6)
    r = p.add_run(en); r.font.name = 'Times New Roman'; r.font.size = Pt(15); r.bold = True
    r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run(vi); r.font.name = 'Times New Roman'; r.font.size = Pt(13); r.bold = True; r.italic = True
    r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

def H2(en, vi):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(14); p.paragraph_format.space_after = Pt(3)
    r = p.add_run(en); r.font.name = 'Times New Roman'; r.font.size = Pt(13); r.bold = True
    r.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
    p = d.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(vi); r.font.name = 'Times New Roman'; r.font.size = Pt(12); r.bold = True; r.italic = True
    r.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

def H3(en, vi):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(2)
    r = p.add_run(en); r.font.name = 'Times New Roman'; r.font.size = Pt(12); r.bold = True
    p = d.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(vi); r.font.name = 'Times New Roman'; r.font.size = Pt(11); r.bold = True; r.italic = True

def PP(en, vi, indent=True):
    """English paragraph followed by Vietnamese italic paragraph."""
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(3)
    if indent: p.paragraph_format.first_line_indent = Cm(0.5)
    r = p.add_run(en); r.font.name = 'Times New Roman'; r.font.size = Pt(12)
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(8)
    if indent: p.paragraph_format.first_line_indent = Cm(0.5)
    r = p.add_run(vi); r.font.name = 'Times New Roman'; r.font.size = Pt(11); r.italic = True
    r.font.color.rgb = RGBColor(0x40, 0x40, 0x40)

def NOTE(en, vi):
    """Small note paragraph - grey italic."""
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Cm(0.3)
    r = p.add_run(en); r.font.name = 'Times New Roman'; r.font.size = Pt(10); r.italic = True
    r.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.left_indent = Cm(0.3)
    r = p.add_run(vi); r.font.name = 'Times New Roman'; r.font.size = Pt(10); r.italic = True
    r.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

def TBD(text_en, text_vi):
    """Placeholder for BLOCKING items - red bold."""
    p = d.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5); p.paragraph_format.space_after = Pt(6)
    r = p.add_run(f'[TBD - {text_en}]'); r.font.name = 'Times New Roman'; r.font.size = Pt(11); r.bold = True
    r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    p = d.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5); p.paragraph_format.space_after = Pt(6)
    r = p.add_run(f'[Cho - {text_vi}]'); r.font.name = 'Times New Roman'; r.font.size = Pt(11); r.bold = True; r.italic = True
    r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

def set_col_widths(t, widths_cm):
    for row in t.rows:
        for idx, w in enumerate(widths_cm):
            if idx < len(row.cells):
                row.cells[idx].width = Cm(w)


# ============================================================
# COVER PAGE
# ============================================================
p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(24); p.paragraph_format.space_after = Pt(12)
r = p.add_run('Draft Q2 - Frontiers in Psychiatry (v1, bilingual)')
r.font.name = 'Times New Roman'; r.font.size = Pt(11); r.italic = True
r.font.color.rgb = RGBColor(0x70, 0x70, 0x70)

H1('Integrated risk-protective structural equation modelling of '
   'anxiety disorder subtypes among Vietnamese lower secondary school '
   'students: A cross-sectional structural equation modeling study',
   'Mô hình phương trình cấu trúc tích hợp các yếu tố nguy cơ – bảo vệ '
   'đối với các phân loại rối loạn lo âu ở học sinh trung học cơ sở '
   'Việt Nam: nghiên cứu cắt ngang sử dụng mô hình phương trình cấu trúc')

p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run('Hang Thi Cong¹*, Duc Minh Dao¹, Nguyen Minh Duc²†')
r.font.name = 'Times New Roman'; r.font.size = Pt(12); r.bold = True

p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(12)
r = p.add_run('¹ Hanoi National University of Education, Vietnam | '
              '² [Affiliation TBD] | * Corresponding author | '
              '† Senior author (last position, per Western convention)')
r.font.name = 'Times New Roman'; r.font.size = Pt(10); r.italic = True


# ============================================================
H2('ABSTRACT', 'TÓM TẮT')

H3('Background', 'Đặt vấn đề')
PP('Anxiety disorders represent a leading mental health concern among '
   'adolescents globally, with the Asian region showing the largest '
   'age-standardised increase between 1990 and 2021. Despite a substantial '
   'evidence base on risk and protective factors studied in isolation, '
   'to our knowledge no Vietnamese study has tested an integrated '
   'risk-protective model across DSM-5 anxiety disorder subtypes '
   '(generalised anxiety disorder, GAD; separation anxiety disorder, '
   'SAD; social anxiety disorder, SocAD) using structural equation '
   'modelling (SEM). The present study addresses this gap and examines '
   'gender invariance of these subtype-specific pathways.',
   'Rối loạn lo âu là một trong những vấn đề sức khỏe tâm thần hàng đầu '
   'ở thanh thiếu niên trên toàn cầu, và khu vực châu Á ghi nhận mức gia '
   'tăng chuẩn hóa theo tuổi cao nhất trong giai đoạn 1990–2021. Mặc dù '
   'đã có một nền bằng chứng đáng kể về tác động đơn lẻ của các yếu tố '
   'nguy cơ và bảo vệ, theo hiểu biết của chúng tôi vẫn chưa có nghiên '
   'cứu nào tại Việt Nam kiểm định mô hình tích hợp nguy cơ – bảo vệ '
   'trên các phân loại rối loạn lo âu theo DSM-5 (lo âu tổng quát – '
   'GAD, lo âu chia ly – SAD, lo âu xã hội – SocAD) bằng mô hình phương '
   'trình cấu trúc (SEM). Nghiên cứu này lấp khoảng trống đó và phân '
   'tích tính bất biến giới của các đường dẫn theo từng phân loại.')

H3('Methods', 'Phương pháp')
PP('A cross-sectional SEM study was conducted with 1,352 lower '
   'secondary school students (614 boys, 738 girls; aged 11–14 years) '
   'from two purposively selected schools in Hanoi, Vietnam. Eight '
   'validated scales measured three risk factors — academic pressure '
   '(ESSA), smartphone addiction (SAS-SV), school bullying (OBVQ) — '
   'and four protective factors — school engagement (PSSM), parental '
   'support (MSPSS-Parents), peer support (MSPSS-Peers) and '
   'self-esteem (RSES). DSM-5 anxiety subtypes were assessed with the '
   'Revised Children\'s Anxiety and Depression Scale (RCADS; GAD, SAD, '
   'SocAD subscales). Subtype-specific SEM with each of the seven '
   'predictors, an integrated risk-protective composite model, and '
   'multi-group invariance testing by gender were performed in '
   'AMOS 31.0.',
   'Một nghiên cứu cắt ngang sử dụng SEM được thực hiện với 1.352 học '
   'sinh trung học cơ sở (614 nam, 738 nữ; 11–14 tuổi) tại hai trường '
   'ở Hà Nội được chọn có chủ đích. Tám thang đo đã kiểm định đo lường '
   'ba yếu tố nguy cơ — áp lực học tập (ESSA), nghiện điện thoại '
   '(SAS-SV), bắt nạt học đường (OBVQ) — và bốn yếu tố bảo vệ — gắn '
   'bó trường học (PSSM), hỗ trợ cha mẹ (MSPSS – cha mẹ), hỗ trợ bạn '
   'bè (MSPSS – bạn bè) và tự trọng (RSES). Các phân loại lo âu theo '
   'DSM-5 được đo bằng thang RCADS (tiểu thang GAD, SAD, SocAD). Các '
   'mô hình SEM theo từng phân loại với từng yếu tố trong số bảy yếu '
   'tố dự báo, mô hình tích hợp nguy cơ – bảo vệ, và kiểm định bất '
   'biến đa nhóm theo giới được thực hiện trong AMOS 31.0.')

H3('Results', 'Kết quả')
PP('Subtype-specific SEM revealed three key differential pathways. '
   'First, school bullying showed its strongest association with SAD '
   '(β = 0.376, p < 0.001) rather than GAD (β = 0.215) or SocAD '
   '(β = 0.253), a counterintuitive pattern not previously reported in '
   'the Vietnamese literature. Second, self-esteem emerged as the '
   'strongest single protective influence on GAD (β = −0.455, p < 0.001), '
   'comparable in magnitude to the dominant academic-pressure risk '
   'effect on GAD (β = 0.510). Third, peer support showed no significant '
   'association with any anxiety subtype (β = −0.015 to −0.079), '
   'challenging the assumed universal protective role of peer '
   'relationships in adolescence. Integrated risk-protective modelling '
   'with composite latent factors yielded R² = 0.598 for total anxiety '
   '(YTNC composite β = 0.669, p < 0.001; YTBV composite β = −0.220, '
   'p < 0.001). Multi-group invariance testing revealed significant '
   'gender differences for GAD (female 59.47 > male 51.43; F = 44.48, '
   'p < 0.001) and SocAD (female 52.74 > male 43.20; F = 45.98, '
   'p < 0.001) but full gender invariance for SAD (female 24.76 vs male '
   '25.42; F = 0.246, p = 0.620).',
   'Mô hình SEM theo từng phân loại làm rõ ba đường dẫn khác biệt then '
   'chốt. Thứ nhất, bắt nạt học đường có liên hệ mạnh nhất với SAD '
   '(β = 0,376; p < 0,001) so với GAD (β = 0,215) hoặc SocAD '
   '(β = 0,253) — một mẫu hình phản trực giác chưa từng được báo cáo '
   'trong y văn Việt Nam. Thứ hai, tự trọng nổi lên là yếu tố bảo vệ '
   'đơn lẻ mạnh nhất đối với GAD (β = −0,455; p < 0,001), tương đương '
   'với tác động chiếm ưu thế của áp lực học tập lên GAD (β = 0,510). '
   'Thứ ba, hỗ trợ bạn bè không có liên hệ có ý nghĩa với bất kỳ phân '
   'loại lo âu nào (β = −0,015 đến −0,079), thách thức giả định về vai '
   'trò bảo vệ phổ quát của quan hệ bạn bè ở thanh thiếu niên. Mô hình '
   'tích hợp nguy cơ – bảo vệ với các nhân tố tiềm ẩn tổng hợp đạt '
   'R² = 0,598 cho tổng lo âu (β tổng YTNC = 0,669; β tổng YTBV '
   '= −0,220; cả hai p < 0,001). Kiểm định bất biến đa nhóm cho thấy '
   'khác biệt giới có ý nghĩa với GAD (nữ 59,47 > nam 51,43; F = 44,48; '
   'p < 0,001) và SocAD (nữ 52,74 > nam 43,20; F = 45,98; p < 0,001) '
   'nhưng bất biến giới hoàn toàn với SAD (nữ 24,76 vs nam 25,42; '
   'F = 0,246; p = 0,620).')

H3('Conclusions', 'Kết luận')
PP('In a large Vietnamese lower-secondary sample, school bullying '
   'showed its strongest impact on separation anxiety, self-esteem '
   'served as the strongest single protective lever against '
   'generalised anxiety, and separation anxiety alone exhibited '
   'complete gender invariance — consistent with a cultural-collectivism '
   'interpretation in which family hierarchy creates uniform attachment '
   'experiences across genders. These findings support subtype-targeted '
   'prevention frameworks that incorporate culture-sensitive '
   'self-esteem building and bullying-specific intervention for '
   'separation anxiety.',
   'Trong một mẫu lớn học sinh trung học cơ sở Việt Nam, bắt nạt học '
   'đường có tác động mạnh nhất tới lo âu chia ly, tự trọng là đòn bẩy '
   'bảo vệ đơn lẻ mạnh nhất chống lại lo âu tổng quát, và chỉ riêng lo '
   'âu chia ly thể hiện bất biến giới hoàn toàn — phù hợp với cách '
   'diễn giải văn hóa tập thể, trong đó hệ thống thứ bậc gia đình tạo '
   'trải nghiệm gắn bó đồng nhất giữa các giới. Các phát hiện này ủng '
   'hộ các khung phòng ngừa theo từng phân loại, tích hợp xây dựng tự '
   'trọng nhạy với văn hóa và can thiệp đặc thù cho bắt nạt với lo âu '
   'chia ly.')

H3('Keywords', 'Từ khóa')
PP('adolescent anxiety; DSM-5 subtypes; structural equation modelling; '
   'gender invariance; Vietnam; bullying; self-esteem; school '
   'engagement; cross-sectional',
   'lo âu thanh thiếu niên; phân loại DSM-5; mô hình phương trình cấu '
   'trúc; bất biến giới; Việt Nam; bắt nạt học đường; tự trọng; gắn bó '
   'trường học; nghiên cứu cắt ngang')


d.add_page_break()


# ============================================================
H2('1. INTRODUCTION', '1. ĐẶT VẤN ĐỀ')

H3('1.1 Adolescent anxiety burden and DSM-5 subtype rationale',
   '1.1 Gánh nặng lo âu thanh thiếu niên và lý do tiếp cận theo phân loại DSM-5')

PP('Anxiety is the leading mental health concern among adolescents '
   'globally. The Global Burden of Disease 2021 estimates that anxiety '
   'disorders affect approximately 11.9% of adolescents in the ASEAN '
   'region and 10.1% in Vietnam, with the largest age-standardised '
   'incidence increase recorded in Asia between 1990 and 2021 [GBD '
   'ASEAN 2025; Kieling et al. 2024]. Beyond aggregate burden, '
   'epidemiological evidence indicates that anxiety in adolescence is '
   'frequently chronic: a population-based cohort in Australia found '
   'that 64% of adolescents reported anxiety or depressive symptoms on '
   'three or more occasions between ages 10 and 18 [Robson et al. 2025].',
   'Lo âu là vấn đề sức khỏe tâm thần hàng đầu ở thanh thiếu niên trên '
   'toàn cầu. Nghiên cứu Gánh nặng Bệnh tật Toàn cầu 2021 ước tính rối '
   'loạn lo âu ảnh hưởng tới khoảng 11,9% thanh thiếu niên khu vực ASEAN '
   'và 10,1% tại Việt Nam, với mức gia tăng tỷ lệ mới mắc chuẩn hóa '
   'theo tuổi cao nhất ghi nhận tại châu Á giai đoạn 1990–2021. Ngoài '
   'gánh nặng tổng thể, bằng chứng dịch tễ cho thấy lo âu ở tuổi vị '
   'thành niên thường mang tính mạn tính: một nghiên cứu đoàn hệ dựa '
   'trên dân số tại Úc cho thấy 64% thanh thiếu niên báo cáo triệu '
   'chứng lo âu hoặc trầm cảm ít nhất ba lần trong giai đoạn 10–18 tuổi.')

PP('Critical to the present study is the recognition that "anxiety" is '
   'not a unitary phenomenon. The Diagnostic and Statistical Manual of '
   'Mental Disorders, Fifth Edition [American Psychiatric Association '
   '2013] distinguishes generalised anxiety disorder (GAD), separation '
   'anxiety disorder (SAD) and social anxiety disorder (SocAD) as '
   'discrete diagnostic entities; the Revised Children\'s Anxiety and '
   'Depression Scale operationalises these subtypes for adolescent '
   'self-report [Chorpita 2000]. Aggregating these subtypes can obscure '
   'clinically actionable heterogeneity. Nevertheless, most Asian '
   'studies report anxiety only as a composite outcome [Xu et al. 2021; '
   'Chen et al. 2023; Wen et al. 2020; Saikia et al. 2023], leaving '
   'subtype-specific risk-protective architectures unmapped in the '
   'Vietnamese context.',
   'Một điểm cốt lõi của nghiên cứu này là sự thừa nhận rằng "lo âu" '
   'không phải là một hiện tượng đơn nhất. Sổ tay Chẩn đoán và Thống '
   'kê các Rối loạn Tâm thần phiên bản thứ năm (Hiệp hội Tâm thần học '
   'Hoa Kỳ, 2013) phân biệt rối loạn lo âu tổng quát (GAD), rối loạn '
   'lo âu chia ly (SAD) và rối loạn lo âu xã hội (SocAD) thành các '
   'thực thể chẩn đoán riêng biệt; thang Đánh giá Lo âu và Trầm cảm '
   'sửa đổi cho Trẻ em (RCADS) thao tác hóa các phân loại này cho tự '
   'báo cáo ở thanh thiếu niên (Chorpita, 2000). Việc gộp các phân '
   'loại này có thể che lấp tính dị biệt lâm sàng có giá trị hành '
   'động. Tuy vậy, phần lớn các nghiên cứu châu Á chỉ báo cáo lo âu '
   'dưới dạng kết quả tổng hợp, để lại khoảng trống cấu trúc nguy cơ '
   '– bảo vệ theo từng phân loại trong bối cảnh Việt Nam.')

PP('The Vietnamese context introduces three culturally specific '
   'considerations [Anderson et al. 2025]. First, a Confucian '
   'academic culture [Stankov 2010] frames high-stakes examinations '
   'as the principal gateway to social mobility, intensifying chronic '
   'academic stress. Second, a hierarchical family structure '
   'constrains emotional disclosure between adolescents and parents, '
   'potentially attenuating parental support efficacy. Third, a '
   'collectivist orientation may uniformly shape attachment processes '
   'across genders, with theoretical implications for separation '
   'anxiety [Triandis 1995; Markus & Kitayama 1991]. These '
   'culture-specific features call for empirical examination beyond '
   'imported Western frameworks.',
   'Bối cảnh Việt Nam đặt ra ba yếu tố đặc thù văn hóa. Thứ nhất, '
   'văn hóa học thuật Nho giáo (Stankov 2010) coi các kỳ thi áp lực '
   'cao là cửa ngõ chính để di chuyển xã hội, làm gia tăng stress '
   'học tập kéo dài. Thứ hai, cấu trúc gia đình thứ bậc hạn chế '
   'việc bộc lộ cảm xúc giữa thanh thiếu niên và cha mẹ, có thể '
   'làm suy giảm hiệu lực của hỗ trợ cha mẹ. Thứ ba, định hướng '
   'tập thể có thể định hình các quá trình gắn bó một cách đồng '
   'nhất giữa các giới, với hàm ý lý thuyết cho lo âu chia ly. Các '
   'đặc trưng văn hóa này đòi hỏi kiểm chứng thực nghiệm vượt ra '
   'ngoài các khung lý thuyết phương Tây nhập khẩu.')


H3('1.2 Risk-protective framework and the integration imperative',
   '1.2 Khung lý thuyết nguy cơ – bảo vệ và yêu cầu tích hợp')

PP('The transactional stress-coping framework [Lazarus & Folkman 1984] '
   'posits that adolescent psychopathology emerges from dynamic '
   'interactions between stressors and resources, both internal '
   '(self-esteem) and external (school, peer, family). A landmark '
   'meta-analysis by Compas and colleagues (212 studies, N = 80,850) '
   'confirmed the joint contribution of risk and protective factors to '
   'youth internalising symptoms [Compas et al. 2017]. Critically, '
   'however, the majority of primary studies tested factors in '
   'isolation, leaving the interactive architecture under-specified.',
   'Khung lý thuyết stress-ứng phó giao dịch giả định rằng tâm bệnh ở '
   'tuổi thanh thiếu niên xuất hiện từ tương tác động giữa các yếu tố '
   'gây stress và các nguồn lực, cả nội tại (tự trọng) và bên ngoài '
   '(trường học, bạn bè, gia đình). Một phân tích tổng hợp quan trọng '
   'của Compas và cộng sự (212 nghiên cứu, N = 80.850) đã khẳng định '
   'sự đóng góp đồng thời của các yếu tố nguy cơ và bảo vệ đối với các '
   'triệu chứng nội hóa ở giới trẻ. Tuy nhiên, phần lớn các nghiên cứu '
   'gốc chỉ kiểm định các yếu tố một cách riêng lẻ, để lại cấu trúc '
   'tương tác chưa được xác định.')

PP('Within Vietnam, methodological inconsistency further obscures the '
   'evidence base. The Vietnam Adolescent Mental Health Survey '
   '(V-NAMHS) using the DSM-5–anchored Diagnostic Interview Schedule '
   'for Children (DISC-5) estimated a 12-month anxiety disorder '
   'prevalence of 2.3% [V-NAMHS 2022], whereas school-based studies '
   'using continuous symptom scales (e.g., DASS-21) report substantially '
   'higher rates [Hoang Trung Hoc 2025, N = 8,389]. This measurement '
   'gap underscores the need for analytical strategies that move beyond '
   'aggregate prevalence to mechanistic models of subtype-specific '
   'pathways.',
   'Tại Việt Nam, sự thiếu nhất quán về phương pháp luận làm cho cơ sở '
   'bằng chứng thêm mờ nhạt. Khảo sát Sức khỏe Tâm thần Thanh thiếu '
   'niên Việt Nam (V-NAMHS) sử dụng Bảng phỏng vấn chẩn đoán cho trẻ '
   'em phiên bản DSM-5 (DISC-5) ước tính tỷ lệ rối loạn lo âu trong 12 '
   'tháng là 2,3%, trong khi các nghiên cứu trường học sử dụng thang '
   'đo triệu chứng liên tục (ví dụ DASS-21) báo cáo tỷ lệ cao hơn nhiều '
   '(Hoàng Trung Học 2025, N = 8.389). Khoảng cách đo lường này nhấn '
   'mạnh nhu cầu về chiến lược phân tích vượt ra ngoài tỷ lệ tổng hợp '
   'để tiến tới mô hình cơ chế của các đường dẫn theo từng phân loại.')


H3('1.3 Gender, separation anxiety and the cultural collectivism hypothesis',
   '1.3 Giới, lo âu chia ly và giả thuyết văn hóa tập thể')

PP('A standard expectation derived from Western literature is that '
   'females exhibit higher anxiety than males across subtypes '
   '[McLean et al. 2011]. Emerging Asian evidence challenges this '
   'expectation: rural Chinese males have shown higher anxiety than '
   'rural Chinese females [Wen et al. 2020], and Northeast Indian '
   'adolescent boys reported significantly higher severe-anxiety '
   'prevalence than girls (30.0% vs 18.9%; p = 0.049) [Saikia et al. '
   '2023]. These patterns suggest that gender effects on anxiety are '
   'culturally and contextually contingent rather than universal.',
   'Một kỳ vọng tiêu chuẩn rút ra từ y văn phương Tây là nữ có mức lo '
   'âu cao hơn nam ở các phân loại. Bằng chứng mới nổi từ châu Á '
   'thách thức kỳ vọng này: nam thanh thiếu niên Trung Quốc nông thôn '
   'có mức lo âu cao hơn nữ cùng vùng, và nam thanh thiếu niên Đông '
   'Bắc Ấn Độ có tỷ lệ lo âu nặng cao hơn nữ một cách có ý nghĩa '
   '(30,0% so với 18,9%; p = 0,049). Các mẫu hình này gợi ý rằng tác '
   'động của giới lên lo âu phụ thuộc vào văn hóa và bối cảnh, không '
   'phải hằng định phổ quát.')

PP('We advance a culturally specific hypothesis: in Vietnam, '
   'collectivist family hierarchy may homogenise separation-anxiety '
   'experiences across genders while gender-differentiated social and '
   'academic pressures remain operative for GAD and SocAD. This '
   'prediction draws on interdependent self-construal theory [Markus '
   '& Kitayama 1991] and on developmental work indicating that '
   'separation-individuation tasks are predominantly childhood-onset '
   'and thus precede the post-pubertal differentiation of '
   'gender-loaded social pressures [Blos 1979; Kroger 2007].',
   'Chúng tôi đưa ra một giả thuyết đặc thù văn hóa: tại Việt Nam, hệ '
   'thống thứ bậc gia đình tập thể có thể đồng nhất hóa trải nghiệm '
   'lo âu chia ly giữa các giới, trong khi áp lực xã hội và học tập '
   'phân hóa theo giới vẫn tác động lên GAD và SocAD. Dự đoán này dựa '
   'trên lý thuyết bản thân phụ thuộc lẫn nhau và các nghiên cứu phát '
   'triển cho thấy các nhiệm vụ chia ly – cá nhân hóa chủ yếu khởi '
   'phát từ thời thơ ấu, do đó đi trước sự phân hóa giới của các áp '
   'lực xã hội sau dậy thì.')


H3('1.4 The present study', '1.4 Nghiên cứu hiện tại')

PP('We present a cross-sectional SEM study that employs '
   'subtype-specific structural equation modelling to test three '
   'hypotheses in a Vietnamese lower secondary school sample:',
   'Chúng tôi trình bày một nghiên cứu cắt ngang sử dụng SEM theo từng '
   'phân loại nhằm kiểm định ba giả thuyết trên mẫu học sinh trung học '
   'cơ sở Việt Nam:')

PP('H1: All three risk factors (academic pressure, smartphone '
   'addiction, school bullying) will show significant positive paths '
   'to GAD, SAD and SocAD.',
   'H1: Cả ba yếu tố nguy cơ (áp lực học tập, nghiện điện thoại, bắt '
   'nạt học đường) đều có đường dẫn dương có ý nghĩa tới GAD, SAD và '
   'SocAD.', indent=False)

PP('H2: All four protective factors (school engagement, parental '
   'support, peer support, self-esteem) will show significant negative '
   'paths to GAD, SAD and SocAD.',
   'H2: Cả bốn yếu tố bảo vệ (gắn bó trường học, hỗ trợ cha mẹ, hỗ '
   'trợ bạn bè, tự trọng) đều có đường dẫn âm có ý nghĩa tới GAD, SAD '
   'và SocAD.', indent=False)

PP('H3: Multi-group SEM by gender will reveal gender invariance for '
   'separation anxiety only, while GAD and SocAD will exhibit gender '
   'differences favouring higher scores in females.',
   'H3: SEM đa nhóm theo giới sẽ làm rõ tính bất biến giới chỉ với lo '
   'âu chia ly, trong khi GAD và SocAD sẽ thể hiện khác biệt giới với '
   'điểm số cao hơn ở nữ.', indent=False)

NOTE('Note: This is a quantitative cross-sectional analysis. A '
     'companion qualitative study is planned separately.',
     'Ghi chú: Đây là phân tích định lượng cắt ngang. Một nghiên cứu '
     'định tính bổ trợ được lên kế hoạch triển khai riêng.')


d.add_page_break()


# ============================================================
H2('2. METHODS', '2. PHƯƠNG PHÁP')

H3('2.1 Study design and participants', '2.1 Thiết kế nghiên cứu và người tham gia')

PP('We employed a cross-sectional survey design with structural '
   'equation modelling. The sample comprised 1,352 lower secondary '
   'school students (Male = 614, 45.4%; Female = 738, 54.6%) recruited '
   'from two purposively selected schools in Hanoi, Vietnam (Nhat Tan '
   'and Tay Mo) representing urban and suburban contexts. Participants '
   'were aged 11–14 years and distributed across grades 6 to 9 '
   '(Grade 6 = 368, Grade 7 = 316, Grade 8 = 340, Grade 9 = 328). '
   'Recruitment proceeded via school administration, with written '
   'parental consent and student written assent.',
   'Chúng tôi sử dụng thiết kế khảo sát cắt ngang kết hợp với mô hình '
   'phương trình cấu trúc. Mẫu gồm 1.352 học sinh trung học cơ sở '
   '(nam = 614; 45,4% và nữ = 738; 54,6%) được tuyển từ hai trường '
   'tại Hà Nội được chọn có chủ đích (Nhật Tân và Tây Mỗ) đại diện cho '
   'bối cảnh đô thị và ngoại ô. Người tham gia ở độ tuổi 11–14, phân '
   'bố qua các khối từ 6 đến 9 (khối 6 = 368, khối 7 = 316, khối 8 = '
   '340, khối 9 = 328). Quá trình tuyển chọn được thực hiện qua ban '
   'giám hiệu, có chấp thuận bằng văn bản của cha mẹ và đồng ý bằng '
   'văn bản của học sinh.')

H3('2.2 Measures', '2.2 Công cụ đo lường')

PP('Eight validated instruments were adapted to Vietnamese via forward '
   '– back translation, expert review (three clinical psychologists '
   'and one educational measurement specialist) and pilot testing on a '
   'separate sample of 50 students prior to the main survey. All raw '
   'scale scores were rescaled to a 0 – 100 metric for cross-scale '
   'comparability while preserving original scoring conventions.',
   'Tám công cụ đo lường đã được kiểm định được chuyển ngữ sang tiếng '
   'Việt theo quy trình dịch xuôi – dịch ngược, được hội đồng chuyên '
   'gia (ba nhà tâm lý lâm sàng và một chuyên gia đo lường giáo dục) '
   'rà soát và được khảo sát thử trên một mẫu riêng gồm 50 học sinh '
   'trước khi triển khai chính thức. Tất cả điểm thô được tái cân '
   'chỉnh về thang 0 – 100 để so sánh được giữa các thang đo, trong '
   'khi vẫn giữ nguyên quy ước chấm điểm gốc.')

PP('Anxiety subtypes were measured using the Revised Children\'s '
   'Anxiety and Depression Scale (RCADS; Chorpita 2000) — a '
   'developmentally calibrated, DSM-5-aligned instrument. We extracted '
   'three subscales: generalised anxiety disorder (7 items), '
   'separation anxiety disorder (4 items) and social anxiety disorder '
   '(4 items), each rated on a 4-point Likert scale.',
   'Các phân loại lo âu được đo bằng thang Đánh giá Lo âu và Trầm cảm '
   'sửa đổi cho Trẻ em (RCADS) — một công cụ phù hợp với phát triển '
   'và liên kết với DSM-5. Chúng tôi trích xuất ba tiểu thang: rối '
   'loạn lo âu tổng quát (7 mục), rối loạn lo âu chia ly (4 mục) và '
   'rối loạn lo âu xã hội (4 mục), mỗi mục chấm trên thang Likert 4 '
   'điểm.')

PP('Risk factors were assessed using: (a) Educational Stress Scale '
   'for Adolescents — ESSA (Sun et al. 2011), 4 items; (b) Smartphone '
   'Addiction Scale-Short Version — SAS-SV (Kwon et al. 2013), 5 '
   'items; and (c) Olweus Bully/Victim Questionnaire — OBVQ (Olweus '
   '1996), 8 items capturing physical and verbal victimisation.',
   'Các yếu tố nguy cơ được đo bằng: (a) thang Áp lực Học tập cho Vị '
   'thành niên — ESSA (Sun và cộng sự, 2011), 4 mục; (b) thang Nghiện '
   'Điện thoại thông minh phiên bản rút gọn — SAS-SV (Kwon và cộng '
   'sự, 2013), 5 mục; và (c) Bảng hỏi Bắt nạt/Nạn nhân Olweus — OBVQ '
   '(Olweus, 1996), 8 mục đo lường bắt nạt thể chất và lời nói.')

PP('Protective factors comprised: (a) Psychological Sense of School '
   'Membership — PSSM (Goodenow 1993), 7 items; (b) Multidimensional '
   'Scale of Perceived Social Support — MSPSS (Zimet et al. 1988), '
   '8 items split into parental and peer subscales (4 items each); '
   'and (c) Rosenberg Self-Esteem Scale — RSES (Rosenberg 1965), 5 '
   'items.',
   'Các yếu tố bảo vệ gồm: (a) thang Cảm nhận Gắn bó với Trường — '
   'PSSM (Goodenow, 1993), 7 mục; (b) thang Đa chiều về Hỗ trợ Xã hội '
   'Cảm nhận — MSPSS (Zimet và cộng sự, 1988), 8 mục chia thành tiểu '
   'thang cha mẹ và bạn bè (mỗi tiểu thang 4 mục); và (c) thang Tự '
   'trọng Rosenberg — RSES (Rosenberg, 1965), 5 mục.')


H3('2.3 Analytic strategy', '2.3 Chiến lược phân tích')

PP('Quantitative analyses proceeded in five steps using AMOS 31.0 and '
   'SPSS 31.0. Step 1 generated descriptive statistics and reliability '
   'estimates (Cronbach\'s α, McDonald\'s ω). Step 2 performed '
   'confirmatory factor analysis (CFA) for each scale, applying the '
   'fit-index thresholds CFI ≥ 0.90, TLI ≥ 0.90, RMSEA ≤ 0.08 and '
   'SRMR ≤ 0.08 [Hu & Bentler 1999]. Step 3 estimated '
   'subtype-specific structural models for each of the seven '
   'predictors against each of the three anxiety subtypes plus the '
   'three-factor total RLLA outcome (21 path coefficients of '
   'theoretical interest plus seven aggregate paths). Step 4 fitted '
   'an integrated risk-protective composite model in which the '
   'three risk factors loaded onto a higher-order Risk latent (YTNC) '
   'and the four protective factors loaded onto a higher-order '
   'Protective latent (YTBV); both higher-order latents simultaneously '
   'predicted total RLLA. Step 5 implemented multi-group invariance '
   'testing (configural → metric → scalar) by gender, applying ΔCFI ≤ '
   '0.01 as the invariance criterion [Cheung & Rensvold 2002].',
   'Phân tích định lượng tiến hành theo năm bước trên AMOS 31.0 và '
   'SPSS 31.0. Bước 1 tạo thống kê mô tả và ước lượng độ tin cậy '
   '(Cronbach α, McDonald ω). Bước 2 chạy phân tích nhân tố khẳng '
   'định (CFA) cho từng thang, áp dụng ngưỡng chỉ số phù hợp CFI ≥ '
   '0,90; TLI ≥ 0,90; RMSEA ≤ 0,08; SRMR ≤ 0,08. Bước 3 ước lượng '
   'mô hình cấu trúc theo từng phân loại cho mỗi yếu tố trong số '
   'bảy yếu tố dự báo với từng phân loại trong số ba phân loại lo âu '
   'cộng với kết quả tổng RLLA ba nhân tố (21 hệ số đường dẫn quan '
   'tâm lý thuyết cộng với bảy đường dẫn tổng hợp). Bước 4 ước lượng '
   'mô hình tích hợp nguy cơ – bảo vệ tổng hợp, trong đó ba yếu '
   'tố nguy cơ tải lên một biến tiềm ẩn bậc cao Nguy cơ (YTNC) và '
   'bốn yếu tố bảo vệ tải lên một biến tiềm ẩn bậc cao Bảo vệ (YTBV); '
   'cả hai biến tiềm ẩn bậc cao đồng thời dự báo tổng RLLA. Bước 5 '
   'triển khai kiểm định bất biến đa nhóm (cấu hình → metric → '
   'thang điểm) theo giới, áp dụng ΔCFI ≤ 0,01 làm tiêu chí bất biến.')


H3('2.4 Ethics', '2.4 Đạo đức nghiên cứu')

PP('The study was approved by the Hanoi National University of '
   'Education Institutional Review Board. Written parental consent '
   'and student written assent were obtained. Data were anonymised '
   'via ID codes; no personal identifiers appear in analysis files; '
   'data were stored on access-controlled institutional servers.',
   'Nghiên cứu được Hội đồng Đạo đức của Trường Đại học Sư phạm Hà '
   'Nội phê duyệt. Đồng ý bằng văn bản của cha mẹ và đồng ý bằng văn '
   'bản của học sinh đều đã được thu thập. Dữ liệu được ẩn danh hóa '
   'bằng mã số nhận dạng; không có thông tin cá nhân xuất hiện trong '
   'tệp phân tích; dữ liệu được lưu trữ trên máy chủ truy cập có '
   'kiểm soát của cơ sở.')

TBD('Q3-6 BLOCKING: IRB decision number and date pending official '
    'letter from HNUE Ethics Committee',
    'Q3-6 CHỜ: số quyết định và ngày của Hội đồng Đạo đức ĐHSPHN')


d.add_page_break()


# ============================================================
H2('3. RESULTS', '3. KẾT QUẢ')

H3('3.1 Sample characteristics and psychometric properties',
   '3.1 Đặc điểm mẫu và tính chất tâm trắc')

PP('Of the 1,352 participants, 614 (45.4%) were male and 738 (54.6%) '
   'female. Mean age was 12.5 years (SD = 1.1). Grade distribution: '
   'Grade 6 (n = 368), Grade 7 (n = 316), Grade 8 (n = 340), Grade 9 '
   '(n = 328). All eight scales achieved acceptable internal '
   'consistency (Cronbach\'s α ≥ 0.70) and adequate CFA fit (CFI ≥ '
   '0.90, RMSEA ≤ 0.08). Detailed psychometric statistics appear in '
   'Table 2.',
   'Trong số 1.352 người tham gia, 614 (45,4%) là nam và 738 (54,6%) '
   'là nữ. Độ tuổi trung bình là 12,5 (SD = 1,1). Phân bố khối lớp: '
   'khối 6 (n = 368), khối 7 (n = 316), khối 8 (n = 340), khối 9 (n '
   '= 328). Tất cả tám thang đo đạt độ nhất quán nội tại chấp nhận '
   'được (Cronbach α ≥ 0,70) và độ phù hợp CFA đầy đủ (CFI ≥ 0,90; '
   'RMSEA ≤ 0,08). Các chỉ số tâm trắc chi tiết được trình bày trong '
   'Bảng 2.')


H3('3.2 Descriptive anxiety levels', '3.2 Mức độ lo âu mô tả')

PP('On the 0–100 metric, mean anxiety scores differed substantially '
   'across subtypes: GAD M = 55.82 (SD ≈ 22.4), SocAD M = 48.41 '
   '(SD = 26.19) and SAD M = 25.06 (SD = 24.29). Separation anxiety '
   'scored markedly lower than both generalised and social anxiety, '
   'in line with developmental theory positing childhood onset and '
   'attenuation across early adolescence.',
   'Trên thang 0–100, điểm lo âu trung bình khác biệt đáng kể giữa '
   'các phân loại: GAD M = 55,82 (SD ≈ 22,4); SocAD M = 48,41 '
   '(SD = 26,19) và SAD M = 25,06 (SD = 24,29). Lo âu chia ly có '
   'điểm thấp hơn rõ rệt so với cả lo âu tổng quát và lo âu xã hội, '
   'phù hợp với lý thuyết phát triển cho rằng SAD khởi phát từ thời '
   'thơ ấu và suy giảm dần khi vào tuổi thanh thiếu niên sớm.')


H3('3.3 Subtype-specific path coefficients (Table 3)',
   '3.3 Hệ số đường dẫn theo từng phân loại (Bảng 3)')

PP('Subtype-specific SEM was run for each of the seven predictors '
   'against GAD, SAD, SocAD and total RLLA. All 21 risk-anxiety paths '
   'and 12 of 12 protective-anxiety paths to GAD/SocAD reached '
   'statistical significance, with the conspicuous exceptions of '
   'peer support paths to GAD and SAD (n.s.) and school engagement → '
   'SAD (n.s.). Detailed coefficients appear below.',
   'SEM theo từng phân loại được chạy cho mỗi yếu tố trong số bảy yếu '
   'tố dự báo với GAD, SAD, SocAD và RLLA tổng. Toàn bộ 21 đường dẫn '
   'nguy cơ – lo âu và 12 trên 12 đường dẫn bảo vệ – lo âu tới '
   'GAD/SocAD đạt ý nghĩa thống kê, với ngoại lệ rõ ràng là các đường '
   'dẫn hỗ trợ bạn bè tới GAD và SAD (không có ý nghĩa) và gắn bó '
   'trường học tới SAD (không có ý nghĩa). Hệ số chi tiết được trình '
   'bày bên dưới.')


# Table 3: Subtype-specific beta table
P_header = ['Predictor', '→ GAD (β)', '→ SAD (β)', '→ SocAD (β)',
            '→ Total RLLA (β; R²)']

SIG3 = '***'  # statistical p < 0.001
SIG2 = '**'   # statistical p < 0.01
SIG1 = '*'    # statistical p < 0.05

table_rows = [
    ('RISK FACTORS / YẾU TỐ NGUY CƠ', '', '', '', ''),
    ('Academic pressure (ALHT) / Áp lực học tập',
     f'+0,510{SIG3}', f'+0,253{SIG3}', f'+0,490{SIG3}',
     f'+0,533{SIG3} (R²=0,284)'),
    ('Smartphone addiction (NĐT) / Nghiện điện thoại',
     f'+0,336{SIG3}', f'+0,265{SIG3}', f'+0,383{SIG3}',
     f'+0,400{SIG3} (R²=0,160)'),
    ('School bullying (BNHĐ) / Bắt nạt học đường',
     f'+0,215{SIG3}', f'+0,376{SIG3} ★', f'+0,253{SIG3}',
     f'+0,276{SIG3} (R²=0,076)'),
    ('PROTECTIVE FACTORS / YẾU TỐ BẢO VỆ', '', '', '', ''),
    ('School engagement (GBTH) / Gắn bó trường học',
     f'−0,108{SIG2}', '+0,014 (n.s.)', f'−0,187{SIG3}',
     f'−0,155{SIG3} (R²=0,024)'),
    ('Self-esteem (TTr) / Tự trọng',
     f'−0,455{SIG3} ★', f'−0,087{SIG1}', f'−0,415{SIG3}',
     f'−0,457{SIG3} (R²=0,209)'),
    ('Parental support (HTCM) / Hỗ trợ cha mẹ',
     f'−0,172{SIG3}', '0,000 (n.s.)', f'−0,273{SIG3}',
     f'−0,234{SIG3} (R²=0,055)'),
    ('Peer support (HTBB) / Hỗ trợ bạn bè',
     '−0,015 (n.s.)', '−0,019 (n.s.)', f'−0,079{SIG1}',
     '−0,044 (n.s.) (R²=0,002)'),
]

t = d.add_table(rows=1, cols=5); t.style = 'Light Grid Accent 1'; t.autofit = False
hdr = t.rows[0].cells
for i, h in enumerate(P_header):
    hdr[i].text = h
    for p in hdr[i].paragraphs:
        for r in p.runs:
            r.font.bold = True; r.font.size = Pt(10)
for row_data in table_rows:
    row = t.add_row().cells
    for i, txt in enumerate(row_data):
        row[i].text = str(txt)
        for p in row[i].paragraphs:
            for r in p.runs:
                r.font.size = Pt(10)
set_col_widths(t, [5.5, 2.4, 2.4, 2.4, 4.0])

PP(f'Note. {SIG3} p < 0.001; {SIG2} p < 0.01; {SIG1} p < 0.05; n.s. = '
   'non-significant. ★ = key novel pathways. β verified against LA '
   'Bảng 27 (ALHT), 30 (NĐT), 33 (BNHĐ), 36 (GBTH), 39 (TTr), 42 '
   '(HTCM), 45 (HTBB).',
   f'Ghi chú. {SIG3} p < 0,001; {SIG2} p < 0,01; {SIG1} p < 0,05; '
   'n.s. = không có ý nghĩa thống kê. ★ = các đường dẫn then chốt '
   'mới. Các giá trị β đã được đối chiếu với LA Bảng 27 (ALHT), 30 '
   '(NĐT), 33 (BNHĐ), 36 (GBTH), 39 (TTr), 42 (HTCM), 45 (HTBB).',
   indent=False)


H3('3.4 Three differential-pathway findings',
   '3.4 Ba phát hiện đường dẫn khác biệt')

PP('First, school bullying showed its strongest path to separation '
   'anxiety (β = 0.376, p < 0.001), exceeding its paths to GAD '
   '(β = 0.215) and SocAD (β = 0.253). This pattern is counterintuitive '
   'because bullying is most often theorised as a social-evaluative '
   'stressor and might therefore be expected to associate most '
   'strongly with SocAD. The selective amplification of separation '
   'anxiety suggests that, for Vietnamese lower-secondary adolescents, '
   'school bullying activates attachment insecurity rather than '
   'social-evaluative concerns per se.',
   'Thứ nhất, bắt nạt học đường có đường dẫn mạnh nhất tới lo âu '
   'chia ly (β = 0,376; p < 0,001), vượt qua các đường dẫn tới GAD '
   '(β = 0,215) và SocAD (β = 0,253). Mẫu hình này phản trực giác vì '
   'bắt nạt thường được lý thuyết hóa là stress đánh giá xã hội nên '
   'lẽ ra liên hệ mạnh nhất với SocAD. Sự khuếch đại chọn lọc của lo '
   'âu chia ly gợi ý rằng, với học sinh trung học cơ sở Việt Nam, bắt '
   'nạt học đường kích hoạt mất an toàn gắn bó hơn là các mối lo đánh '
   'giá xã hội thuần túy.')

PP('Second, self-esteem emerged as the strongest single protective '
   'lever for GAD (β = −0.455, p < 0.001), with magnitude approaching '
   'the dominant academic-pressure risk effect on GAD (β = 0.510). '
   'The total RLLA path from self-esteem was −0.457 (R² = 0.209). '
   'This positions self-esteem building as a high-yield intervention '
   'target.',
   'Thứ hai, tự trọng nổi lên là đòn bẩy bảo vệ đơn lẻ mạnh nhất '
   'đối với GAD (β = −0,455; p < 0,001), với độ lớn gần bằng tác động '
   'chiếm ưu thế của áp lực học tập lên GAD (β = 0,510). Đường dẫn '
   'tổng từ tự trọng tới RLLA là −0,457 (R² = 0,209). Điều này định '
   'vị xây dựng tự trọng như một mục tiêu can thiệp có hiệu suất cao.')

PP('Third, peer support showed no significant effect on total '
   'RLLA (β = −0.044, p = 0.183, R² = 0.002), and was non-significant '
   'for both GAD (β = −0.015, n.s.) and SAD (β = −0.019, n.s.), '
   'reaching only marginal significance for SocAD (β = −0.079, p = '
   '0.020). This pattern challenges the assumed universal protective '
   'role of peer relationships and is consistent with the '
   'co-rumination hypothesis: among adolescents who do confide, '
   'frequency of peer disclosure may not equate to quality of support.',
   'Thứ ba, hỗ trợ bạn bè không có tác động có ý nghĩa lên tổng '
   'RLLA (β = −0,044; p = 0,183; R² = 0,002), không có ý nghĩa với cả '
   'GAD (β = −0,015) và SAD (β = −0,019), chỉ đạt ý nghĩa cận biên '
   'với SocAD (β = −0,079; p = 0,020). Mẫu hình này thách thức vai '
   'trò bảo vệ phổ quát được giả định của quan hệ bạn bè và phù hợp '
   'với giả thuyết co-rumination: ở thanh thiếu niên có chia sẻ, tần '
   'suất giãi bày với bạn bè không nhất thiết đồng nghĩa với chất '
   'lượng hỗ trợ.')


H3('3.5 Integrated risk-protective composite model',
   '3.5 Mô hình tích hợp nguy cơ – bảo vệ tổng hợp')

PP('A higher-order composite SEM combined the three risk factors into '
   'a Risk latent (YTNC) and the four protective factors into a '
   'Protective latent (YTBV), both predicting total RLLA simultaneously. '
   'The composite Risk → RLLA path was β = 0.669 (p < 0.001) and the '
   'composite Protective → RLLA path was β = −0.220 (p < 0.001), with '
   'integrated R² = 0.598 for total RLLA. The risk-side composite '
   'thus accounted for approximately three times the standardised '
   'effect of the protective-side composite, indicating that '
   'risk-factor reduction may be the higher-leverage intervention '
   'target at the population level in this Vietnamese sample.',
   'Một SEM tổng hợp bậc cao kết hợp ba yếu tố nguy cơ thành một biến '
   'tiềm ẩn Nguy cơ (YTNC) và bốn yếu tố bảo vệ thành một biến tiềm '
   'ẩn Bảo vệ (YTBV), cả hai cùng dự báo tổng RLLA đồng thời. Đường '
   'dẫn tổng hợp Nguy cơ → RLLA là β = 0,669 (p < 0,001) và đường dẫn '
   'tổng hợp Bảo vệ → RLLA là β = −0,220 (p < 0,001), với R² tích '
   'hợp = 0,598 cho tổng RLLA. Như vậy, mặt nguy cơ chiếm khoảng '
   'gấp ba lần ảnh hưởng chuẩn hóa của mặt bảo vệ, cho thấy giảm yếu '
   'tố nguy cơ có thể là mục tiêu can thiệp có đòn bẩy cao hơn ở cấp '
   'độ quần thể trong mẫu Việt Nam này.')


H3('3.6 Multi-group invariance by gender — H3 supported',
   '3.6 Bất biến đa nhóm theo giới — H3 được ủng hộ')

PP('Configural invariance was achieved (both groups fitted the same '
   'structural model), supporting subsequent invariance comparisons, '
   'following procedures established for cross-gender invariance of '
   'GAD-7 in adolescent samples [Niwenahisemo, Hong, & Kuang 2024]. '
   'Mean scores diverged significantly by gender for GAD '
   '(Female M = 59.47 vs Male M = 51.43; F = 44.48, p < 0.001) and '
   'SocAD (Female M = 52.74 vs Male M = 43.20; F = 45.98, '
   'p < 0.001). Strikingly, SAD exhibited complete gender '
   'invariance (Female M = 24.76 vs Male M = 25.42; F = 0.246, '
   'p = 0.620), consistent with H3. This is the first documented '
   'instance of complete SAD gender invariance in a Vietnamese '
   'lower-secondary sample.',
   'Bất biến cấu hình được đạt (cả hai nhóm vừa khít cùng một mô hình '
   'cấu trúc), hỗ trợ các so sánh bất biến tiếp theo, theo các quy trình '
   'đã được thiết lập cho bất biến giữa các giới của GAD-7 trên mẫu '
   'thanh thiếu niên (Niwenahisemo, Hong, & Kuang 2024). Điểm trung bình '
   'khác biệt có ý nghĩa theo giới với GAD (nữ M = 59,47 so với '
   'nam M = 51,43; F = 44,48; p < 0,001) và SocAD (nữ M = 52,74 '
   'so với nam M = 43,20; F = 45,98; p < 0,001). Đáng chú ý, SAD '
   'thể hiện bất biến giới hoàn toàn (nữ M = 24,76 so với nam M = '
   '25,42; F = 0,246; p = 0,620), phù hợp với H3. Đây là trường hợp '
   'được ghi nhận đầu tiên về bất biến giới hoàn toàn của SAD trên '
   'mẫu học sinh trung học cơ sở Việt Nam.')


d.add_page_break()


# ============================================================
H2('4. DISCUSSION', '4. THẢO LUẬN')

H3('4.1 Summary of findings', '4.1 Tổng hợp các phát hiện')

PP('In a large Vietnamese lower-secondary sample, integrated SEM '
   'documented three differential-pathway findings: bullying showed '
   'its strongest impact on separation anxiety, self-esteem was the '
   'strongest single protective lever for generalised anxiety, and '
   'peer support exerted no significant protective effect on any '
   'subtype. The composite risk-protective model accounted for 59.8% '
   'of the variance in total anxiety, with risk-side contribution '
   'approximately three times the protective-side contribution. '
   'Multi-group invariance testing revealed gender differences for '
   'GAD and SocAD but complete gender invariance for SAD — a finding '
   'we interpret through a cultural-collectivism lens. H1 was '
   'supported in full; H2 was supported with the notable exception '
   'of peer-support paths; H3 was supported.',
   'Trong một mẫu lớn học sinh trung học cơ sở Việt Nam, SEM tích hợp '
   'ghi nhận ba phát hiện đường dẫn khác biệt: bắt nạt có tác động '
   'mạnh nhất lên lo âu chia ly, tự trọng là đòn bẩy bảo vệ đơn lẻ '
   'mạnh nhất đối với lo âu tổng quát, và hỗ trợ bạn bè không có hiệu '
   'lực bảo vệ có ý nghĩa với bất kỳ phân loại nào. Mô hình tổng hợp '
   'nguy cơ – bảo vệ giải thích 59,8% phương sai của tổng lo âu, với '
   'đóng góp của mặt nguy cơ khoảng gấp ba lần đóng góp của mặt bảo '
   'vệ. Kiểm định bất biến đa nhóm cho thấy khác biệt giới với GAD '
   'và SocAD nhưng bất biến giới hoàn toàn với SAD — một phát hiện '
   'chúng tôi diễn giải qua lăng kính văn hóa tập thể. H1 được ủng '
   'hộ đầy đủ; H2 được ủng hộ với ngoại lệ đáng chú ý của các đường '
   'dẫn từ hỗ trợ bạn bè; H3 được ủng hộ.')


H3('4.2 Bullying → SAD as the strongest risk pathway',
   '4.2 Bắt nạt → lo âu chia ly: đường dẫn nguy cơ mạnh nhất')

PP('The dominance of bullying for separation anxiety (β = 0.376) '
   'rather than for the more theoretically anticipated social-anxiety '
   'subtype (β = 0.253) is, to our knowledge, the first such '
   'observation in a Vietnamese lower-secondary sample. We interpret '
   'this pattern via attachment theory: bullying victimisation '
   'plausibly disrupts the perception of school as a secondary secure '
   'base, intensifying attachment-related distress on separation from '
   'primary caregivers [Bowlby 1988]. The pattern aligns with '
   'cross-cultural evidence that bullying victimisation in '
   'collectivist contexts is preferentially processed as an '
   'attachment-related threat [Triandis 1995].',
   'Tính chiếm ưu thế của bắt nạt với lo âu chia ly (β = 0,376) so với '
   'phân loại lo âu xã hội được kỳ vọng về mặt lý thuyết (β = 0,253), '
   'theo hiểu biết của chúng tôi, là quan sát đầu tiên kiểu này trên '
   'mẫu học sinh trung học cơ sở Việt Nam. Chúng tôi diễn giải mẫu '
   'hình này qua lý thuyết gắn bó (Bowlby 1988): nạn nhân bắt nạt có '
   'khả năng làm gián đoạn nhận thức về trường học như một căn cứ an '
   'toàn thứ cấp, làm tăng cường đau khổ liên quan đến gắn bó khi '
   'chia ly với người chăm sóc chính. Mẫu hình này phù hợp với bằng '
   'chứng liên văn hóa rằng nạn nhân bắt nạt trong bối cảnh tập thể '
   'được xử lý ưu tiên như mối đe dọa liên quan đến gắn bó '
   '(Triandis 1995).')


H3('4.3 Self-esteem as the strongest protective lever',
   '4.3 Tự trọng — đòn bẩy bảo vệ đơn lẻ mạnh nhất')

PP('Self-esteem\'s path to GAD (β = −0.455) approached the magnitude '
   'of academic pressure\'s risk path (β = 0.510), and its path to '
   'total RLLA reached β = −0.457 (R² = 0.209). These findings align '
   'with the Compas et al. (2017) meta-analysis (212 studies, '
   'N = 80,850), which identified secondary control coping — including '
   'cognitive reappraisal — as the most consistent predictor of lower '
   'internalising symptoms in childhood and adolescence. Systematic '
   'reviews of academic pressure and adolescent mental health '
   '[Pascoe et al. 2020; Steare et al. 2023] further support '
   'academic-stress management as a complementary intervention '
   'target. Together these data identify self-esteem building as a '
   'high-yield single-component intervention focus for adolescent '
   'GAD in the Vietnamese context.',
   'Đường dẫn của tự trọng tới GAD (β = −0,455) gần bằng độ lớn của '
   'đường dẫn nguy cơ từ áp lực học tập (β = 0,510), và đường dẫn của '
   'tự trọng tới tổng RLLA đạt β = −0,457 (R² = 0,209). Các phát hiện '
   'này phù hợp với phân tích tổng hợp của Compas và cộng sự (2017) '
   '(212 nghiên cứu, N = 80.850), trong đó nhận diện chiến lược ứng '
   'phó kiểm soát thứ cấp — bao gồm tái diễn giải nhận thức — là '
   'biến dự báo nhất quán nhất cho mức triệu chứng nội hóa thấp ở '
   'trẻ em và thanh thiếu niên. Các tổng quan hệ thống về áp lực học '
   'tập và sức khỏe tâm thần thanh thiếu niên (Pascoe và cộng sự '
   '2020; Steare và cộng sự 2023) tiếp tục ủng hộ quản lý stress học '
   'tập như một mục tiêu can thiệp bổ trợ. Tổng hợp lại, các dữ liệu '
   'này xác định xây dựng tự trọng là một mục tiêu can thiệp đơn '
   'thành phần có hiệu suất cao cho GAD ở thanh thiếu niên trong bối '
   'cảnh Việt Nam.')


H3('4.4 Peer support null effects: the co-rumination hypothesis',
   '4.4 Hỗ trợ bạn bè không có tác động: giả thuyết co-rumination')

PP('Western literature consistently identifies peer support as '
   'protective against adolescent internalising symptoms. Our null '
   'and marginally significant results across all three subtypes '
   'depart from this expectation. We propose two non-mutually-'
   'exclusive interpretations: first, in a Confucian-influenced '
   'collectivist context shaped by the tam giao tradition (the '
   'coexistence of Confucianism, Buddhism and Taoism), peer '
   'disclosure norms may emphasise emotional restraint over '
   'expressive sharing [Small & Blanc 2021], attenuating the '
   'support-seeking buffer mechanism; second, the co-rumination '
   'hypothesis — defined as extensively discussing problems, '
   'speculating about them and focusing on negative feelings '
   '[Rose 2002] — suggests that frequent peer disclosure can '
   'recursively amplify worry rather than dissipate it, especially '
   'when academic-competition norms are operative. These theoretical '
   'interpretations are consistent with cross-cultural evidence on '
   'emotion socialisation in East Asian adolescent samples.',
   'Y văn phương Tây thường xuyên xác định hỗ trợ bạn bè là yếu '
   'tố bảo vệ chống lại các triệu chứng nội hóa ở thanh thiếu '
   'niên. Các kết quả không có ý nghĩa và cận biên của chúng tôi '
   'xuyên suốt cả ba phân loại lệch khỏi kỳ vọng này. Chúng tôi '
   'đề xuất hai cách diễn giải không loại trừ lẫn nhau: thứ nhất, '
   'trong bối cảnh tập thể ảnh hưởng Nho giáo được định hình bởi '
   'truyền thống tam giáo (sự cộng tồn của Nho giáo, Phật giáo và '
   'Đạo giáo), chuẩn mực giãi bày với bạn bè có thể nhấn mạnh sự '
   'chừng mực cảm xúc hơn là chia sẻ biểu cảm (Small & Blanc '
   '2021), làm suy giảm cơ chế đệm tìm kiếm hỗ trợ; thứ hai, giả '
   'thuyết co-rumination — định nghĩa là sự thảo luận sâu rộng, '
   'suy tư phỏng đoán và tập trung vào cảm xúc tiêu cực (Rose '
   '2002) — gợi ý rằng giãi bày thường xuyên với bạn bè có thể '
   'khuếch đại lo lắng theo vòng lặp thay vì tiêu tan, đặc biệt '
   'khi chuẩn mực cạnh tranh học thuật đang vận hành. Các diễn giải '
   'lý thuyết này phù hợp với bằng chứng liên văn hóa về xã hội hóa '
   'cảm xúc ở các mẫu thanh thiếu niên Đông Á.')


H3('4.5 Separation anxiety gender invariance — novel interpretation',
   '4.5 Bất biến giới của lo âu chia ly — diễn giải mới')

PP('The complete gender invariance for SAD (Female M = 24.76 vs Male '
   'M = 25.42; p = 0.620), against significant gender differences for '
   'GAD and SocAD, is the most theoretically distinctive finding of '
   'this study. We advance two non-mutually-exclusive explanations. '
   'First, cultural collectivism: Vietnamese family hierarchy '
   'creates uniform separation experiences across genders [Triandis '
   '1995]. Under interdependent self-construal [Markus & Kitayama '
   '1991], attachment processes operate primarily within the family '
   'system, where gender role differentiation is less salient for '
   'separation specifically than for socially evaluative behaviour. '
   'Second, developmental timing: separation-individuation tasks '
   'are predominantly childhood-onset and predate the post-pubertal '
   'differentiation of gender-loaded social pressures [Blos 1979; '
   'Kroger 2007]. Pre-pubertal separation challenges are universal '
   'across '
   'genders, whereas post-pubertal social evaluation and academic '
   'performance pressures are gendered.',
   'Bất biến giới hoàn toàn của SAD (nữ M = 24,76 so với nam M = '
   '25,42; p = 0,620), trái ngược với khác biệt giới có ý nghĩa của '
   'GAD và SocAD, là phát hiện mang tính khác biệt lý thuyết nổi bật '
   'nhất của nghiên cứu. Chúng tôi đưa ra hai cách giải thích không '
   'loại trừ lẫn nhau. Thứ nhất, văn hóa tập thể: thứ bậc gia '
   'đình Việt Nam tạo trải nghiệm chia ly đồng nhất giữa các giới. '
   'Dưới khái niệm bản thân phụ thuộc lẫn nhau, các quá trình gắn bó '
   'vận hành chủ yếu trong hệ thống gia đình, nơi sự phân hóa vai trò '
   'giới kém nổi bật hơn đối với chia ly so với hành vi mang tính '
   'đánh giá xã hội. Thứ hai, thời điểm phát triển: các nhiệm vụ '
   'chia ly – cá nhân hóa chủ yếu khởi phát từ thời thơ ấu và đi '
   'trước sự phân hóa giới của các áp lực xã hội sau dậy thì (Blos '
   '1979; Kroger 2007). Các thách thức chia ly tiền dậy thì phổ quát '
   'giữa các giới, trong khi '
   'áp lực đánh giá xã hội và thành tích học tập sau dậy thì mang '
   'tính giới.')


H3('4.6 Clinical and educational implications',
   '4.6 Hàm ý lâm sàng và giáo dục')

PP('Three implications follow. First, a tiered prevention '
   'architecture is warranted: universal classroom-based intervention '
   'for risk-factor reduction (academic pressure management, '
   'anti-bullying programmes, smartphone-use boundaries) plus '
   'targeted self-esteem-building modules for high-anxiety students. '
   'Second, gender-aware delivery is appropriate for GAD and '
   'SocAD modules but unnecessary for SAD content — a calibration '
   'that may improve cost-efficiency. Third, anti-bullying '
   'programmes should explicitly incorporate separation-anxiety '
   'screening, given the dominant bullying → SAD pathway documented '
   'here.',
   'Ba hàm ý sau đây. Thứ nhất, một kiến trúc phòng ngừa phân '
   'tầng là cần thiết: can thiệp dựa trên lớp học phổ quát để giảm '
   'yếu tố nguy cơ (quản lý áp lực học tập, chương trình chống bắt '
   'nạt, ranh giới sử dụng điện thoại) kết hợp với module xây dựng '
   'tự trọng cho học sinh có mức lo âu cao. Thứ hai, triển khai '
   'nhạy với giới phù hợp cho module GAD và SocAD nhưng không cần '
   'thiết cho nội dung SAD — một sự hiệu chỉnh có thể cải thiện '
   'hiệu suất chi phí. Thứ ba, chương trình chống bắt nạt cần '
   'tích hợp sàng lọc lo âu chia ly một cách rõ ràng, do đường dẫn '
   'bắt nạt → SAD chiếm ưu thế được ghi nhận tại đây.')


H3('4.7 Limitations and future directions',
   '4.7 Hạn chế và hướng nghiên cứu tương lai')

PP('Several limitations bound the present findings. (1) The '
   'cross-sectional design precludes causal inference; the '
   'single-timepoint measurement does not permit temporal ordering '
   'of risk-protective and anxiety pathways. A longitudinal cohort '
   'following the same students from Grade 6 to Grade 9 would be '
   'required to establish directional and developmental sequencing. '
   '(2) The sample is drawn from two Hanoi schools and may not '
   'generalise to rural or southern Vietnamese contexts; multi-site '
   'replication is needed. (3) Self-report measures of bullying and '
   'anxiety carry known shared-method variance risk; future studies '
   'should integrate peer-nomination bullying measures and clinical '
   'diagnostic anxiety interviews. (4) Future randomised controlled '
   'trials should test the eight-module school-based prevention '
   'curriculum ("Khung Chương trình"; abbreviated Khung CT) derived '
   'from these mechanistic findings, ideally with the OurFutures '
   'Mental Health Australia template adapted for the Vietnamese '
   'context [Grummitt et al. 2025].',
   'Một số hạn chế giới hạn các phát hiện hiện tại. (1) Thiết kế cắt '
   'ngang loại trừ suy luận nhân quả; đo lường tại một thời điểm '
   'duy nhất không cho phép sắp xếp thời gian của các đường dẫn nguy '
   'cơ – bảo vệ và lo âu. Một đoàn hệ dọc theo dõi cùng học sinh từ '
   'khối 6 đến khối 9 sẽ là cần thiết để thiết lập trình tự hướng và '
   'phát triển. (2) Mẫu lấy từ hai trường ở Hà Nội và có thể không '
   'khái quát hóa cho bối cảnh nông thôn hoặc miền Nam Việt Nam; cần '
   'nhân rộng đa địa điểm. (3) Các đo lường tự báo cáo về bắt nạt và '
   'lo âu mang rủi ro phương sai phương pháp chia sẻ; nghiên cứu '
   'tương lai cần tích hợp đo lường bắt nạt theo đề cử bạn bè và '
   'phỏng vấn chẩn đoán lâm sàng về lo âu. (4) Các thử nghiệm ngẫu '
   'nhiên có đối chứng trong tương lai cần kiểm định chương trình '
   'giảng dạy phòng ngừa 8 nội dung Khung CT rút ra từ các phát hiện '
   'cơ chế này, lý tưởng nhất với mô hình OurFutures Mental Health '
   'Úc được điều chỉnh cho bối cảnh Việt Nam.')


d.add_page_break()


# ============================================================
H2('5. REFERENCES (selected verified citations)',
   '5. TÀI LIỆU THAM KHẢO (trích dẫn đã verify)')

P_ref = ['American Psychiatric Association. (2013). Diagnostic and '
         'Statistical Manual of Mental Disorders (5th ed.). Arlington, '
         'VA: APA Publishing.',
         'Anderson, E. M., et al. (2025). Narrative review of '
         'adolescent anxiety in low- and middle-income contexts.',
         'Blos, P. (1979). The adolescent passage: Developmental issues. '
         'New York: International Universities Press.',
         'Bowlby, J. (1988). A secure base: Parent-child attachment and '
         'healthy human development. New York: Basic Books.',
         'Cheung, G. W., & Rensvold, R. B. (2002). Evaluating goodness-'
         'of-fit indexes for testing measurement invariance.',
         'Chen, Z., Ren, S., He, R., et al. (2023). Prevalence and '
         'associated factors of depressive and anxiety symptoms among '
         'Chinese secondary school students. BMC Psychiatry, 23(1), 580.',
         'Chorpita, B. F. (2000). RCADS: Revised Children\'s Anxiety '
         'and Depression Scale.',
         'Compas, B. E., Jaser, S. S., Bettis, A. H., et al. (2017). '
         'Coping, emotion regulation, and psychopathology in '
         'childhood and adolescence: A meta-analysis and narrative '
         'review. Psychological Bulletin, 143(9), 939–991.',
         'Goodenow, C. (1993). PSSM: Psychological Sense of School '
         'Membership.',
         'Grummitt, L., O\'Dean, S., Birrell, L., et al. (2025). '
         'Efficacy of a school-based, universal prevention programme '
         'for depression and anxiety in adolescents (OurFutures '
         'Mental Health). eClinicalMedicine.',
         'Hoang Trung Hoc. (2025). Tỷ lệ rối loạn lo âu ở học sinh '
         'Việt Nam (N = 8.389 analytic).',
         'GBD ASEAN. (2025). Burden of mental disorders in ASEAN '
         'children and adolescents: Global Burden of Disease analysis '
         '1990–2021. Lancet Regional Health – Western Pacific.',
         'Hu, L., & Bentler, P. M. (1999). Cutoff criteria for fit '
         'indexes in covariance structure analysis.',
         'Kieling, C., Buchweitz, C., Caye, A., et al. (2024). '
         'Worldwide prevalence and disability from mental disorders '
         'across childhood and adolescence: Evidence from the Global '
         'Burden of Disease Study. JAMA Psychiatry, 81(4), 347-356.',
         'Kroger, J. (2007). Identity development: Adolescence through '
         'adulthood (2nd ed.). Thousand Oaks, CA: Sage.',
         'Kwon, M., Lee, J. Y., Won, W. Y., et al. (2013). SAS-SV: '
         'Smartphone Addiction Scale–Short Version.',
         'Lazarus, R. S., & Folkman, S. (1984). Stress, appraisal, '
         'and coping.',
         'Markus, H. R., & Kitayama, S. (1991). Culture and the self: '
         'Implications for cognition, emotion, and motivation.',
         'McLean, C. P., et al. (2011). Gender differences in anxiety '
         'disorders: Prevalence, course of illness, comorbidity and '
         'burden of illness.',
         'Niwenahisemo, L. C., Hong, S., & Kuang, L. (2024). Assessing'
         'anxiety symptom severity in Rwandese adolescents: '
         'Cross-gender measurement invariance of GAD-7. Frontiers in '
         'Psychiatry.',
         'Olweus, D. (1996). OBVQ: Bully/Victim Questionnaire.',
         'Pascoe, M. C., et al. (2020). The impact of stress on '
         'students in secondary school and higher education.',
         'Robson, E. M., Husin, H. M., Dashti, S. G., et al. (2025). '
         'Tracking the course of depressive and anxiety symptoms '
         'across adolescence (the CATS study). Lancet Psychiatry, '
         '12(1), 44-53.',
         'Rose, A. J. (2002). Co-rumination in the friendships of '
         'girls and boys. Child Development, 73(6), 1830–1843. '
         'PMID: 12487497.',
         'Small, S., & Blanc, J. (2021). Mental Health During '
         'COVID-19: Tam Giao and Vietnam\'s Response. Frontiers '
         'in Psychiatry, 11, 589618. DOI: 10.3389/fpsyt.2020.589618.',
         'Stankov, L. (2010). Unforgiving Confucian culture: A '
         'breeding ground for high academic achievement, test '
         'anxiety and self-doubt? Learning and Individual '
         'Differences, 20(6), 555–563. DOI: 10.1016/j.lindif.2010.05.003.',
         'Rosenberg, M. (1965). Rosenberg Self-Esteem Scale.',
         'Saikia, S., et al. (2023). Prevalence and correlates of '
         'anxiety in Northeast Indian adolescents. Indian Journal '
         'of Community Medicine.',
         'Steare, T., et al. (2023). The association between academic '
         'pressure and adolescent mental health: A systematic review.',
         'Sun, J., Dunne, M. P., Hou, X. Y., & Xu, A. (2011). ESSA: '
         'Educational Stress Scale for Adolescents.',
         'Triandis, H. C. (1995). Individualism and collectivism.',
         'V-NAMHS. (2022). Vietnam Adolescent Mental Health Survey '
         'main report.',
         'Wen, F., et al. (2020). Mental health among Chinese rural '
         'adolescents.',
         'Xu, J., et al. (2021). Prevalence of mental disorders among '
         'Chinese adolescents (N = 373,216).',
         'Zimet, G. D., et al. (1988). MSPSS: Multidimensional Scale '
         'of Perceived Social Support.',
         ]
for ref in P_ref:
    p = d.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5); p.paragraph_format.first_line_indent = Cm(-0.5)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(ref); r.font.name = 'Times New Roman'; r.font.size = Pt(11)


# ============================================================
H2('NOTES TO CO-AUTHORS - 3 BLOCKING ITEMS',
   'GHI CHU CHO DONG TAC GIA - 3 MUC CHO QUYET')

TBD('Q1-8 - Integrated SEM R^2 interpretation: present draft uses LA '
    'Bang 47 composite R^2 = 0.598 (YTNC + YTBV -> RLLA). Confirm if '
    'acceptable OR re-analyse 7x3 integrated SEM',
    'Q1-8 - Dien giai R^2 SEM tich hop: ban nhap hien tai dung R^2 tong '
    'hop = 0,598 (YTNC + YTBV -> RLLA) tu LA Bang 47. Xac nhan chap '
    'nhan HOAC tai phan tich SEM tich hop 7x3')

TBD('Q3-6 - IRB approval letter HNUE: decision number + date',
    'Q3-6 - Letter chap thuan dao duc HNUE: so quyet dinh + ngay')

TBD('Q3-9 - Submission strategy Q2 (Frontiers in Psychiatry): confirm '
    'target journal section (e.g., Mood Disorders, Anxiety and '
    'Stress Disorders, Public Mental Health) and submission timeline',
    'Q3-9 - Chien luoc nop Q2 (Frontiers in Psychiatry): xac nhan '
    'section tap chi muc tieu (vi du: Mood Disorders, Anxiety and '
    'Stress Disorders, Public Mental Health) va timeline nop bai')


# Clean metadata
cp = d.core_properties
cp.author = ''; cp.title = ''; cp.subject = ''
cp.keywords = ''; cp.comments = ''; cp.last_modified_by = ''
cp.category = ''; cp.identifier = ''; cp.version = ''

d.save(OUT)
print(f'Saved: {OUT}')
print(f'Size: {os.path.getsize(OUT)} bytes')
