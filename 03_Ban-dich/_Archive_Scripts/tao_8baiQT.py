# -*- coding: utf-8 -*-
"""Dịch + nhận xét 8 bài quốc tế (Mỹ/Anh/Úc/Châu Âu)"""
import sys, io, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '02_Papers-goc', 'The-gioi-moi')

def read_pdf_pages(fname, max_chars=2000):
    import fitz
    path = os.path.join(BASE, fname)
    doc = fitz.open(path)
    pages = []
    for p in doc:
        pages.append(p.get_text()[:max_chars])
    return pages, len(doc)

# ============================================================
# 1. NORWAY 2011-2024
# ============================================================
d = create_doc()
add_heading(d, 'Giải thích khả thi cho xu hướng tăng căng thẳng tâm thần ở thanh thiếu niên Na Uy từ 2011 đến 2024', 1)
add_heading(d, 'Possible Explanations for the Upward Trend in Mental Distress among Adolescents in Norway from 2011 to 2024', 2)
add_heading(d, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d, [
    ('Tiêu đề gốc', 'Possible explanations for the upward trend in mental distress among adolescents in Norway from 2011 to 2024'),
    ('Tạp chí', 'Social Science & Medicine, 2025 (Q1, IF \u2248 5.4)'),
    ('Quốc gia', 'Na Uy (Norway)'),
    ('Số trang', '10'),
    ('Đặc điểm', 'Phân tích xu hướng 13 năm, vai trò trường học + mạng xã hội'),
])
add_heading(d, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d, '\u2022 Xu hướng 13 năm (2011\u20132024): căng thẳng tâm thần tăng ở cả nam và nữ VTN Na Uy.', bold=True)
add_p(d, '\u2022 Bất mãn với trường học giải thích phần lớn xu hướng tăng.')
add_p(d, '\u2022 Thời gian dùng mạng xã hội cũng giải thích một phần.')
add_p(d, '\u2022 Social Science & Medicine Q1 — tạp chí uy tín cao.')
add_p(d, '\u2022 Hạn chế: Chỉ Na Uy, mô hình giải thích (decomposition) không phải nhân quả.')
add_red_heading(d, 'PHẢN BIỆN')
add_red(d, 'Điểm mạnh: Q1 IF=5.4, xu hướng 13 năm, phân tích decomposition nâng cao.')
add_red(d, 'Hạn chế: Chỉ Na Uy — không khái quát cho châu Á. Không đo lo âu trực tiếp (đo mental distress chung).')
add_red(d, 'Gap: So sánh xu hướng Na Uy với Việt Nam/châu Á. Đánh giá can thiệp giảm mạng xã hội.')
d.save('21_Norway_2025_SocSciMed.docx')
sys.stderr.write('21_Norway OK\n')

# ============================================================
# 2. SCREEN TIME LONGITUDINAL — Br J Clin Psychol 2025
# ============================================================
d = create_doc()
add_heading(d, 'Mối liên quan cắt ngang và dọc giữa thời gian sử dụng màn hình với trầm cảm và lo âu ở thanh thiếu niên', 1)
add_heading(d, 'Cross-sectional and Longitudinal Associations of Screen Time with Adolescent Depression and Anxiety', 2)
add_heading(d, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d, [
    ('Tiêu đề gốc', 'Cross-sectional and longitudinal associations of screen time with adolescent depression and anxiety'),
    ('Tạp chí', 'British Journal of Clinical Psychology, 2025, 64:873\u2013887 (Q1, IF \u2248 3.0)'),
    ('Số trang', '15'),
    ('Đặc điểm', 'Nghiên cứu DỌC — screen time \u2192 trầm cảm/lo âu'),
])
add_heading(d, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d, '\u2022 Nghiên cứu DỌC (longitudinal) — có thể suy luận chiều hướng nhân quả.', bold=True)
add_p(d, '\u2022 Screen time liên quan với trầm cảm và lo âu ở VTN.')
add_p(d, '\u2022 BJCP Q1 — tạp chí tâm lý lâm sàng uy tín Anh quốc.')
add_p(d, '\u2022 Phù hợp với Chen 2023 (game OR=5,00) và Hoàng Trung Học 2025 (điện tử Beta=0,176).')
add_red_heading(d, 'PHẢN BIỆN')
add_red(d, 'Điểm mạnh: Thiết kế dọc — vượt trội so với cắt ngang. Q1.')
add_red(d, 'Hạn chế: Cần đọc chi tiết PDF để xác nhận cỡ mẫu, công cụ, kết quả cụ thể.')
add_red(d, 'Gap: So sánh tác động screen time ở các nền văn hóa khác nhau (châu Á vs phương Tây).')
d.save('22_ScreenTime_2025_BJCP.docx')
sys.stderr.write('22_ScreenTime OK\n')

# ============================================================
# 3. JAACAP US TRENDS
# ============================================================
d = create_doc()
add_heading(d, 'Xu hướng rối loạn tâm thần ở trẻ em và thanh thiếu niên nhận điều trị trong hệ thống SKTT công lập Hoa Kỳ', 1)
add_heading(d, 'Trends in Mental Disorders in Children and Adolescents Receiving Treatment in the State Mental Health System', 2)
add_heading(d, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d, [
    ('Tiêu đề gốc', 'Trends in Mental Disorders in Children and Adolescents Receiving Treatment in the State Mental Health System'),
    ('Tạp chí', 'Journal of the American Academy of Child & Adolescent Psychiatry (JAACAP), 2024 (Q1, IF \u2248 11.0)'),
    ('Quốc gia', 'Hoa Kỳ'),
    ('Số trang', '15'),
    ('Đặc điểm', 'Xu hướng chẩn đoán 2013\u20132021, hệ thống y tế công lập Mỹ'),
])
add_heading(d, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d, '\u2022 JAACAP — tạp chí hàng đầu về tâm thần trẻ em (IF \u2248 11,0).', bold=True)
add_p(d, '\u2022 Xu hướng chẩn đoán 8 năm (2013\u20132021) trong hệ thống y tế công Mỹ.')
add_p(d, '\u2022 So sánh được với NSCH 2020 (lo âu tăng 61% tại Mỹ).')
add_red_heading(d, 'PHẢN BIỆN')
add_red(d, 'Điểm mạnh: JAACAP Q1 IF=11, dữ liệu hệ thống quốc gia Mỹ, xu hướng dài hạn.')
add_red(d, 'Hạn chế: Chỉ trẻ em ĐANG điều trị — thiên lệch chọn (không bao gồm trẻ chưa được chẩn đoán).')
add_red(d, 'Gap: So sánh hệ thống y tế công Mỹ với Việt Nam (8,4% tiếp cận dịch vụ).')
d.save('23_JAACAP_US_Trends_2024.docx')
sys.stderr.write('23_JAACAP OK\n')

# ============================================================
# 4. WHO EUROPE — Lancet Regional Health Europe 2025
# ============================================================
d = create_doc()
add_heading(d, 'Sức khỏe tâm thần của trẻ em và thanh niên trong khu vực WHO châu Âu', 1)
add_heading(d, 'Mental Health of Children and Young People in the WHO Europe Region', 2)
add_heading(d, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d, [
    ('Tiêu đề gốc', 'Mental health of children and young people in the WHO Europe region'),
    ('Tạp chí', 'The Lancet Regional Health \u2014 Europe, 2025 (Q1, IF \u2248 15.0)'),
    ('Quốc gia', 'Toàn bộ khu vực WHO châu Âu'),
    ('Số trang', '13'),
    ('Đặc điểm', 'Tổng quan toàn diện, 9 triệu VTN châu Âu có rối loạn SKTT'),
])
add_heading(d, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d, '\u2022 Lancet Regional Health Europe — tạp chí uy tín rất cao (IF \u2248 15,0).', bold=True)
add_p(d, '\u2022 9 triệu VTN châu Âu (10\u201319 tuổi) sống với rối loạn SKTT.')
add_p(d, '\u2022 Lo âu + trầm cảm chiếm >50% các ca rối loạn.')
add_p(d, '\u2022 Tổng quan toàn diện cho chính sách — rất phù hợp cho phần Tổng quan tài liệu của Đề tài.')
add_red_heading(d, 'PHẢN BIỆN')
add_red(d, 'Điểm mạnh: Lancet Q1 IF=15, phạm vi toàn châu Âu, hướng chính sách.')
add_red(d, 'Hạn chế: Tổng quan chính sách, không phải nghiên cứu gốc — thiếu số liệu cụ thể tại từng nước.')
add_red(d, 'Gap: So sánh mô hình can thiệp châu Âu với ASEAN/Việt Nam.')
d.save('24_WHO_Europe_2025_LancetRegional.docx')
sys.stderr.write('24_WHO_Europe OK\n')

# ============================================================
# 5. EPIDEMIOLOGY & PSYCHIATRIC SCIENCES
# ============================================================
d = create_doc()
add_heading(d, 'Bài báo từ Epidemiology and Psychiatric Sciences', 1)
add_heading(d, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d, [
    ('Tạp chí', 'Epidemiology and Psychiatric Sciences, 2025 (Q1, IF \u2248 7.0)'),
    ('Số trang', '11'),
    ('Ghi chú', 'Cần đọc chi tiết PDF để xác định nội dung cụ thể'),
])
add_heading(d, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d, '\u2022 Epidemiology and Psychiatric Sciences — Q1 IF \u2248 7,0.')
add_p(d, '\u2022 11 trang — cần đọc chi tiết để nhận xét.', bold=True)
add_red_heading(d, 'GHI CHÚ')
add_red(d, 'Cần đọc chi tiết PDF (S2045796025000083a.pdf) để xác định tiêu đề, tác giả, và nội dung cụ thể.')
d.save('25_EpiPsychSci_2025.docx')
sys.stderr.write('25_EpiPsychSci OK\n')

# ============================================================
# 6. UK NHS MENTAL HEALTH STATISTICS (46 trang — tóm tắt)
# ============================================================
d = create_doc()
add_heading(d, 'Thống kê sức khỏe tâm thần: Tỷ lệ, dịch vụ và tài trợ tại Anh', 1)
add_heading(d, 'Mental Health Statistics: Prevalence, Services and Funding in England', 2)
add_heading(d, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d, [
    ('Tiêu đề', 'Mental health statistics: prevalence, services and funding in England'),
    ('Nguồn', 'House of Commons Library, UK Parliament, 2025'),
    ('Tác giả', 'Carl Baker'),
    ('Số trang', '46 (báo cáo chính sách)'),
    ('Đặc điểm', 'NHS Survey 2025: 25,8% thanh niên 16\u201324 rối loạn TT, tăng từ 18,9% (2014)'),
])
add_heading(d, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d, '\u2022 Báo cáo chính sách từ Quốc hội Anh — nguồn chính thức, đáng tin cậy nhất.', bold=True)
add_p(d, '\u2022 NHS Mental Health Survey 2025: 25,8% thanh niên 16\u201324 mắc rối loạn tâm thần phổ biến.')
add_p(d, '\u2022 Tăng từ 18,9% (2014) lên 25,8% (2025) — tăng 36% trong 11 năm.')
add_p(d, '\u2022 Nữ bị ảnh hưởng nhiều hơn: 36,1% vs nam 16,3%.')
add_p(d, '\u2022 Tự hại tăng gấp 4 lần từ năm 2000, 31,7% nữ 16\u201324 tuổi.')
add_p(d, '\u2022 46 trang — báo cáo toàn diện, không phải nghiên cứu gốc.')
add_red_heading(d, 'PHẢN BIỆN')
add_red(d, 'Điểm mạnh: Nguồn chính thức Quốc hội Anh, dữ liệu NHS quốc gia, xu hướng dài hạn.')
add_red(d, 'Hạn chế: Báo cáo tổng hợp chính sách, không phải NC peer-reviewed. Chỉ Anh.')
add_red(d, 'Gap: So sánh hệ thống NHS Anh với Việt Nam về tiếp cận dịch vụ SKTT.')
d.save('26_UK_NHS_2025_Parliament.docx')
sys.stderr.write('26_UK_NHS OK\n')

# ============================================================
# 7. NATURE HUMAN BEHAVIOUR — Social media
# ============================================================
d = create_doc()
add_heading(d, 'Sử dụng mạng xã hội ở thanh thiếu niên có và không có tình trạng sức khỏe tâm thần', 1)
add_heading(d, 'Social Media Use in Adolescents with and without Mental Health Conditions', 2)
add_heading(d, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d, [
    ('Tiêu đề gốc', 'Social media use in adolescents with and without mental health conditions'),
    ('Tạp chí', 'Nature Human Behaviour, Vol. 9, June 2025, pp. 1283\u20131299 (Q1, IF \u2248 24.0)'),
    ('Quốc gia', 'Vương quốc Anh'),
    ('Số trang', '21'),
    ('Đặc điểm', 'So sánh VTN có/không rối loạn TT về hành vi mạng xã hội. Nature — IF cực cao.'),
])
add_heading(d, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d, '\u2022 Nature Human Behaviour — IF \u2248 24,0, tạp chí TOP toàn cầu.', bold=True)
add_p(d, '\u2022 VTN có rối loạn SKTT dùng mạng xã hội nhiều hơn và ít hài lòng hơn.')
add_p(d, '\u2022 Nghiên cứu tại Anh, quốc gia đại diện (N = 3.340, 11\u201319 tuổi).')
add_p(d, '\u2022 Bao gồm đánh giá chẩn đoán bởi chuyên gia lâm sàng.')
add_p(d, '\u2022 Phù hợp với Chen 2023 (game OR=5,00) và Hoàng Trung Học 2025 (điện tử Beta=0,176).')
add_red_heading(d, 'PHẢN BIỆN')
add_red(d, 'Điểm mạnh: Nature Q1 IF=24, mẫu quốc gia UK, chẩn đoán lâm sàng, thiết kế mạnh.')
add_red(d, 'Hạn chế: Chỉ UK — khác biệt văn hóa với châu Á. Cắt ngang — không suy luận nhân quả.')
add_red(d, 'Gap: Nghiên cứu tương tự tại VN so sánh VTN có/không SKTT về hành vi mạng xã hội.')
d.save('27_NatureHumanBehav_SocialMedia_2025.docx')
sys.stderr.write('27_Nature OK\n')

# ============================================================
# 8. AJP — Pediatric Anxiety Treatment
# ============================================================
d = create_doc()
add_heading(d, 'Phương pháp tiếp cận hiện tại và tương lai trong điều trị rối loạn lo âu ở trẻ em', 1)
add_heading(d, 'Current and Future Approaches to Pediatric Anxiety Disorder Treatment', 2)
add_heading(d, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d, [
    ('Tiêu đề gốc', 'Current and Future Approaches to Pediatric Anxiety Disorder Treatment'),
    ('Tạp chí', 'American Journal of Psychiatry (AJP), 2024 (Q1, IF \u2248 18.0)'),
    ('Số trang', '12'),
    ('Đặc điểm', 'Tổng quan CBT, thuốc, hướng mới trong điều trị lo âu trẻ em'),
])
add_heading(d, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d, '\u2022 AJP — tạp chí tâm thần hàng đầu thế giới (IF \u2248 18,0).', bold=True)
add_p(d, '\u2022 Tổng quan toàn diện: CBT, SSRI, digital interventions, neuromodulation.')
add_p(d, '\u2022 CBT hiệu quả 47\u201366% phục hồi, 57\u201360% đáp ứng.')
add_p(d, '\u2022 Phù hợp với Zhameden 2025 (CBT 3/4 trầm cảm, 1/4 lo âu ở LMIC).')
add_p(d, '\u2022 Hướng mới: can thiệp kỹ thuật số, kích thích thần kinh không xâm lấn.')
add_red_heading(d, 'PHẢN BIỆN')
add_red(d, 'Điểm mạnh: AJP Q1 IF=18, tổng quan toàn diện cập nhật nhất về điều trị.')
add_red(d, 'Hạn chế: Tổng quan (review) — không có dữ liệu gốc mới. Chủ yếu từ bối cảnh phương Tây.')
add_red(d, 'Gap: Đánh giá hiệu quả CBT tại Việt Nam (hiện 0 RCT). Can thiệp kỹ thuật số cho VTN VN.')
d.save('28_AJP_PediatricAnxiety_2024.docx')
sys.stderr.write('28_AJP OK\n')

# ============================================================
# 9. CBT NETWORK META-ANALYSIS — BMC Psychiatry 2025
# ============================================================
d = create_doc()
add_heading(d, 'Hiệu quả các loại can thiệp khác nhau cho rối loạn lo âu ở trẻ em và thanh thiếu niên: Tổng quan hệ thống và phân tích tổng hợp mạng Bayesian', 1)
add_heading(d, 'Effects of Different Interventions on Anxiety Disorders in Children and Adolescents: A Systematic Review and Bayesian Network Meta-analysis', 2)
add_heading(d, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d, [
    ('Tiêu đề gốc', 'Effects of different interventions on anxiety disorders in children and adolescents: a systematic review and Bayesian network meta-analysis'),
    ('Tạp chí', 'BMC Psychiatry, 2025 (Q1, Open Access)'),
    ('Số trang', '14'),
    ('Đặc điểm', '30 RCTs, 1.711 HS — so sánh CBT, thuốc, kết hợp, tâm lý giáo dục'),
])
add_heading(d, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d, '\u2022 Network meta-analysis — so sánh ĐỒNG THỜI nhiều loại can thiệp.', bold=True)
add_p(d, '\u2022 30 RCTs, 1.711 trẻ — bằng chứng mạnh.')
add_p(d, '\u2022 BMC Psychiatry Q1 Open Access.')
add_p(d, '\u2022 Bổ sung cho Zhameden 2025 (6 RCTs LMIC, GRADE rất thấp).')
add_red_heading(d, 'PHẢN BIỆN')
add_red(d, 'Điểm mạnh: Network meta Bayesian — phương pháp nâng cao nhất. BMC Q1. 30 RCTs.')
add_red(d, 'Hạn chế: Cần đọc chi tiết kết quả xếp hạng can thiệp. Có thể thiên lệch xuất bản.')
add_red(d, 'Gap: 0 RCT từ Việt Nam trong phân tích — cần thêm bằng chứng LMIC/châu Á.')
d.save('29_CBT_NetworkMeta_2025_BMCPsych.docx')
sys.stderr.write('29_CBT OK\n')
