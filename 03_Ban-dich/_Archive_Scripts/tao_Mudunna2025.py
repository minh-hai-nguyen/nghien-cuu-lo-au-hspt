# -*- coding: utf-8 -*-
"""Bài 08: Mudunna 2025 - Lancet Regional Health SEA - South Asia Adolescent MH"""
import sys, io, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()
add_link_and_qr(doc, 'https://doi.org/10.1016/S2772-3682(25)00003-4', 'QR_Hoa2024.png')

add_heading(doc, 'Bản chất, tỷ lệ hiện mắc và các yếu tố quyết định của các vấn đề sức khỏe tâm thần ở vị thành niên Nam Á: Tổng quan hệ thống', 1)

add_info_table(doc, [
    ('Tiêu đề gốc', 'Nature, prevalence and determinants of mental health problems experienced by adolescents in south Asia: a systematic review'),
    ('Tác giả', 'Chethana Mudunna, Medhavi Weerasinghe, Thach Tran, Josefine Antoniades, Lorena Romero, Miyuru Chandradasa, Jane Fisher'),
    ('Cơ quan', 'Monash University (Úc), La Trobe University, Alfred Health, University of Kelaniya (Sri Lanka)'),
    ('Tạp chí', 'The Lancet Regional Health \u2014 Southeast Asia, 2025'),
    ('DOI', '10.1016/S2772-3682(25)00003-4'),
    ('Phạm vi', 'Tổng quan hệ thống (Systematic Review) \u2014 8 quốc gia Nam Á: Afghanistan, Bangladesh, Bhutan, Ấn Độ, Maldives, Nepal, Pakistan, Sri Lanka'),
    ('Đối tượng', 'Vị thành niên (VTN) 10\u201319 tuổi tại Nam Á'),
    ('Thiết kế', 'Tổng quan hệ thống theo PRISMA (Preferred Reporting Items for Systematic Reviews and Meta-Analyses)'),
])

doc.add_paragraph()
add_page_ref(doc, '1', 'Lancet Regional Health \u2014 SEA', '2025')

add_heading(doc, 'TÓM TẮT', 2)
add_p(doc, 'Bối cảnh: Vị thành niên (VTN \u2014 10\u201319 tuổi) là giai đoạn nhạy cảm đặc biệt với các vấn đề sức khỏe tâm thần (SKTT). Nam Á, nơi sinh sống của 24% dân số thế giới, chủ yếu gồm các nước thu nhập thấp và trung bình (LMIC). Tổng quan hệ thống này đánh giá bản chất, tỷ lệ và các yếu tố quyết định của các vấn đề SKTT ở VTN Nam Á.')
add_p(doc, 'Phương pháp: Tìm kiếm có hệ thống trên MEDLINE, PsycINFO, Embase, CINAHL, Global Health theo hướng dẫn PRISMA. Tỷ lệ hiện mắc dao động rất lớn: 1,5\u201381,6% tùy phương pháp và quốc gia.')
add_p(doc, 'Kết quả: Trầm cảm và lo âu là phổ biến nhất. Các yếu tố quyết định bao gồm: giới tính nữ, nghèo đói, bạo lực gia đình, bắt nạt, mất cha mẹ, thiếu hỗ trợ xã hội, và áp lực học tập. Đặc biệt, kỳ thị SKTT và rào cản văn hóa cản trở tiếp cận dịch vụ.')
add_p(doc, 'Kết luận: Cần nghiên cứu chất lượng cao hơn với phương pháp chuẩn hóa tại Nam Á. Chính sách SKTT cần tích hợp yếu tố xã hội và văn hóa đặc thù của khu vực.')

add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(doc, 'Điểm nổi bật:', bold=True)
add_p(doc, '\u2022 Tổng quan hệ thống đầu tiên tập trung vào VTN Nam Á \u2014 khu vực chiếm 24% dân số thế giới nhưng ít được nghiên cứu (WHO, 2022).', size=11)
add_p(doc, '\u2022 Đăng trên Lancet Regional Health \u2014 tạp chí uy tín cao (GBD 2021 ASEAN Collaborators, 2025).', size=11)
add_p(doc, '\u2022 Bao phủ 8 quốc gia với đa dạng văn hóa, tôn giáo, kinh tế (Polanczyk và cộng sự, 2015).', size=11)
add_p(doc, '\u2022 Xác định rõ khoảng trống: tỷ lệ dao động 1,5\u201381,6% do thiếu chuẩn hóa phương pháp.', size=11)

add_p(doc, 'Hạn chế:', bold=True)
add_p(doc, '\u2022 Không có meta-analysis do tính không đồng nhất quá cao giữa các nghiên cứu (Higgins và cộng sự, 2003).', size=11)
add_p(doc, '\u2022 Phần lớn nghiên cứu từ Ấn Độ \u2014 thiếu dữ liệu từ Afghanistan, Bhutan, Maldives (Samuels và cộng sự, 2018).', size=11)
add_p(doc, '\u2022 Tập trung Nam Á \u2014 không bao gồm Đông Nam Á (Việt Nam, Indonesia, Philippines) (GBD 2021 ASEAN Collaborators, 2025).', size=11)

add_p(doc, 'Hướng cải thiện:', bold=True)
add_p(doc, '\u2022 Mở rộng sang Đông Nam Á để so sánh liên khu vực (GBD 2021 ASEAN Collaborators, 2025).', size=11)
add_p(doc, '\u2022 Chuẩn hóa công cụ đo lường xuyên quốc gia Nam Á (Polanczyk và cộng sự, 2015).', size=11)
add_p(doc, '\u2022 RCT can thiệp phù hợp văn hóa tại trường học Nam Á (Werner-Seidler và cộng sự, 2021; Zhameden và cộng sự, 2025).', size=11)

doc.add_paragraph()

add_page_ref(doc, '2\u20133', 'Lancet Regional Health \u2014 SEA', '2025')
add_heading(doc, 'GIỚI THIỆU', 2)
add_p(doc, 'Trầm cảm và lo âu là nguyên nhân hàng đầu gây bệnh tâm thần và khuyết tật ở VTN. Tự tử \u2014 thường là hệ quả của trầm cảm \u2014 là nguyên nhân tử vong thứ 4 ở VTN 15\u201319 tuổi (WHO, 2022). Hơn 80% người có vấn đề SKTT sống ở LMIC. Dự kiến đến 2030, trầm cảm sẽ là nguyên nhân gánh nặng bệnh tật thứ 3 tại nước thu nhập thấp và thứ 2 tại nước thu nhập trung bình.')
add_p(doc, 'Nam Á gồm 8 quốc gia: Afghanistan, Bangladesh, Bhutan, Ấn Độ, Maldives, Nepal, Pakistan, Sri Lanka. 7/8 nước được World Bank xếp hạng LMIC. Dữ liệu về SKTT VTN tại khu vực này rất thiếu hoặc không nhất quán.')
add_p(doc, 'Khung sinh thái xã hội (social-ecological framework) được sử dụng để phân tích: yếu tố cá nhân (sinh học, tâm lý), gia đình (bạo lực, cấu trúc), trường học (bắt nạt, áp lực), cộng đồng (nghèo đói, xung đột), và vĩ mô (chính sách, văn hóa).')

add_page_ref(doc, '3\u20134', 'Lancet Regional Health \u2014 SEA', '2025')
add_heading(doc, 'PHƯƠNG PHÁP', 2)
add_p(doc, 'Tìm kiếm trên MEDLINE, PsycINFO, Embase, CINAHL, Global Health. Tiêu chí: nghiên cứu quan sát về VTN 10\u201319 tuổi tại 8 nước Nam Á, đo SKTT bằng công cụ chuẩn hóa, xuất bản bằng tiếng Anh. Đánh giá chất lượng bằng JBI (Joanna Briggs Institute). Tổng hợp tường thuật (narrative synthesis) do tính không đồng nhất cao.')

add_page_ref(doc, '5\u20138', 'Lancet Regional Health \u2014 SEA', '2025')
add_heading(doc, 'KẾT QUẢ', 2)
add_p(doc, 'Tỷ lệ SKTT ở VTN Nam Á:', bold=True)
add_p(doc, 'Tỷ lệ dao động rất lớn: 1,5\u201381,6% tùy quốc gia, phương pháp và định nghĩa. Phần lớn nghiên cứu từ Ấn Độ. Trầm cảm và lo âu là phổ biến nhất. Một số nghiên cứu báo cáo tỷ lệ cao bất thường (>50%) \u2014 có thể do dùng thang sàng lọc với ngưỡng thấp.')

add_p(doc, 'Yếu tố quyết định theo khung sinh thái xã hội:', bold=True)
add_table(doc,
    ['Cấp độ', 'Yếu tố nguy cơ', 'Yếu tố bảo vệ'],
    [
        ['Cá nhân', 'Nữ giới, tuổi lớn hơn, bệnh mạn tính', 'Lòng tự trọng cao, kỹ năng ứng phó'],
        ['Gia đình', 'Bạo lực gia đình, mất cha mẹ, xung đột gia đình', 'Hỗ trợ gia đình, giao tiếp tốt'],
        ['Trường học', 'Bắt nạt, áp lực học tập, kỷ luật hà khắc', 'Môi trường trường học tích cực'],
        ['Cộng đồng', 'Nghèo đói, bạo lực cộng đồng, thiên tai', 'Hỗ trợ xã hội, kết nối cộng đồng'],
        ['Vĩ mô', 'Kỳ thị SKTT, thiếu chính sách, xung đột vũ trang', 'Chính sách SKTT quốc gia'],
    ], widths=[2.5, 5.0, 5.0])

doc.add_paragraph()

add_page_ref(doc, '9\u201311', 'Lancet Regional Health \u2014 SEA', '2025')
add_heading(doc, 'THẢO LUẬN', 2)
add_p(doc, 'Sự dao động lớn (1,5\u201381,6%) phản ánh tình trạng thiếu chuẩn hóa trong nghiên cứu SKTT VTN Nam Á. Nhiều nghiên cứu sử dụng công cụ tự phát triển hoặc chưa được chuẩn hóa cho bối cảnh văn hóa địa phương \u2014 tương tự vấn đề tại Việt Nam (Vũ Mạnh Lợi và cộng sự, 2022).')
add_p(doc, 'Kỳ thị SKTT là rào cản lớn nhất \u2014 ở nhiều nước Nam Á, bệnh tâm thần bị coi là "ma quỷ nhập" hoặc "yếu đuối", khiến gia đình tìm đến thầy lang thay vì chuyên gia SKTT. Đây cũng là vấn đề tại Việt Nam và nhiều nước ASEAN.')

add_heading(doc, 'KẾT LUẬN', 2)
add_p(doc, 'Cần chuẩn hóa phương pháp nghiên cứu SKTT VTN tại Nam Á. Chính sách cần tích hợp khung sinh thái xã hội, giảm kỳ thị, và đầu tư vào can thiệp trường học phù hợp văn hóa. Bài học cho Đông Nam Á: đa dạng văn hóa đòi hỏi tiếp cận linh hoạt, không thể áp dụng mô hình phương Tây máy móc.')

add_abbreviation_table(doc, [
    ('MHP', 'Mental Health Problems \u2014 Vấn đề sức khỏe tâm thần'),
    ('LMIC', 'Low- and Middle-Income Countries \u2014 Nước thu nhập thấp và trung bình'),
    ('PRISMA', 'Preferred Reporting Items for Systematic Reviews and Meta-Analyses'),
    ('JBI', 'Joanna Briggs Institute \u2014 Viện Joanna Briggs (đánh giá chất lượng NC)'),
    ('SDH', 'Social Determinants of Health \u2014 Yếu tố quyết định xã hội về sức khỏe'),
    ('VTN', 'Vị thành niên (10\u201319 tuổi)'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('WHO', 'World Health Organization \u2014 Tổ chức Y tế Thế giới'),
])

doc.add_paragraph()
add_red_heading(doc, 'PHẢN BIỆN CHI TIẾT')
add_red(doc, 'Điểm mạnh:', bold=True)
add_red(doc, '\u2022 Tổng quan hệ thống đầu tiên về SKTT VTN tập trung riêng vào Nam Á \u2014 khu vực 24% dân số thế giới nhưng ít được nghiên cứu (WHO, 2022).')
add_red(doc, '\u2022 Đăng trên Lancet Regional Health \u2014 tạp chí uy tín, bình duyệt nghiêm ngặt.')
add_red(doc, '\u2022 Sử dụng khung sinh thái xã hội \u2014 phân tích đa tầng từ cá nhân đến vĩ mô (Bronfenbrenner, 1979).')
add_red(doc, '\u2022 Bao phủ 8 quốc gia với đa dạng văn hóa, tôn giáo, kinh tế \u2014 phạm vi rộng (Polanczyk và cộng sự, 2015).')

add_red(doc, 'Hạn chế:', bold=True)
add_red(doc, '\u2022 Không meta-analysis do heterogeneity quá cao \u2014 chỉ tổng hợp tường thuật, hạn chế sức mạnh bằng chứng (Higgins và cộng sự, 2003).')
add_red(doc, '\u2022 Phần lớn NC từ Ấn Độ \u2014 Afghanistan, Bhutan, Maldives gần như vắng bóng (Samuels và cộng sự, 2018).')
add_red(doc, '\u2022 Tỷ lệ dao động 1,5\u201381,6% \u2014 vô nghĩa nếu không chuẩn hóa công cụ (COVID-19 Mental Disorders Collaborators, 2021; Ferrari và cộng sự, 2013).')
add_red(doc, '\u2022 Chỉ Nam Á \u2014 thiếu so sánh với Đông Nam Á (Việt Nam, Indonesia, Philippines có bối cảnh tương tự) (GBD 2021 ASEAN Collaborators, 2025).')
add_red(doc, '\u2022 Thiếu đánh giá chất lượng can thiệp \u2014 chỉ mô tả tỷ lệ, không đánh giá hiệu quả can thiệp nào (Werner-Seidler và cộng sự, 2021).')

add_red(doc, 'Khoảng trống nghiên cứu (Gap):', bold=True)
add_red(doc, '\u2022 Cần tổng quan hệ thống tương tự cho Đông Nam Á \u2014 hiện chưa có tổng quan riêng cho ASEAN (GBD 2021 ASEAN Collaborators, 2025).')
add_red(doc, '\u2022 Chuẩn hóa công cụ đo: so sánh PHQ-9/GAD-7/DASS-21/DISC-5 trên cùng mẫu VTN Nam Á (Spitzer và cộng sự, 2006; Vũ Mạnh Lợi và cộng sự, 2022).')
add_red(doc, '\u2022 RCT can thiệp SKTT tại trường học Nam Á \u2014 tổng quan hệ thống của Zhameden và cộng sự (2025) chỉ tìm được 6 RCTs tại LMIC, không có bài nào từ Nam Á trừ Ấn Độ.')
add_red(doc, '\u2022 Nghiên cứu vai trò của kỳ thị SKTT \u2014 yếu tố cản trở lớn nhất tại Nam Á, cần can thiệp giảm kỳ thị dựa trên bằng chứng (Thornicroft và cộng sự, 2016).')
add_red(doc, '\u2022 So sánh Nam Á vs Đông Nam Á: cùng LMIC nhưng văn hóa rất khác \u2014 Phật giáo vs Hindu vs Hồi giáo ảnh hưởng khác nhau đến SKTT (Canino & Alegria, 2008).')

doc.save('08_Mudunna_2025_LancetSEA.docx')
