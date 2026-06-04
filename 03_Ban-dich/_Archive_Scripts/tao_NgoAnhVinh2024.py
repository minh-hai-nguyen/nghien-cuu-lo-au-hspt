# -*- coding: utf-8 -*-
"""Dịch Ngô Anh Vinh 2024 — DTTS Lạng Sơn, DASS-21 + ACEs"""
import sys, io, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()
add_link_and_qr(doc, 'https://doi.org/10.1016/j.jadr.2024.100795', 'QR_NgoAnhVinh2024.png')

add_heading(doc, 'Sức khỏe tâm thần ở thanh thiếu niên dân tộc thiểu số tại Việt Nam và các yếu tố liên quan: Nghiên cứu cắt ngang', 1)
add_heading(doc, 'Mental Health among Ethnic Minority Adolescents in Vietnam and Correlated Factors: A Cross-sectional Study', 2)

add_heading(doc, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(doc, [
    ('Tiêu đề gốc', 'Mental Health among Ethnic Minority Adolescents in Vietnam and Correlated Factors: A Cross-sectional Study'),
    ('Tiêu đề dịch', 'Sức khỏe tâm thần ở thanh thiếu niên dân tộc thiểu số tại Việt Nam và các yếu tố liên quan: Nghiên cứu cắt ngang'),
    ('Tác giả', 'Ngô Anh Vinh, Vũ Thị Mỹ Hạnh, Đỗ Thị Bích Vân, Dương Anh Tài, Đỗ Minh Loan, Lê Thị Thanh Thùy'),
    ('Cơ quan', 'Khoa Sức khỏe Vị thành niên, Bệnh viện Nhi Trung ương, Hà Nội\nKhoa Công tác Xã hội, Học viện Thanh thiếu niên Việt Nam'),
    ('Tạp chí', 'Journal of Affective Disorders Reports, Vol. 17, 2024, 100795 (Open Access, Elsevier)'),
    ('DOI', '10.1016/j.jadr.2024.100795'),
    ('Cỡ mẫu', '845 học sinh dân tộc thiểu số nội trú, tỉnh Lạng Sơn'),
    ('Công cụ', 'DASS-21 + Bảng hỏi ACEs (Trải nghiệm bất lợi thời thơ ấu)'),
])

# ĐÁNH GIÁ NHANH
add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(doc, 'Điểm nổi bật:', bold=True)
add_p(doc, '\u2022 Nghiên cứu đầu tiên về SKTT thanh thiếu niên DTTS nội trú tại Việt Nam — nhóm đặc biệt dễ bị tổn thương.')
add_p(doc, '\u2022 Tỷ lệ rất cao: lo âu 54,4%, trầm cảm 59,0%, căng thẳng 24,7%.')
add_p(doc, '\u2022 ACEs 48,9% — liên quan dương tính với tất cả 3 rối loạn.')
add_p(doc, '\u2022 Mối quan hệ bạn bè kém → tăng lo âu, căng thẳng, trầm cảm — yếu tố đặc thù trường nội trú.')
add_p(doc, '\u2022 Elsevier Open Access, hồi quy Logistic + Tobit đa biến.')

add_p(doc, 'Hạn chế:', bold=True)
add_p(doc, '\u2022 Chỉ 1 tỉnh (Lạng Sơn) — không đại diện cho tất cả DTTS Việt Nam.')
add_p(doc, '\u2022 DASS-21 sàng lọc, không chẩn đoán. Tỷ lệ cao có thể bao gồm dương tính giả.')
add_p(doc, '\u2022 Thiết kế cắt ngang — không suy luận nhân quả ACEs → SKTT.')
add_p(doc, '\u2022 Không phân tích giới tính chi tiết trong tỷ lệ lo âu.')

# TÓM TẮT
add_page_ref(doc, '1', 'J Affective Disorders Reports', 'Vol. 17, 2024')
add_heading(doc, 'TÓM TẮT', 2)
add_p(doc, 'Mục tiêu: Đo lường tỷ lệ trầm cảm, lo âu và căng thẳng ở thanh thiếu niên DTTS và xác định mối quan hệ giữa đặc điểm nhân khẩu, hành vi, trải nghiệm bất lợi thời thơ ấu (ACEs) với các rối loạn tâm thần.')
add_red(doc, 'Phương pháp: Nghiên cứu cắt ngang tại Lạng Sơn, 845 HS DTTS nội trú. DASS-21 (Cronbach alpha: trầm cảm 0,767; lo âu 0,792; căng thẳng 0,811) + Bảng hỏi ACEs. Hồi quy Logistic đa biến và hồi quy Tobit.')
add_red(doc, 'Kết quả: Căng thẳng 24,7%, lo âu 54,4%, trầm cảm 59,0%. ACEs 48,9% (TB = 1,1; SD = 1,8). Số ACEs tương quan dương với điểm lo âu (Coef. = 0,28), căng thẳng (Coef. = 0,28), trầm cảm (Coef. = 0,16). Bạn bè kém liên quan tăng cả 3 chỉ số. Nguồn thông tin từ bạn bè/TV bảo vệ lo âu và căng thẳng.')
add_p(doc, 'Kết luận: Tỷ lệ cao SKTT và ACEs ở DTTS đòi hỏi can thiệp khẩn cấp nhắm vào ACEs, cải thiện mối quan hệ bạn bè, và tăng cường dịch vụ SKTT tại trường nội trú.')

# NỘI DUNG
add_heading(doc, '1. GIỚI THIỆU', 2)
add_page_ref(doc, '1\u20132', 'J Affective Disorders Reports', 'Vol. 17, 2024')
add_p(doc, 'Thanh thiếu niên DTTS là nhóm dễ bị tổn thương đặc biệt — đối mặt với bất bình đẳng kinh tế xã hội, rào cản ngôn ngữ và văn hóa. Tại Việt Nam, UNICEF (2017) báo cáo 68,4% trẻ 1\u201314 tuổi trải qua kỷ luật thể chất hoặc tâm lý. Nghiên cứu trước (Thai et al., 2020) trên 4.957 VTN Việt Nam cho thấy 86% có ít nhất 1 ACE.')
add_p(doc, 'Tuy nhiên, dữ liệu về SKTT VTN DTTS nội trú vùng cao còn rất hạn chế. Nghiên cứu này nhằm lấp đầy khoảng trống đó tại Lạng Sơn — tỉnh vùng núi phía Bắc Việt Nam.')

add_heading(doc, '2. PHƯƠNG PHÁP', 2)
add_page_ref(doc, '2\u20133', 'J Affective Disorders Reports', 'Vol. 17, 2024')
add_p(doc, 'Nghiên cứu cắt ngang tại các trường nội trú DTTS tỉnh Lạng Sơn, n = 845. Dân tộc: Tày, Nùng và khác.')
add_p(doc, 'DASS-21 phiên bản tiếng Việt (Le et al., 2017). Ngưỡng: Trầm cảm ≥5, Lo âu ≥4, Căng thẳng ≥8 (theo Henry & Crawford, 2005).')
add_p(doc, 'ACEs: 10 loại trải nghiệm bất lợi (lạm dụng thể chất, cảm xúc, tình dục; bỏ mặc; bạo lực gia đình; nghiện chất; bệnh tâm thần cha mẹ; ly hôn; tù đày).')
add_p(doc, 'Phân tích: Hồi quy Logistic đa biến (OR) + hồi quy Tobit (Coef.) cho điểm liên tục.')

add_heading(doc, '3. KẾT QUẢ', 2)
add_page_ref(doc, '3\u20136', 'J Affective Disorders Reports', 'Vol. 17, 2024')

add_p(doc, 'Tỷ lệ SKTT và ACEs (N = 845):', bold=True)
add_table(doc,
    ['Tình trạng', 'Tỷ lệ', 'Ghi chú'],
    [['Trầm cảm', '59,0%', 'Cao nhất trong 3 chỉ số'],
     ['Lo âu', '54,4%', '\u2014'],
     ['Căng thẳng', '24,7%', '\u2014'],
     ['ACEs (≥1)', '48,9%', 'TB = 1,1; SD = 1,8; range 0\u20136']],
    widths=[3.5, 2.5, 5.5])
doc.add_paragraph()

add_p(doc, 'Yếu tố liên quan (hồi quy Tobit, điểm lo âu):', bold=True)
add_table(doc,
    ['Yếu tố', 'Coef./OR', 'KTC 95%', 'Ý nghĩa'],
    [['Số ACEs → lo âu', 'Coef. = 0,28', '\u2014', 'Dương, có ý nghĩa'],
     ['Số ACEs → căng thẳng', 'Coef. = 0,28', '\u2014', 'Dương, có ý nghĩa'],
     ['Số ACEs → trầm cảm', 'Coef. = 0,16', '\u2014', 'Dương, có ý nghĩa'],
     ['Bạn bè kém → lo âu', 'Coef. = 0,54\u20131,46', '\u2014', 'Fair/Poor đều tăng'],
     ['Bạn bè kém → trầm cảm', 'OR = 6,84***', '2,03\u201323,02', 'Poor: rất mạnh'],
     ['Sống ≥10 người → lo âu', 'OR = 0,51***', '0,33\u20130,77', 'Bảo vệ'],
     ['Hoạt động thể chất → lo âu', 'OR = 0,72**', '0,51\u20131,00', 'Bảo vệ'],
     ['Internet 5h/ngày → lo âu', 'OR = 0,43**', '0,21\u20130,86', 'Bảo vệ (vs cuối tuần)'],
     ['TT từ bạn bè → lo âu', 'OR = 0,57*', '\u2014', 'Bảo vệ'],
     ['TT từ TV → căng thẳng', 'Coef. = \u20130,82*', '\u2014', 'Bảo vệ']],
    widths=[4.0, 2.5, 2.5, 3.5])
doc.add_paragraph()

add_heading(doc, '4. THẢO LUẬN', 2)
add_page_ref(doc, '4\u20135', 'J Affective Disorders Reports', 'Vol. 17, 2024')
add_p(doc, 'ACEs 48,9% cao hơn trung bình các nước LMIC nhưng thấp hơn NC trước tại VN (Thai 2020: 86%) — có thể do HS nội trú ít về nhà, ít tiếp xúc ACEs.')
add_p(doc, 'Mối quan hệ bạn bè đặc biệt quan trọng trong bối cảnh trường nội trú — HS sống xa gia đình, bạn bè trở thành nguồn hỗ trợ chính. Bạn bè kém (Poor) tăng nguy cơ trầm cảm OR = 6,84 — rất mạnh.')
add_p(doc, 'Internet hàng ngày bảo vệ hơn cuối tuần (OR = 0,43\u20130,66) — có thể do sử dụng đều đặn giúp kết nối xã hội tốt hơn so với sử dụng dồn cuối tuần.')
add_p(doc, 'Nguồn thông tin từ trường KHÔNG liên quan SKTT — gợi ý vai trò SKTT của trường chưa đủ mạnh, cần cải thiện.')

add_heading(doc, '5. KẾT LUẬN', 2)
add_p(doc, 'Tỷ lệ rất cao SKTT (trầm cảm 59%, lo âu 54,4%) và ACEs (48,9%) ở VTN DTTS nội trú Lạng Sơn. Can thiệp cần nhắm vào: giảm ACEs, cải thiện mối quan hệ bạn bè, tăng cường dịch vụ SKTT và tư vấn tại trường nội trú.')

add_heading(doc, 'TÀI LIỆU THAM KHẢO', 2)
add_p(doc, '[Giữ nguyên TLTK tiếng Anh từ bài gốc \u2014 xem PDF trang 6\u20137]', italic=True)

add_abbreviation_table(doc, [
    ('DASS-21', 'Depression Anxiety Stress Scale 21-item \u2014 Thang đo Trầm cảm Lo âu Căng thẳng 21 mục'),
    ('ACEs', 'Adverse Childhood Experiences \u2014 Trải nghiệm Bất lợi Thời thơ ấu'),
    ('DTTS', 'Dân tộc thiểu số'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('OR', 'Odds Ratio \u2014 Tỷ số chênh'),
    ('Coef.', 'Coefficient \u2014 Hệ số hồi quy (Tobit)'),
    ('LMIC', 'Low- and Middle-Income Countries \u2014 Nước thu nhập thấp và trung bình'),
])

add_red_heading(doc, 'PHẢN BIỆN CHI TIẾT')

add_red(doc, 'Điểm mạnh:', bold=True)
add_red(doc, '1. Nghiên cứu đầu tiên về SKTT VTN DTTS nội trú Việt Nam \u2014 lấp đầy khoảng trống quan trọng.')
add_red(doc, '2. Kết hợp DASS-21 + ACEs \u2014 đánh giá đồng thời SKTT và trải nghiệm bất lợi.')
add_red(doc, '3. Hồi quy Logistic + Tobit đa biến \u2014 phương pháp phân tích nâng cao.')
add_red(doc, '4. Open Access (Elsevier, CC BY-NC-ND) \u2014 dễ tiếp cận.')
add_red(doc, '5. Cronbach alpha tốt (0,767\u20130,811) cho DASS-21 tiếng Việt.')

add_red(doc, 'Hạn chế:', bold=True)
add_red(doc, '1. Chỉ 1 tỉnh (Lạng Sơn), chỉ trường nội trú \u2014 không đại diện cho VTN DTTS nói chung.')
add_red(doc, '2. DASS-21 sàng lọc, ngưỡng thấp (lo âu ≥4) \u2014 tỷ lệ 54,4% có thể bao gồm nhiều dương tính giả. So với V-NAMHS 2022 (DISC-5: 2,3%), chênh lệch 24 lần.')
add_red(doc, '3. Không phân tích giới tính chi tiết \u2014 chỉ có OR giới tính cho căng thẳng (1,29) nhưng không ý nghĩa (P > 0,05).')
add_red(doc, '4. ACEs tự báo cáo hồi cứu \u2014 recall bias, có thể đánh giá thấp ACEs thực tế.')
add_red(doc, '5. Thiết kế cắt ngang \u2014 không xác định chiều nhân quả ACEs → SKTT.')
add_red(doc, '6. Tạp chí J Affective Disorders Reports (không phải J Affective Disorders Q1) \u2014 tạp chí mở mới, chưa có impact factor ổn định.')

add_red(doc, 'Khoảng trống (Gap):', bold=True)
add_red(doc, '1. Mở rộng đến nhiều tỉnh vùng cao (Hà Giang, Lai Châu, Sơn La, Điện Biên).')
add_red(doc, '2. So sánh DTTS nội trú vs DTTS không nội trú vs Kinh cùng vùng.')
add_red(doc, '3. Nghiên cứu dọc: ACEs ảnh hưởng SKTT theo thời gian ở DTTS.')
add_red(doc, '4. Can thiệp cải thiện mối quan hệ bạn bè tại trường nội trú.')
add_red(doc, '5. Phân tích giới tính chi tiết ở VTN DTTS \u2014 có khác biệt văn hóa giới ở vùng cao?')

fname = '15_NgoAnhVinh_2024_JAffectDisordReports.docx'
doc.save(fname)
sys.stderr.write(f'{fname} saved OK\n')
