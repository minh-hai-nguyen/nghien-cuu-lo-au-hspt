# -*- coding: utf-8 -*-
"""Bài 07: Zhu 2025 - BMC Public Health - China Middle School Depression"""
import sys, io, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()
add_link_and_qr(doc, 'https://doi.org/10.1186/s12889-024-21252-8', 'QR_Hoa2024.png')

add_heading(doc, 'Các yếu tố ảnh hưởng và xu hướng thay đổi triệu chứng trầm cảm ở học sinh THCS và THPT tại miền Đông Trung Quốc từ 2019 đến 2023', 1)

add_info_table(doc, [
    ('Tiêu đề gốc', 'Influencing factors and changing trends of depressive symptoms among middle and junior high school students in Eastern China from 2019 to 2023'),
    ('Tác giả', 'Zhu và cộng sự'),
    ('Tạp chí', 'BMC Public Health, 2025, 25:17'),
    ('DOI', '10.1186/s12889-024-21252-8'),
    ('Mẫu', 'n = 9.831 học sinh THCS và THPT tại miền Đông Trung Quốc (Suzhou)'),
    ('Công cụ', 'PHQ-9 (Patient Health Questionnaire-9 — Bảng hỏi Sức khỏe Bệnh nhân 9 mục)'),
    ('Thiết kế', 'Nghiên cứu cắt ngang, khảo sát trực tuyến 2019-2023'),
])

doc.add_paragraph()
add_page_ref(doc, '1/11', 'BMC Public Health', '25:17, 2025')

add_heading(doc, 'TÓM TẮT', 2)
add_p(doc, 'Bối cảnh: Đại dịch COVID-19 (Coronavirus Disease 2019) dẫn đến gia tăng mức độ trầm cảm ở vị thành niên (VTN). Tuy nhiên, tình hình trầm cảm VTN trong giai đoạn sau đại dịch chưa được nghiên cứu đầy đủ. Nghiên cứu này phân tích các yếu tố ảnh hưởng và xu hướng thay đổi triệu chứng trầm cảm ở học sinh trung học cơ sở (THCS) và trung học phổ thông (THPT) tại miền Đông Trung Quốc từ 2019 đến 2023.')
add_p(doc, 'Phương pháp: Khảo sát cắt ngang trên 9.831 học sinh THCS và THPT tại Suzhou, Trung Quốc. Sử dụng PHQ-9 (Patient Health Questionnaire-9 — Bảng hỏi Sức khỏe Bệnh nhân 9 mục) để đánh giá triệu chứng trầm cảm. Phân tích hồi quy logistic đa biến xác định yếu tố nguy cơ và bảo vệ.')
add_p(doc, 'Kết quả: 14,5% học sinh có triệu chứng trầm cảm "có thể" (possible) và 5,8% "chắc chắn" (definite). Yếu tố nguy cơ: HS THPT (AOR = 1,409), gia đình đơn thân (AOR = 1,434), gia đình tái hôn (AOR = 1,837), ngủ < 5 giờ (AOR = 13,710 cho trầm cảm chắc chắn). Yếu tố bảo vệ: nam giới (AOR = 0,803), hoạt động ngoài trời > 1 giờ/ngày (AOR = 0,666\u20130,785).')

add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(doc, 'Điểm nổi bật:', bold=True)
add_p(doc, '\u2022 Cỡ mẫu lớn (n = 9.831), so sánh xu hướng 2019\u20132023 \u2014 bao gồm giai đoạn trước, trong và sau COVID (Racine và cộng sự, 2021).', size=11)
add_p(doc, '\u2022 Phát hiện đáng chú ý: ngủ < 5 giờ tăng nguy cơ trầm cảm 13,7 lần (AOR = 13,710) \u2014 yếu tố nguy cơ mạnh nhất (Hale và cộng sự, 2015).', size=11)
add_p(doc, '\u2022 Hoạt động ngoài trời > 1 giờ/ngày là yếu tố bảo vệ rõ ràng \u2014 có thể ứng dụng can thiệp (WHO, 2022).', size=11)

add_p(doc, 'Hạn chế:', bold=True)
add_p(doc, '\u2022 Chỉ tại Suzhou (miền Đông, đô thị phát triển) \u2014 không đại diện nông thôn hay miền Tây (Samuels và cộng sự, 2018).', size=11)
add_p(doc, '\u2022 PHQ-9 là sàng lọc, không phải chẩn đoán (COVID-19 Mental Disorders Collaborators, 2021).', size=11)
add_p(doc, '\u2022 Chỉ đo trầm cảm, thiếu lo âu (Spitzer và cộng sự, 2006).', size=11)

add_p(doc, 'Hướng cải thiện:', bold=True)
add_p(doc, '\u2022 Bổ sung GAD-7 đo lo âu song song (Spitzer và cộng sự, 2006).', size=11)
add_p(doc, '\u2022 Mở rộng ra nông thôn miền Tây Trung Quốc để so sánh (Wen và cộng sự, 2020).', size=11)
add_p(doc, '\u2022 RCT can thiệp tăng giấc ngủ + hoạt động ngoài trời tại trường học (Werner-Seidler và cộng sự, 2021).', size=11)

doc.add_paragraph()

add_page_ref(doc, '2\u20133', 'BMC Public Health', '25:17')
add_heading(doc, 'GIỚI THIỆU', 2)
add_p(doc, 'Khoảng 1/7 VTN 10\u201319 tuổi trên toàn thế giới mắc rối loạn tâm thần, chiếm 13% tổng gánh nặng bệnh tật ở nhóm tuổi này (WHO, 2022). Tại Trung Quốc, một khảo sát cho thấy khoảng 14,8% VTN có nguy cơ trầm cảm (Li và cộng sự, 2022). Với sự phát triển kinh tế xã hội nhanh chóng, tỷ lệ vấn đề SKTT ở trẻ em và VTN Trung Quốc đang gia tăng.')

add_page_ref(doc, '3\u20134', 'BMC Public Health', '25:17')
add_heading(doc, 'PHƯƠNG PHÁP', 2)
add_p(doc, 'Khảo sát cắt ngang tại Suzhou, tỉnh Giang Tô, miền Đông Trung Quốc. Mẫu: 9.921 học sinh (sau loại trừ: 9.831, tỷ lệ phản hồi 99,1%). Công cụ: PHQ-9 \u2014 9 mục, điểm 0\u201327. Ngưỡng: 0\u20134 (không), 5\u20139 (nhẹ/possible), 10\u201327 (trung bình\u2013nặng/definite). Phân tích: hồi quy logistic đa biến, AOR (Adjusted Odds Ratio) với 95% CI.')

add_page_ref(doc, '5\u20137', 'BMC Public Health', '25:17')
add_heading(doc, 'KẾT QUẢ', 2)

add_p(doc, 'Tỷ lệ trầm cảm:', bold=True)
add_table(doc,
    ['Mức độ', 'Tỷ lệ'],
    [
        ['Không có triệu chứng (PHQ-9: 0\u20134)', '79,7%'],
        ['Có thể trầm cảm (PHQ-9: 5\u20139)', '14,5%'],
        ['Chắc chắn trầm cảm (PHQ-9: 10\u201327)', '5,8%'],
    ], widths=[6.0, 4.0])

doc.add_paragraph()
add_p(doc, 'Yếu tố nguy cơ và bảo vệ (hồi quy logistic):', bold=True)
add_table(doc,
    ['Yếu tố', 'AOR', '95% CI', 'Loại'],
    [
        ['HS THPT (vs THCS)', '1,409', '1,164\u20131,706', 'Nguy cơ'],
        ['Gia đình đơn thân', '1,434', '1,134\u20131,814', 'Nguy cơ'],
        ['Gia đình tái hôn', '1,837', '1,324\u20132,548', 'Nguy cơ'],
        ['Cả 2 cha mẹ vắng mặt', '1,710', '1,289\u20132,268', 'Nguy cơ'],
        ['Ngủ < 5 giờ', '13,710', '8,659\u201321,710', 'Nguy cơ (rất mạnh)'],
        ['Ngủ 5\u20136 giờ', '7,289', '5,315\u20139,996', 'Nguy cơ'],
        ['Ngủ 6\u20137 giờ', '2,678', '2,020\u20133,551', 'Nguy cơ'],
        ['Nam giới (vs nữ)', '0,803', '0,716\u20130,902', 'Bảo vệ'],
        ['Hoạt động ngoài trời 1\u20132 giờ', '0,656', '0,534\u20130,807', 'Bảo vệ'],
        ['Hoạt động ngoài trời 2\u20133 giờ', '0,557', '0,409\u20130,759', 'Bảo vệ'],
    ], widths=[4.0, 2.0, 3.0, 3.0])

doc.add_paragraph()
add_p(doc, 'Xu hướng 2019\u20132023: Tỷ lệ trầm cảm ở Nantong (thành phố lân cận) là 20,3% \u2014 cao hơn đáng kể so với Suzhou, gợi ý sự khác biệt giữa các thành phố ngay trong cùng khu vực.')

add_page_ref(doc, '8\u20139', 'BMC Public Health', '25:17')
add_heading(doc, 'THẢO LUẬN', 2)
add_p(doc, 'Ngủ ít là yếu tố nguy cơ mạnh nhất (AOR = 13,71 cho < 5 giờ) \u2014 phù hợp với meta-analysis cho thấy thiếu ngủ tăng gấp đôi nguy cơ trầm cảm ở VTN (Lovato & Gradisar, 2014). Hoạt động ngoài trời là yếu tố bảo vệ rõ ràng \u2014 nhất quán với bằng chứng về tác dụng của thể chất lên SKTT (Biddle và cộng sự, 2019).')
add_p(doc, 'Cấu trúc gia đình ảnh hưởng đáng kể: gia đình tái hôn (AOR = 1,84) và cả 2 cha mẹ vắng mặt (AOR = 1,71) tăng nguy cơ. Đây là vấn đề phổ biến ở Trung Quốc với hiện tượng "trẻ em bị bỏ lại" (left-behind children) khi cha mẹ đi làm xa.')

add_heading(doc, 'KẾT LUẬN', 2)
add_p(doc, 'Khoảng 1/5 học sinh THCS/THPT tại miền Đông Trung Quốc có triệu chứng trầm cảm. Giấc ngủ đủ và hoạt động ngoài trời là hai yếu tố bảo vệ có thể can thiệp được. Cần chương trình SKTT tại trường học tích hợp giáo dục giấc ngủ và hoạt động thể chất.')

add_abbreviation_table(doc, [
    ('PHQ-9', 'Patient Health Questionnaire-9 \u2014 Bảng hỏi Sức khỏe Bệnh nhân 9 mục'),
    ('AOR', 'Adjusted Odds Ratio \u2014 Tỷ số chênh điều chỉnh'),
    ('CI', 'Confidence Interval \u2014 Khoảng tin cậy'),
    ('THCS', 'Trung học cơ sở (Junior high school)'),
    ('THPT', 'Trung học phổ thông (Senior high school)'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('VTN', 'Vị thành niên'),
])

doc.add_paragraph()
add_red_heading(doc, 'PHẢN BIỆN CHI TIẾT')
add_red(doc, 'Điểm mạnh:', bold=True)
add_red(doc, '\u2022 Cỡ mẫu lớn (n = 9.831) với tỷ lệ phản hồi cực cao (99,1%) \u2014 giảm thiên lệch chọn (Etikan và cộng sự, 2016).')
add_red(doc, '\u2022 PHQ-9 được sử dụng rộng rãi và có tính so sánh quốc tế cao (Spitzer và cộng sự, 2006).')
add_red(doc, '\u2022 Phân tích đa biến kiểm soát nhiều biến gây nhiễu (VanderWeele, 2019).')
add_red(doc, '\u2022 Phát hiện AOR = 13,71 cho ngủ < 5 giờ \u2014 bằng chứng mạnh cho can thiệp giấc ngủ (Hale và cộng sự, 2015).')

add_red(doc, 'Hạn chế:', bold=True)
add_red(doc, '\u2022 Chỉ tại Suzhou \u2014 thành phố phát triển, không đại diện nông thôn hay miền Tây nghèo hơn (GBD 2021 ASEAN Collaborators, 2025).')
add_red(doc, '\u2022 Thiết kế cắt ngang: không thể kết luận ngủ ít GÂY RA trầm cảm hay trầm cảm gây mất ngủ (Sedgwick, 2014).')
add_red(doc, '\u2022 PHQ-9 là sàng lọc, không phải chẩn đoán \u2014 20,3% là triệu chứng, V-NAMHS dùng DISC-5 chỉ ghi nhận 2,3% rối loạn lo âu tại Việt Nam (Vũ Mạnh Lợi và cộng sự, 2022).')
add_red(doc, '\u2022 Chỉ đo trầm cảm, thiếu lo âu \u2014 cần bổ sung GAD-7 (Spitzer và cộng sự, 2006).')
add_red(doc, '\u2022 Khảo sát trực tuyến \u2014 có thể thiên lệch cho HS có thiết bị/internet (Etikan và cộng sự, 2016).')

add_red(doc, 'Khoảng trống nghiên cứu (Gap):', bold=True)
add_red(doc, '\u2022 RCT can thiệp tăng giấc ngủ + hoạt động ngoài trời tại trường THCS/THPT \u2014 đặc biệt cho nhóm ngủ < 6 giờ (Werner-Seidler và cộng sự, 2021).')
add_red(doc, '\u2022 Nghiên cứu dọc theo dõi cùng nhóm HS từ 2019 qua COVID đến 2025 (Caruana và cộng sự, 2015).')
add_red(doc, '\u2022 So sánh miền Đông (Suzhou) vs miền Tây (nông thôn) Trung Quốc \u2014 yếu tố kinh tế xã hội (Wen và cộng sự, 2020).')
add_red(doc, '\u2022 Nghiên cứu cơ chế: giấc ngủ \u2192 viêm thần kinh \u2192 trầm cảm (Hale và cộng sự, 2015; Chronic Stress Neuroinflammation, 2023).')
add_red(doc, '\u2022 So sánh liên quốc gia: Trung Quốc vs Việt Nam vs Hàn Quốc với cùng PHQ-9 (GBD 2021 ASEAN Collaborators, 2025).')

doc.save('07_Zhu_2025_BMC_China.docx')
