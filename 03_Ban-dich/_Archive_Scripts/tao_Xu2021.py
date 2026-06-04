# -*- coding: utf-8 -*-
"""Dịch chi tiết Xu et al. 2021 — format Jenkins
CỠ MẪU LỚN NHẤT TOÀN CẦU — 373.216 HS — NAM > NỮ lo âu"""
import sys, io, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()

# 1. LINK + QR
add_link_and_qr(doc, 'https://doi.org/10.1016/j.jad.2021.03.080', 'QR_Xu2021.png')

# 2. TIÊU ĐỀ
add_heading(doc, 'Tỷ lệ và các yếu tố nguy cơ đối với triệu chứng lo âu trong thời kỳ bùng phát COVID-19: Khảo sát quy mô lớn trên 373.216 học sinh trung học tại Trung Quốc', 1)
add_heading(doc, 'Prevalence and Risk Factors for Anxiety Symptoms during the Outbreak of COVID-19: A Large Survey among 373,216 Junior and Senior High School Students in China', 2)

# 3. THÔNG TIN THƯ MỤC
add_heading(doc, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(doc, [
    ('Tiêu đề gốc', 'Prevalence and risk factors for anxiety symptoms during the outbreak of COVID-19: A large survey among 373,216 junior and senior high school students in China'),
    ('Tiêu đề dịch', 'Tỷ lệ và các yếu tố nguy cơ đối với triệu chứng lo âu trong thời kỳ bùng phát COVID-19: Khảo sát quy mô lớn trên 373.216 học sinh trung học cơ sở và trung học phổ thông tại Trung Quốc'),
    ('Tác giả', 'Qingqing Xu, Zhenxing Mao, Dandan Wei, Pengling Liu, Keliang Fan, Juan Wang, Xian Wang, Xiaomin Lou, Hualiang Lin, Chongjian Wang, Cuiping Wu'),
    ('Cơ quan', 'Khoa Dịch tễ học và Thống kê Sinh học, Trường Y tế Công cộng, Đại học Trịnh Châu, Hà Nam; Đại học Tôn Trung Sơn, Quảng Châu, Trung Quốc'),
    ('Tạp chí', 'Journal of Affective Disorders, Volume 288, 1 June 2021, Pages 17\u201322 (Q1, IF = 6.6)'),
    ('DOI', '10.1016/j.jad.2021.03.080'),
    ('PMID', '33839554 | PMCID: PMC9754660'),
    ('Thu thập dữ liệu', '4\u201312/02/2020 (đỉnh dịch COVID-19 tại Trung Quốc)'),
])

# TÓM TẮT ĐÁNH GIÁ NHANH
add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(doc, 'Điểm nổi bật:', bold=True)
add_p(doc, '\u2022 CỠ MẪU LỚN NHẤT TOÀN CẦU về lo âu thanh thiếu niên: 373.216 học sinh (244.193 THCS + 129.023 THPT).')
add_p(doc, '\u2022 NAM > NỮ: nam 10,11% > nữ 9,66% (P < 0,001) \u2014 một trong 2 nghiên cứu duy nhất (cùng Saikia 2023) cho thấy nam lo âu cao hơn nữ.')
add_p(doc, '\u2022 Nông thôn > Thành thị: 11,33% vs 8,77% (OR = 1,30) \u2014 bằng chứng mạnh cho bất bình đẳng SKTT theo vùng.')
add_p(doc, '\u2022 Tạp chí Q1: Journal of Affective Disorders (IF = 6,6) \u2014 chất lượng xuất bản cao.')

add_p(doc, 'Hạn chế:', bold=True)
add_p(doc, '\u2022 Khảo sát trực tuyến trong 8 ngày (4\u201312/02/2020) \u2014 thiên lệch tự chọn (self-selection bias).')
add_p(doc, '\u2022 Đỉnh dịch COVID-19 \u2014 kết quả phản ánh bối cảnh khủng hoảng, không thể khái quát cho thời bình.')
add_p(doc, '\u2022 Chỉ dùng GAD-7 sàng lọc, không có chẩn đoán lâm sàng.')
add_p(doc, '\u2022 Kiến thức COVID-19 đánh giá bằng 3 câu hỏi chưa được xác thực.')
add_p(doc, '\u2022 Chênh lệch nam-nữ rất nhỏ (10,11% vs 9,66% = 0,45 điểm %) \u2014 có ý nghĩa thống kê nhờ cỡ mẫu cực lớn.')

add_p(doc, 'Hướng cải thiện:', bold=True)
add_p(doc, '\u2022 So sánh dọc đỉnh dịch vs hậu COVID. Can thiệp giảm sợ hãi không hợp lý. Mở rộng nhiều tỉnh.')

# 4. TÓM TẮT
add_page_ref(doc, '17', 'J Affective Disorders', 'Vol. 288, 2021')
add_heading(doc, 'TÓM TẮT', 2)
add_p(doc, 'Bối cảnh: Mối đe dọa ngày càng tăng của đại dịch COVID-19 đã dẫn đến bầu không khí lo âu trên toàn thế giới, tuy nhiên bằng chứng ở học sinh Trung Quốc từ 12 đến 18 tuổi vẫn còn hạn chế.')
add_red(doc, 'Phương pháp: Tổng cộng 373.216 học sinh THCS và THPT được tuyển chọn bằng phương pháp lấy mẫu cụm tại Trịnh Châu, Tân Hương, Tín Dương thuộc tỉnh Hà Nam, Trung Quốc, từ 4 đến 12/02/2020. Triệu chứng lo âu được xác định bằng GAD-7. Hồi quy logistic đa biến được thực hiện.')
add_red(doc, 'Kết quả: Tỷ lệ lo âu tổng thể 9,89%. Nam > nữ (10,11% vs 9,66%). THCS > THPT (10,85% vs 8,08%). Nông thôn > thành thị (11,33% vs 8,77%). Mức nhận thức COVID-19 tương quan nghịch với lo âu. Sau điều chỉnh đa biến: tuổi, giới tính, vị trí cư trú, mức lo lắng, mức sợ hãi và hành vi phòng ngừa liên quan với lo âu.')
add_p(doc, 'Kết luận: Cần quan tâm SKTT người trẻ trong đại dịch. Ưu tiên can thiệp cho học sinh nông thôn, THCS.')

# NỘI DUNG DỊCH ĐẦY ĐỦ
add_heading(doc, '1. GIỚI THIỆU', 2)
add_page_ref(doc, '17\u201318', 'J Affective Disorders', 'Vol. 288, 2021')
add_p(doc, 'Vào tháng 12/2019, vi-rút corona mới (COVID-19) được xác định, gây lo ngại toàn cầu. Nhiều biện pháp phòng chống (giãn cách, cách ly, phong tỏa) được thực hiện tại Trung Quốc trong Tết Nguyên đán. Các biện pháp này có thể giảm lây lan nhưng cũng gây tác động tiêu cực đến kinh tế, việc làm và sức khỏe cộng đồng.')
add_p(doc, 'Rối loạn lo âu là một trong những tình trạng SKTT phổ biến và suy giảm chức năng nhiều nhất toàn cầu. Tỷ lệ ước tính khoảng 7,3% toàn cầu. Tác động của COVID-19 đến SKTT có thể khác nhau giữa các nhóm dân số, nhưng kiến thức về thanh thiếu niên vẫn rất hạn chế.')
add_p(doc, 'Tỉnh Hà Nam có 28,53 triệu người trong hệ thống giáo dục \u2014 tỉnh có dân số giáo dục lớn nhất Trung Quốc, giáp Hồ Bắc (tâm dịch). Nghiên cứu này nhằm đánh giá tỷ lệ lo âu và yếu tố nguy cơ ở 373.216 HS trong đỉnh dịch COVID-19.')

add_heading(doc, '2. PHƯƠNG PHÁP', 2)
add_page_ref(doc, '18\u201319', 'J Affective Disorders', 'Vol. 288, 2021')
add_p(doc, '2.1. Đối tượng nghiên cứu', bold=True)
add_p(doc, 'Nghiên cứu cắt ngang từ 4\u201312/02/2020. HS THCS và THPT 12\u201318 tuổi tại 3 thành phố (Trịnh Châu, Tân Hương, Tín Dương) tỉnh Hà Nam. Lấy mẫu cụm. Khảo sát trực tuyến qua nền tảng "SurveyStar". Sau loại trừ (<12 tuổi, >18 tuổi, hoàn thành <100 giây): n = 34.276 bị loại, còn 373.216.')

add_p(doc, '2.2. Thu thập dữ liệu', bold=True)
add_p(doc, 'Bảng câu hỏi chuẩn: thông tin nhân khẩu (giới tính, tuổi, lớp, vị trí cư trú), nhận thức COVID-19 (3 câu: lây người-người, đường lây, cách ly), mức lo lắng (Likert 5 điểm: cao 4\u20135, TB 3, thấp/không 1\u20132), mức sợ hãi (tương tự), hành vi phòng ngừa (4 câu: rửa tay, khẩu trang, bỏ du lịch, ăn ngoài \u2192 tất cả đúng/không tất cả/tất cả sai).')
add_p(doc, 'Lo âu: GAD-7 phiên bản tiếng Trung (Spitzer et al., 2006). Ngưỡng: nhẹ 5\u20139, trung bình 10\u201314, nặng 15\u201321. Ngưỡng cắt \u226510 cho dương tính lo âu.')

add_p(doc, '2.3. Phân tích thống kê', bold=True)
add_p(doc, 'Chi-square, t-test. Hồi quy logistic: Mô hình 1 (thô), Mô hình 2 (điều chỉnh tuổi, giới, cư trú, lo lắng, sợ hãi, hành vi). SPSS 21.0, P < 0,05 hai phía.')

add_heading(doc, '3. KẾT QUẢ', 2)
add_page_ref(doc, '19\u201320', 'J Affective Disorders', 'Vol. 288, 2021')

add_p(doc, '3.1. Đặc điểm mẫu (N = 373.216)', bold=True)
add_table(doc,
    ['Đặc điểm', 'Tổng (n=373.216)', 'Không lo âu (n=336.298)', 'Lo âu (n=36.918)', 'P'],
    [['Tuổi (TB\u00b1SD)', '15,24\u00b11,59', '15,26\u00b11,60', '15,06\u00b11,58', '0,009'],
     ['Nam', '193.507 (51,85%)', '173.943 (51,72%)', '19.564 (52,99%)', '<0,001'],
     ['Nữ', '179.709 (48,15%)', '162.355 (48,28%)', '17.354 (47,01%)', ''],
     ['THCS', '244.193 (65,43%)', '217.703 (64,74%)', '26.490 (71,75%)', '<0,001'],
     ['THPT', '129.023 (34,57%)', '118.595 (35,26%)', '10.428 (28,25%)', '']],
    widths=[3.5, 3.0, 3.0, 3.0, 1.5])
doc.add_paragraph()

add_p(doc, '3.2. Tỷ lệ lo âu', bold=True)
add_table(doc,
    ['Nhóm', 'Tỷ lệ lo âu (GAD-7 \u226510)', 'Ghi chú'],
    [['Tổng thể', '9,89% (36.918/373.216)', '\u2014'],
     ['THCS', '10,85%', 'Cao hơn'],
     ['THPT', '8,08%', 'Thấp hơn'],
     ['Nam', '10,11%', '\u26a0\ufe0f Cao hơn nữ'],
     ['Nữ', '9,66%', '\u2014'],
     ['Nông thôn (THCS)', '12,80%', 'Cao nhất'],
     ['Thành phố (THCS)', '9,30%', 'Thấp nhất'],
     ['Nông thôn (THPT)', '8,40%', '\u2014'],
     ['Thành phố (THPT)', '7,80%', 'Thấp nhất THPT'],
     ['Nông thôn chung', '11,33%', '\u2014'],
     ['Thành phố chung', '8,77%', '\u2014']],
    widths=[4.0, 4.0, 3.5])
doc.add_paragraph()

add_p(doc, '\u26a0\ufe0f PHÁT HIỆN ĐẶC BIỆT: Nam 10,11% > Nữ 9,66% (P < 0,001) \u2014 TRÁI NGƯỢC y văn quốc tế. Chênh lệch nhỏ (0,45 điểm %) nhưng có ý nghĩa thống kê nhờ cỡ mẫu cực lớn (373.216). Đây là 1 trong 2 nghiên cứu (cùng Saikia 2023) cho thấy nam > nữ.', bold=True)

add_p(doc, '3.3. Mức độ lo âu', bold=True)
add_table(doc,
    ['Mức độ', 'THCS', 'THPT'],
    [['Nhẹ (GAD 5\u20139)', '28,06%', '29,43%'],
     ['Trung bình (GAD 10\u201314)', '7,35%', '5,72%'],
     ['Nặng (GAD 15\u201321)', '3,50%', '2,36%']],
    widths=[4.0, 3.5, 3.5])
doc.add_paragraph()

add_p(doc, 'Ba triệu chứng phổ biến nhất trong nhóm lo âu: bồn chồn khó ngồi yên (41,2%), dễ cáu gắt (49,1%), sợ điều xấu xảy ra (39,4%).', italic=True)

add_page_ref(doc, '20\u201321', 'J Affective Disorders', 'Vol. 288, 2021')
add_p(doc, '3.4. Yếu tố nguy cơ và bảo vệ (hồi quy logistic đa biến)', bold=True)
add_table(doc,
    ['Yếu tố', 'THCS OR (KTC 95%)', 'THPT OR (KTC 95%)'],
    [['Nữ (vs nam)', '0,92 (0,89\u20130,94) BV', '0,84 (0,81\u20130,88) BV'],
     ['Nông thôn (vs thành phố)', '1,30 (1,26\u20131,34)', '1,24 (1,17\u20131,32)'],
     ['TP cấp huyện (vs thành phố)', '1,11 (1,07\u20131,15)', '1,10 (1,04\u20131,16)'],
     ['Lo lắng TB (vs cao)', '0,60 (0,56\u20130,64) BV', '0,66 (0,61\u20130,73) BV'],
     ['Lo lắng thấp (vs cao)', '0,61 (0,55\u20130,68) BV', '0,78 (0,67\u20130,92) BV'],
     ['Sợ hãi TB (vs cao)', '0,21 (0,20\u20130,22) BV', '0,20 (0,19\u20130,21) BV'],
     ['Sợ hãi thấp (vs cao)', '0,21 (0,20\u20130,23) BV', '0,19 (0,17\u20130,21) BV'],
     ['Hành vi TB (vs tốt)', '1,04 (1,01\u20131,07)', '1,06 (1,01\u20131,10)'],
     ['Hành vi kém (vs tốt)', '2,72 (2,01\u20133,68)', '2,93 (1,97\u20134,35)']],
    widths=[4.5, 4.0, 4.0])
add_p(doc, 'BV = Bảo vệ (giảm nguy cơ lo âu)', italic=True, size=10)
doc.add_paragraph()

add_heading(doc, '4. THẢO LUẬN', 2)
add_page_ref(doc, '21\u201322', 'J Affective Disorders', 'Vol. 288, 2021')
add_p(doc, 'Tỷ lệ lo âu tổng thể 9,89% cao hơn tỷ lệ toàn cầu ước tính 7,3%, có thể do biện pháp phòng chống dịch. Đầu đợt bùng phát, khoảng 1/3 dân số Trung Quốc báo cáo lo âu trung bình\u2013nặng.')
add_p(doc, 'Học sinh nông thôn có tỷ lệ lo âu cao nhất \u2014 có thể do tình trạng kinh tế xã hội thấp hơn của cha mẹ, ảnh hưởng đến SKTT con cái. THCS có tỷ lệ cao hơn THPT \u2014 cần ưu tiên can thiệp.')
add_p(doc, 'Kiến thức COVID-19 thấp liên quan lo âu cao hơn \u2014 gợi ý tăng cường tuyên truyền về đường lây và cách phòng ngừa.')
add_p(doc, 'Sợ hãi quá mức (OR sợ hãi cao gấp ~5 lần so với sợ hãi thấp) là yếu tố nguy cơ mạnh \u2014 can thiệp giảm sợ hãi không hợp lý là mục tiêu quan trọng.')
add_p(doc, 'Hạn chế: Thiết kế cắt ngang, tự báo cáo, không chẩn đoán lâm sàng. Khảo sát trực tuyến 8 ngày có thể có thiên lệch. Kết quả phản ánh bối cảnh khủng hoảng cấp tính.')

add_heading(doc, '5. KẾT LUẬN', 2)
add_p(doc, 'Nghiên cứu quy mô lớn nhất toàn cầu (373.216 HS) đánh giá lo âu thanh thiếu niên trong đỉnh dịch COVID-19. Tỷ lệ 9,89% với nam > nữ, THCS > THPT, nông thôn > thành thị. Cần ưu tiên SKTT học sinh nông thôn, THCS, và tăng cường kiến thức COVID-19 để giảm sợ hãi không hợp lý.')

# TÀI LIỆU THAM KHẢO
add_heading(doc, 'TÀI LIỆU THAM KHẢO', 2)
add_p(doc, '[Giữ nguyên danh mục TLTK tiếng Anh từ bài gốc \u2014 xem PDF gốc trang 22]', italic=True)

# BẢNG VIẾT TẮT
add_abbreviation_table(doc, [
    ('COVID-19', 'Coronavirus Disease 2019 \u2014 Bệnh virus corona 2019'),
    ('GAD-7', 'Generalized Anxiety Disorder 7-item Scale \u2014 Thang đo Rối loạn Lo âu Tổng quát 7 mục'),
    ('OR', 'Odds Ratio \u2014 Tỷ số chênh'),
    ('CI/KTC', 'Confidence Interval \u2014 Khoảng tin cậy'),
    ('THCS', 'Trung học cơ sở (Junior High School)'),
    ('THPT', 'Trung học phổ thông (Senior High School)'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('SPSS', 'Statistical Package for Social Sciences'),
    ('BV', 'Bảo vệ (Protective factor)'),
    ('SARS-CoV-2', 'Severe Acute Respiratory Syndrome Coronavirus 2'),
])

# PHẢN BIỆN
add_red_heading(doc, 'PHẢN BIỆN CHI TIẾT')

add_red(doc, 'Điểm mạnh:', bold=True)
add_red(doc, '1. Cỡ mẫu LỚN NHẤT TOÀN CẦU: 373.216 học sinh \u2014 sức mạnh thống kê cực cao, phát hiện chênh lệch rất nhỏ cũng có ý nghĩa.')
add_red(doc, '2. Tạp chí Q1 (J Affective Disorders, IF = 6,6) \u2014 qua peer-review chặt chẽ.')
add_red(doc, '3. GAD-7 đã xác thực quốc tế (Spitzer et al., 2006), ngưỡng \u226510 là bảo thủ.')
add_red(doc, '4. Hồi quy logistic đa biến \u2014 kiểm soát nhiều biến gây nhiễu.')
add_red(doc, '5. Đa thành phố (3 TP đại diện tỉnh Hà Nam) \u2014 không chỉ 1 trường/thành phố.')

add_red(doc, 'Hạn chế:', bold=True)
add_red(doc, '1. Khảo sát trực tuyến 8 ngày (4\u201312/02/2020) \u2014 thiên lệch tự chọn. Gia đình không có internet bị loại.')
add_red(doc, '2. Đỉnh dịch COVID-19 \u2014 kết quả không thể khái quát cho thời bình. Lo âu có thể phản ánh phản ứng cấp tính hơn rối loạn mạn tính.')
add_red(doc, '3. Chỉ dùng GAD-7 sàng lọc, không chẩn đoán lâm sàng \u2014 tỷ lệ có thể bao gồm dương tính giả.')
add_red(doc, '4. Nam > nữ chênh lệch rất nhỏ (0,45 điểm %) \u2014 có ý nghĩa thống kê nhờ n cực lớn nhưng ý nghĩa lâm sàng thấp.')
add_red(doc, '5. Kiến thức COVID-19 đánh giá bằng 3 câu hỏi tự thiết kế \u2014 chưa được xác thực (validation).')
add_red(doc, '6. Loại 34.276 người tham gia hoàn thành <100 giây \u2014 tỷ lệ loại 8,4% có thể gây thiên lệch.')
add_red(doc, '7. Abstract ghi THCS 13,89% vs THPT 12,93% nhưng Results ghi THCS 10,85% vs THPT 8,08% \u2014 mâu thuẫn nội bộ trong bài báo gốc (trang 17 vs trang 21).')

add_red(doc, 'Khoảng trống nghiên cứu (Gap):', bold=True)
add_red(doc, '1. So sánh dọc lo âu đỉnh dịch (02/2020) vs hậu COVID (2022\u20132024) trên cùng quần thể.')
add_red(doc, '2. Nghiên cứu can thiệp giảm sợ hãi không hợp lý ở thanh thiếu niên trong khủng hoảng sức khỏe.')
add_red(doc, '3. Giải thích tại sao nam > nữ trong bối cảnh COVID \u2014 yếu tố văn hóa hay phản ứng khủng hoảng?')
add_red(doc, '4. Mở rộng đến nhiều tỉnh (không chỉ Hà Nam) để đánh giá tính khái quát.')
add_red(doc, '5. So sánh GAD-7 với công cụ chẩn đoán trên cùng mẫu \u2014 ước tính khoảng cách sàng lọc\u2013chẩn đoán trong COVID.')

fname = '13_Xu_2021_JAffectDisord.docx'
doc.save(fname)
sys.stderr.write(f'{fname} saved OK\n')
