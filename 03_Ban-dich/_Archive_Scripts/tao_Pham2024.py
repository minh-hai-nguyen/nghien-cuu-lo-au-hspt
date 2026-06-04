# -*- coding: utf-8 -*-
import sys, io, os
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()
add_link_and_qr(doc, 'https://doi.org/10.1177/09731342241275742', 'QR_Hoa2024.png')
add_heading(doc, 'Mối tương quan giữa chất lượng chăm sóc và các vấn đề sức khỏe tâm thần, hành vi ở vị thành niên Việt Nam tại các cơ sở bảo trợ xã hội', 1)

add_info_table(doc, [
    ('Tiêu đề gốc', 'The Correlation Between Quality of Care and Mental Health and Behavioral Problems Among Vietnamese Adolescents in Social Support Facilities'),
    ('Tác giả', 'Sy Tien Pham, Thanh Thanh Thi Duong, Hoai Phuong Thi Nguyen, Xuan Nhi Thi Truong'),
    ('Cơ quan', 'Đại học Huế, Việt Nam'),
    ('Tạp chí', 'Journal of Indian Association for Child and Adolescent Mental Health, 2024, Vol. 20(4), 373-381'),
    ('DOI', '10.1177/09731342241275742'),
    ('Mẫu', '273 VTN tại cơ sở bảo trợ xã hội (SSF) + 273 VTN sống với gia đình'),
    ('Công cụ', 'CES-D (trầm cảm, \u03b1=0,87), Thang lo âu VN (\u03b1=0,79), SDQ (hành vi), Thang chất lượng chăm sóc (\u03b1=0,88/0,83)'),
    ('Thiết kế', 'Nghiên cứu cắt ngang, so sánh 2 nhóm'),
])
doc.add_paragraph()

add_page_ref(doc, '373', 'J Indian Assoc Child Adolesc Mental Health', 'Vol. 20(4), 2024')
add_heading(doc, 'TÓM TẮT', 2)
add_p(doc, 'Bối cảnh: Các nghiên cứu trước đây cho thấy tác động tiêu cực của chăm sóc tại cơ sở (institutional care) đối với sức khỏe thể chất và tâm thần của trẻ em. Nghiên cứu này khám phá chất lượng chăm sóc và mối liên hệ với các vấn đề sức khỏe tâm thần (SKTT) ở vị thành niên (VTN) tại các cơ sở bảo trợ xã hội (SSF \u2014 Social Support Facilities) tại Việt Nam.')
add_p(doc, 'Phương pháp: Nghiên cứu cắt ngang trên 273 VTN sống tại SSF và 273 VTN được gia đình chăm sóc. Đo lường triệu chứng trầm cảm bằng Thang Trầm cảm CES-D (Center for Epidemiologic Studies Depression Scale), lo âu bằng thang đo phát triển cho VTN Việt Nam, rối loạn giấc ngủ và vấn đề hành vi bằng SDQ (Strengths and Difficulties Questionnaire).')
add_p(doc, 'Kết quả: Chăm sóc thể chất không có mối liên hệ có ý nghĩa với nguy cơ SKTT (trừ lo âu). Chăm sóc cảm xúc có tương quan nghịch với tất cả các vấn đề SKTT: lo âu (\u03b2 = \u22120,40, p < 0,001), trầm cảm (\u03b2 = \u22120,31, p < 0,01), vấn đề hành vi (\u03b2 = \u22120,30, p < 0,01) và rối loạn giấc ngủ (\u03b2 = \u22120,30, p < 0,01).')
add_p(doc, 'Kết luận: Chăm sóc cảm xúc quan trọng hơn chăm sóc thể chất. Cần cải thiện chất lượng chăm sóc cảm xúc để giảm nguy cơ SKTT.')

add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(doc, 'Điểm nổi bật:', bold=True)
add_p(doc, '\u2022 Nghiên cứu đầu tiên tại Việt Nam về mối quan hệ chất lượng chăm sóc \u2014 SKTT ở VTN cơ sở bảo trợ (Pham và cộng sự, 2024).', size=11)
add_p(doc, '\u2022 Có nhóm đối chứng 273 VTN gia đình \u2014 hiếm trong nghiên cứu VN (Humphreys và cộng sự, 2015).', size=11)
add_p(doc, '\u2022 Phát hiện: chăm sóc cảm xúc quan trọng hơn thể chất (\u03b2 = \u22120,40 vs không có ý nghĩa).', size=11)
add_p(doc, '\u2022 Sử dụng thang đo lo âu phát triển cho VTN Việt Nam (Nguyễn, Lê & Dunne, 2007).', size=11)

add_p(doc, 'Hạn chế:', bold=True)
add_p(doc, '\u2022 Thiết kế cắt ngang \u2014 không xác định nhân quả (Sedgwick, 2014).', size=11)
add_p(doc, '\u2022 Chỉ tại Huế \u2014 không đại diện SSF toàn quốc (Samuels và cộng sự, 2018).', size=11)
add_p(doc, '\u2022 Tạp chí IF thấp. R\u00b2 điều chỉnh thấp (2,6\u201310,2%).', size=11)

add_p(doc, 'Hướng cải thiện:', bold=True)
add_p(doc, '\u2022 Nghiên cứu dọc theo dõi SKTT VTN SSF qua nhiều năm (Caruana và cộng sự, 2015).', size=11)
add_p(doc, '\u2022 RCT can thiệp nâng cao chăm sóc cảm xúc \u2192 đo SKTT (Werner-Seidler và cộng sự, 2021).', size=11)
add_p(doc, '\u2022 Mở rộng ra nhiều tỉnh, kết hợp đánh giá ACEs (Felitti và cộng sự, 1998).', size=11)

doc.add_paragraph()
add_page_ref(doc, '374', 'J Indian Assoc Child Adolesc Mental Health', 'Vol. 20(4)')
add_heading(doc, 'GIỚI THIỆU', 2)
add_p(doc, 'Trẻ em sống trong cơ sở chăm sóc tập trung đối mặt với nhiều vấn đề SKTT: trầm cảm, lo âu, ý tưởng tự tử, lạm dụng chất, bạo lực (Humphreys và cộng sự, 2015; Zeanah và cộng sự, 2009). Nghiên cứu tại Romania cho thấy trẻ được chuyển từ cơ sở sang gia đình đỡ đầu có giảm đáng kể các vấn đề SKTT (Humphreys và cộng sự, 2015).')
add_p(doc, 'Tại Việt Nam, hơn 80% VTN sống tại SSF từng trải qua ít nhất một hình thức lạm dụng thể chất, cảm xúc hoặc tình dục (Pham và cộng sự, 2021). Nghiên cứu tại Romania cho thấy 39,5% trẻ tại cơ sở bị nhân viên đánh nghiêm trọng (Rus và cộng sự).')
add_p(doc, 'Huynh và cộng sự cho thấy chất lượng chăm sóc (an ninh lương thực, chỗ ở, y tế) dự báo sức khỏe tâm lý xã hội ở trẻ tại các nước thu nhập thấp và trung bình. Tuy nhiên, chưa có nghiên cứu nào tại Việt Nam phân tích riêng vai trò của chăm sóc cảm xúc so với chăm sóc thể chất.')

add_page_ref(doc, '375', 'J Indian Assoc Child Adolesc Mental Health', 'Vol. 20(4)')
add_heading(doc, 'PHƯƠNG PHÁP', 2)
add_p(doc, 'Nghiên cứu cắt ngang tại 3 SSF và 3 trường THCS/THPT ở Huế. Nhóm IA: 273 VTN 12\u201318 tuổi sống tại SSF. Nhóm NIA: 273 VTN cùng độ tuổi sống với gia đình. Đạo đức nghiên cứu được phê duyệt bởi Đại học Huế (Dự án DHH2023-01-207).')
add_p(doc, 'Công cụ đo lường:', bold=True)
add_p(doc, '\u2022 CES-D: 20 mục, điểm 0\u201360, phiên bản tiếng Việt (Radloff, 1977; Cronbach \u03b1 = 0,87).')
add_p(doc, '\u2022 Thang lo âu VTN Việt Nam: 13 mục, Likert 3 điểm, điểm 13\u201339 (Nguyễn, Lê & Dunne, 2007; \u03b1 = 0,79).')
add_p(doc, '\u2022 SDQ phân mục hành vi: 5 mục (Dang và cộng sự, 2017).')
add_p(doc, '\u2022 Thang chất lượng chăm sóc: 8 mục \u2014 thể chất (4 mục, \u03b1 = 0,884) và cảm xúc (4 mục, \u03b1 = 0,831).')

add_page_ref(doc, '376\u2013377', 'J Indian Assoc Child Adolesc Mental Health', 'Vol. 20(4)')
add_heading(doc, 'KẾT QUẢ', 2)
add_p(doc, 'So sánh hai nhóm (Bảng 3):', bold=True)
add_table(doc,
    ['Chỉ số', 'Nhóm SSF (IA)', 'Nhóm gia đình (NIA)', 't', 'p'],
    [
        ['Chăm sóc thể chất', '16,61 (4,04)', '18,05 (2,89)', '\u22124,77', '<0,001'],
        ['Chăm sóc cảm xúc', '15,42 (4,08)', '17,10 (3,37)', '\u22125,26', '<0,001'],
        ['Trầm cảm (CES-D)', '15,75 (8,47)', '15,70 (9,15)', '0,07', '0,95'],
        ['Lo âu', '22,96 (4,83)', '21,89 (4,39)', '2,71', '0,007'],
        ['Vấn đề hành vi', '2,32 (1,57)', '1,92 (1,29)', '3,26', '0,001'],
        ['Rối loạn giấc ngủ', '5,11 (3,76)', '4,82 (3,45)', '0,93', '>0,05'],
    ], widths=[3.0, 2.5, 2.8, 1.5, 1.5])

doc.add_paragraph()
add_p(doc, 'Hồi quy tuyến tính đa biến \u2014 Chăm sóc cảm xúc vs SKTT (Bảng 5):', bold=True)
add_table(doc,
    ['Biến phụ thuộc', '\u03b2 (chăm sóc cảm xúc)', 'p', 'R\u00b2 điều chỉnh'],
    [
        ['Lo âu', '\u22120,40', '<0,001', '0,102'],
        ['Trầm cảm', '\u22120,31', '<0,01', '0,080'],
        ['Vấn đề hành vi', '\u22120,30', '<0,01', '0,026'],
        ['Rối loạn giấc ngủ', '\u22120,30', '<0,01', '0,054'],
    ], widths=[3.5, 3.5, 2.0, 2.5])

doc.add_paragraph()
add_p(doc, 'Phát hiện then chốt: Chăm sóc thể chất KHÔNG có mối liên hệ có ý nghĩa với SKTT (trừ lo âu, \u03b2 = 0,28, p < 0,05). Ngược lại, chăm sóc cảm xúc có tương quan nghịch mạnh với TẤT CẢ các vấn đề SKTT.')

add_page_ref(doc, '378\u2013379', 'J Indian Assoc Child Adolesc Mental Health', 'Vol. 20(4)')
add_heading(doc, 'THẢO LUẬN', 2)
add_p(doc, 'Kết quả trái ngược với nghiên cứu 5 nước thu nhập thấp (Huynh và cộng sự) cho thấy TẤT CẢ thành phần chất lượng chăm sóc đều dự báo sức khỏe tâm lý. Sự khác biệt có thể do bối cảnh Việt Nam \u2014 khi nhu cầu thể chất cơ bản đã được đáp ứng, yếu tố cảm xúc trở thành quyết định.')
add_p(doc, 'Kết quả nhất quán với lý thuyết gắn bó (attachment theory) \u2014 mối quan hệ cảm xúc an toàn với người chăm sóc là nền tảng SKTT.')

add_heading(doc, 'KẾT LUẬN', 2)
add_p(doc, 'Chăm sóc cảm xúc quan trọng hơn chăm sóc thể chất đối với SKTT VTN tại SSF. Cần cải thiện chất lượng chăm sóc cảm xúc \u2014 tình yêu thương, khích lệ, bảo vệ \u2014 để giảm nguy cơ SKTT cho VTN tại các cơ sở bảo trợ xã hội tại Việt Nam.')

add_abbreviation_table(doc, [
    ('SSF', 'Social Support Facilities \u2014 Cơ sở bảo trợ xã hội'),
    ('IA', 'Institutionalized Adolescents \u2014 VTN sống tại cơ sở'),
    ('NIA', 'Non-Institutionalized Adolescents \u2014 VTN sống với gia đình'),
    ('CES-D', 'Center for Epidemiologic Studies Depression Scale'),
    ('SDQ', 'Strengths and Difficulties Questionnaire'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('VTN', 'Vị thành niên'),
    ('ACEs', 'Adverse Childhood Experiences \u2014 Trải nghiệm bất lợi thời thơ ấu'),
])

doc.add_paragraph()
add_red_heading(doc, 'PHẢN BIỆN CHI TIẾT')
add_red(doc, 'Điểm mạnh:', bold=True)
add_red(doc, '\u2022 Nghiên cứu đầu tiên tại Việt Nam phân biệt vai trò chăm sóc thể chất vs cảm xúc đối với SKTT VTN tại SSF (Pham và cộng sự, 2024).')
add_red(doc, '\u2022 Có nhóm đối chứng (NIA) cùng độ tuổi cho phép so sánh trực tiếp \u2014 thiết kế mạnh hơn nhiều nghiên cứu VN chỉ có 1 nhóm (Humphreys và cộng sự, 2015).')
add_red(doc, '\u2022 Phát hiện có ý nghĩa lâm sàng: chăm sóc cảm xúc (\u03b2 = \u22120,40) có tác động mạnh hơn nhiều so với chăm sóc thể chất (không có ý nghĩa) \u2014 định hướng can thiệp rõ ràng.')
add_red(doc, '\u2022 Sử dụng thang đo lo âu phát triển riêng cho VTN Việt Nam, phù hợp văn hóa (Nguyễn, Lê & Dunne, 2007).')

add_red(doc, 'Hạn chế:', bold=True)
add_red(doc, '\u2022 Thiết kế cắt ngang không cho phép kết luận nhân quả: chưa rõ chăm sóc cảm xúc tốt HƠN CÓ TÁC DỤNG giảm SKTT hay VTN ít vấn đề SKTT đánh giá chăm sóc tốt hơn (Sedgwick, 2014).')
add_red(doc, '\u2022 Chỉ tại Huế \u2014 SSF ở TP.HCM, Hà Nội, vùng miền núi có thể rất khác về chất lượng chăm sóc và đặc điểm VTN (Samuels và cộng sự, 2018).')
add_red(doc, '\u2022 R\u00b2 điều chỉnh rất thấp (2,6\u201310,2%) \u2014 mô hình giải thích ít phương sai, nhiều yếu tố quan trọng chưa được đưa vào (VanderWeele, 2019).')
add_red(doc, '\u2022 Không đánh giá ACEs (trải nghiệm bất lợi) dù 80% VTN SSF Việt Nam từng bị lạm dụng \u2014 đây là biến gây nhiễu quan trọng (Felitti và cộng sự, 1998).')
add_red(doc, '\u2022 CES-D đo triệu chứng 1 tuần trước, không phải chẩn đoán lâm sàng \u2014 tỷ lệ thực có thể khác (Radloff, 1977; V-NAMHS: Vũ Mạnh Lợi và cộng sự, 2022).')
add_red(doc, '\u2022 Tạp chí IF thấp (J Indian Assoc Child Adolesc Mental Health) \u2014 hạn chế tầm ảnh hưởng.')

add_red(doc, 'Khoảng trống nghiên cứu (Gap):', bold=True)
add_red(doc, '\u2022 RCT can thiệp: đào tạo nhân viên SSF về kỹ năng chăm sóc cảm xúc \u2192 đo SKTT VTN trước-sau 6 tháng. Chưa có RCT nào tại SSF Việt Nam (Werner-Seidler và cộng sự, 2021; Zhameden và cộng sự, 2025).')
add_red(doc, '\u2022 Nghiên cứu dọc: theo dõi VTN từ khi nhập SSF đến khi rời \u2014 SKTT thay đổi như thế nào theo thời gian? Vai trò của thời gian lưu trú? (Caruana và cộng sự, 2015).')
add_red(doc, '\u2022 So sánh đa trung tâm: SSF công lập vs tư nhân vs tôn giáo (chùa/nhà thờ) tại nhiều tỉnh \u2014 hiện chỉ có dữ liệu từ Huế (WHO, 2022).')
add_red(doc, '\u2022 Kết hợp ACEs + chất lượng chăm sóc + hỗ trợ xã hội trong mô hình SEM (Structural Equation Modeling) \u2014 xác định cơ chế trung gian (Felitti và cộng sự, 1998; Compas và cộng sự, 2017).')
add_red(doc, '\u2022 Đánh giá chi phí-hiệu quả: đầu tư vào chăm sóc cảm xúc tiết kiệm bao nhiêu so với điều trị SKTT sau này? (Chisholm và cộng sự, 2016; UNICEF, 2021).')

doc.save('10_Pham_2024_VN_SocialSupport.docx')
