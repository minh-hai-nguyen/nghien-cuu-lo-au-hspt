# -*- coding: utf-8 -*-
"""Dịch chi tiết Hoàng Trung Học 2025 — format Jenkins
8.473 VTN, 6 tỉnh VN, DASS-21, COVID vs hậu COVID"""
import sys, io, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()

# 1. LINK + QR
add_link_and_qr(doc, 'https://doi.org/10.69980/ajpr.v28i1.105', 'QR_HoangTrungHoc2025.png')

# 2. TIÊU ĐỀ
add_heading(doc, 'Mức độ căng thẳng, lo âu và trầm cảm ở thanh thiếu niên trong và sau đại dịch COVID-19 tại Việt Nam: Nghiên cứu cắt ngang', 1)
add_heading(doc, 'Levels of Stress, Anxiety, and Depression in Adolescents during and after the COVID-19 Pandemic in Vietnam: A Cross-sectional Study', 2)

# 3. THÔNG TIN THƯ MỤC
add_heading(doc, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(doc, [
    ('Tiêu đề gốc', 'Levels of Stress, Anxiety, and Depression in Adolescents during and after the COVID-19 Pandemic in Vietnam: A Cross-sectional Study'),
    ('Tiêu đề dịch', 'Mức độ căng thẳng, lo âu và trầm cảm ở thanh thiếu niên trong và sau đại dịch COVID-19 tại Việt Nam: Nghiên cứu cắt ngang'),
    ('Tác giả', 'Hoàng Trung Học (PhD, Học viện Quản lý Giáo dục Quốc gia)\nNguyễn Thùy Dung (ThS, Viện Nghiên cứu, Đào tạo và Ứng dụng Tâm lý)'),
    ('Tạp chí', 'American Journal of Psychiatric Rehabilitation, Vol. 28, No. 1, April 2025, pp. 360\u2013367'),
    ('DOI', '10.69980/ajpr.v28i1.105'),
    ('ISSN', '1548-7776'),
    ('Cỡ mẫu', '8.473 thanh thiếu niên lớp 6\u201312, 6 tỉnh/thành phố Việt Nam'),
    ('Thời điểm', 'Đợt 1: 12/2021 (trong COVID), Đợt 2: 9/2023 (sau COVID)'),
])

# TÓM TẮT ĐÁNH GIÁ NHANH
add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(doc, 'Điểm nổi bật:', bold=True)
add_p(doc, '\u2022 Cỡ mẫu lớn nhất trong các nghiên cứu SKTT VTN Việt Nam: 8.473 HS, 6 tỉnh (3 miền Nam + 3 miền Bắc).')
add_p(doc, '\u2022 So sánh TRONG vs SAU COVID-19: lo âu giảm từ 41,5% xuống 25,4% — bằng chứng phục hồi.')
add_p(doc, '\u2022 Xác định 3 yếu tố quan trọng nhất: quan hệ cha mẹ-con, thời gian dùng điện tử, giấc ngủ.')
add_p(doc, '\u2022 Hồi quy đa biến: "quan hệ cha mẹ" có Beta = 0,272 mạnh nhất trong COVID, "thời gian dùng điện tử" Beta = 0,176.')

add_p(doc, 'Hạn chế:', bold=True)
add_p(doc, '\u2022 Tạp chí AJPR không có impact factor rõ ràng, không lập chỉ mục PubMed/Scopus.')
add_p(doc, '\u2022 Thiết kế cắt ngang 2 thời điểm trên 2 mẫu khác nhau — không phải nghiên cứu dọc thực sự.')
add_p(doc, '\u2022 DASS-21 là sàng lọc, không phải chẩn đoán. Tỷ lệ cao (41,5\u201365,5%) có thể bao gồm dương tính giả.')
add_p(doc, '\u2022 Lấy mẫu thuận tiện ngẫu nhiên (convenient random sampling) — hạn chế tính đại diện.')

add_p(doc, 'Hướng cải thiện:', bold=True)
add_p(doc, '\u2022 Nghiên cứu dọc thực sự theo dõi cùng nhóm HS. Xuất bản trên tạp chí có impact factor. So sánh DASS-21 với DISC-5.')

# TÓM TẮT
add_page_ref(doc, '360', 'Am J Psychiatric Rehabilitation', 'Vol. 28(1), 2025')
add_heading(doc, 'TÓM TẮT', 2)
add_p(doc, 'Bối cảnh: Căng thẳng, lo âu và trầm cảm là những mối quan ngại quan trọng về sức khỏe tâm thần ảnh hưởng đáng kể đến sự phát triển tâm lý của thanh thiếu niên, đòi hỏi nghiên cứu có hệ thống và khoa học để cung cấp thông tin cho các biện pháp phòng ngừa và can thiệp hiệu quả, đặc biệt trong các cuộc khủng hoảng xã hội.')
add_red(doc, 'Phương pháp: Sử dụng phiên bản tiếng Việt của thang DASS-21 để đánh giá 8.473 thanh thiếu niên tại 6 tỉnh Việt Nam (3 miền Nam: TPHCM, Đồng Nai, Long An; 3 miền Bắc: Hà Nội, Hải Dương, Vĩnh Phúc). Đợt 1: 12/2021 (trong COVID, n=4.052), Đợt 2: 9/2023 (sau COVID, n=4.337). Hồi quy đa biến xác định yếu tố ảnh hưởng.')
add_red(doc, 'Kết quả: Tỷ lệ căng thẳng, lo âu và trầm cảm TRONG COVID: 65,5%, 41,5% và 34,2%. SAU COVID giảm còn: 55,4%, 25,4% và 20,1%. Điểm DASS-21 trung bình giảm từ 26,68 xuống 22,07 (p < 0,01). Yếu tố ảnh hưởng mạnh nhất: quan hệ cha mẹ-con (Beta = 0,272), thời gian dùng điện tử (Beta = 0,176), giấc ngủ (Beta = -0,149).')
add_p(doc, 'Kết luận: Cần điều chỉnh thời gian sử dụng công nghệ và tăng cường sự gắn kết gia đình như thành phần thiết yếu của chiến lược phòng ngừa và can thiệp sớm cho SKTT thanh thiếu niên.')

# NỘI DUNG DỊCH
add_heading(doc, 'GIỚI THIỆU', 2)
add_page_ref(doc, '360\u2013361', 'Am J Psychiatric Rehabilitation', 'Vol. 28(1), 2025')
add_p(doc, 'Tuổi vị thành niên là giai đoạn phát triển quan trọng với những thay đổi đáng kể về thể chất, xã hội và tâm lý, làm tăng tính dễ bị tổn thương trước các yếu tố gây căng thẳng. Tại Ấn Độ, Sandal và cộng sự (2017) báo cáo tỷ lệ căng thẳng, lo âu và trầm cảm ở thanh thiếu niên lần lượt là 65,53%, 80,85% và 47,02%. Tại Việt Nam, Nguyen và cộng sự (2024) phát hiện 34\u201350% thanh thiếu niên trải qua căng thẳng, lo âu và trầm cảm \u2014 tỷ lệ cao hơn so với thời kỳ trước đại dịch.')
add_p(doc, 'Hai giả thuyết nghiên cứu: H1: Mức độ căng thẳng, lo âu và trầm cảm trong COVID cao hơn sau COVID. H2: Quan hệ cha mẹ-con và thời gian dùng điện tử có tác động rõ ràng cả trong và sau COVID.')

add_heading(doc, 'PHƯƠNG PHÁP', 2)
add_page_ref(doc, '362\u2013363', 'Am J Psychiatric Rehabilitation', 'Vol. 28(1), 2025')
add_p(doc, 'Mẫu nghiên cứu:', bold=True)
add_p(doc, '8.473 thanh thiếu niên lớp 6\u201312, chọn ngẫu nhiên thuận tiện từ các trường THCS và THPT tại 6 tỉnh/thành phố: miền Nam (TPHCM, Đồng Nai, Long An) và miền Bắc (Hà Nội, Hải Dương, Vĩnh Phúc). Nông thôn chiếm 62,8% (5.269/8.389), thành thị 37,2% (3.120/8.389).')

add_table(doc,
    ['Khối lớp', 'Nông thôn n (%)', 'Thành thị n (%)', 'Tổng n (%)'],
    [['Lớp 6', '1.098 (20,8%)', '212 (6,8%)', '1.310 (15,6%)'],
     ['Lớp 7', '1.155 (21,9%)', '187 (6,0%)', '1.342 (16,0%)'],
     ['Lớp 8', '1.095 (20,8%)', '150 (4,8%)', '1.245 (14,8%)'],
     ['Lớp 9', '1.063 (20,2%)', '263 (8,4%)', '1.326 (15,8%)'],
     ['Lớp 10', '290 (5,5%)', '943 (30,2%)', '1.233 (14,7%)'],
     ['Lớp 11', '240 (4,6%)', '687 (22,0%)', '927 (11,1%)'],
     ['Lớp 12', '328 (6,2%)', '678 (21,7%)', '1.006 (12,0%)'],
     ['Tổng', '5.269 (100%)', '3.120 (100%)', '8.389 (100%)']],
    widths=[2.5, 3.5, 3.5, 3.0])
doc.add_paragraph()

add_p(doc, 'Công cụ:', bold=True)
add_p(doc, 'Thang đo Trầm cảm Lo âu Căng thẳng 21 mục (DASS-21 \u2014 Depression Anxiety Stress Scale; Lovibond & Lovibond, 1995), phiên bản tiếng Việt. Cronbach alpha > 0,7 cho tất cả thang con. Bổ sung bảng hỏi về giấc ngủ, sử dụng thiết bị điện tử, hoạt động thể chất và quan hệ gia đình.')

add_p(doc, 'Quy trình:', bold=True)
add_p(doc, 'Thiết kế cắt ngang tại 2 thời điểm: Đợt 1 \u2014 tháng 12/2021 (trong COVID, n = 4.052); Đợt 2 \u2014 tháng 9/2023 (sau COVID, n = 4.337). Phân tích: SPSS, thống kê mô tả, tương quan Pearson, hồi quy đa biến. DW = 0,353 và 0,340; VIF < 2 \u2014 loại trừ tự tương quan và đa cộng tuyến.')

add_heading(doc, 'KẾT QUẢ', 2)
add_page_ref(doc, '363\u2013365', 'Am J Psychiatric Rehabilitation', 'Vol. 28(1), 2025')

add_p(doc, 'Tỷ lệ căng thẳng, lo âu và trầm cảm (DASS-21, N = 8.389):', bold=True)
add_table(doc,
    ['Tình trạng', 'Trong COVID (%)', 'Sau COVID (%)', 'Thay đổi'],
    [['Căng thẳng tổng', '65,5%', '55,4%', '\u221910,1 điểm'],
     ['  Nặng + Rất nặng', '33,0%', '18,6%', '\u221914,4 điểm'],
     ['Lo âu tổng', '41,5%', '25,4%', '\u221916,1 điểm'],
     ['  Nặng + Rất nặng', '14,5%', '7,2%', '\u22197,3 điểm'],
     ['Trầm cảm tổng', '34,2%', '20,1%', '\u221914,1 điểm'],
     ['  Nặng + Rất nặng', '8,1%', '4,0%', '\u22194,1 điểm'],
     ['Điểm DASS-21 TB', '26,68', '22,07', '\u22194,61 (p<0,01)']],
    widths=[4.0, 3.0, 3.0, 2.5])
doc.add_paragraph()

add_p(doc, 'Tương quan: Căng thẳng, lo âu và trầm cảm có tương quan dương, chặt và có ý nghĩa (0,592 < r < 0,592, p < 0,01).', italic=True)
doc.add_paragraph()

add_p(doc, 'Yếu tố ảnh hưởng \u2014 Hồi quy đa biến (TRONG COVID, R\u00b2 = 0,190):', bold=True)
add_table(doc,
    ['Biến', 'Beta', 'T', 'Sig.', 'VIF'],
    [['Quan hệ cha mẹ-con', '0,272', '18,997', '0,000', '1,023'],
     ['Thời gian dùng điện tử', '0,176', '10,744', '0,000', '1,342'],
     ['Giấc ngủ ban ngày', '\u20130,149', '\u201310,378', '0,000', '1,032'],
     ['Thời gian thể thao', '\u20130,087', '\u20135,547', '0,000', '1,217'],
     ['Tuổi', '\u20130,086', '\u20135,200', '0,000', '1,377'],
     ['Hình thức học tập', '0,074', '4,524', '0,000', '1,334'],
     ['Mức giãn cách xã hội', '0,073', '4,941', '0,000', '1,094'],
     ['Giới tính', '0,053', '3,455', '0,001', '1,181'],
     ['Khu vực cư trú', '0,042', '2,472', '0,013', '1,475']],
    widths=[4.0, 1.5, 2.0, 1.5, 1.5])
doc.add_paragraph()

add_p(doc, 'Sau COVID: Yếu tố mạnh nhất vẫn là quan hệ cha mẹ-con và thời gian dùng điện tử, nhưng hoạt động ngoại khóa và lối sống tích cực đóng vai trò nổi bật hơn.', italic=True)

add_heading(doc, 'THẢO LUẬN', 2)
add_page_ref(doc, '365\u2013367', 'Am J Psychiatric Rehabilitation', 'Vol. 28(1), 2025')
add_p(doc, 'Trước COVID, tỷ lệ lo âu ở VTN Việt Nam dao động 12\u201315%, trầm cảm 10\u201315%. Trong COVID, tỷ lệ tăng mạnh (lo âu 37\u201347%, trầm cảm 35\u201343%). Nghiên cứu này xác nhận xu hướng đó với lo âu 41,5% trong COVID.')
add_p(doc, 'Sau COVID, tỷ lệ giảm (lo âu 25,4%) nhưng vẫn cao hơn trước dịch \u2014 gợi ý tác động kéo dài. Cần tiếp tục hỗ trợ tâm lý để VTN phục hồi hoàn toàn.')
add_p(doc, 'Ba yếu tố mạnh nhất: (1) Quan hệ cha mẹ-con (Beta = 0,272) \u2014 gia đình hòa thuận bảo vệ SKTT, phù hợp với Pham 2024 (chăm sóc cảm xúc beta = \u20130,40) và Qiu 2022 (nuôi dạy tích cực OR = 0,30). (2) Thời gian dùng điện tử (Beta = 0,176) \u2014 phù hợp Chen 2023 (game OR = 5,00). (3) Giấc ngủ (Beta = \u20130,149) \u2014 phù hợp Zhu 2025 (ngủ <5h AOR = 13,71).')

add_heading(doc, 'KẾT LUẬN', 2)
add_p(doc, 'Mức độ căng thẳng, lo âu và trầm cảm ở VTN tăng đáng kể trong COVID-19, giảm sau đại dịch nhưng vẫn cao hơn trước dịch. Cải thiện quan hệ gia đình, quản lý thời gian dùng thiết bị công nghệ và duy trì chất lượng giấc ngủ là thiết yếu cho phòng ngừa và can thiệp sớm.')

# TÀI LIỆU THAM KHẢO
add_heading(doc, 'TÀI LIỆU THAM KHẢO', 2)
add_p(doc, '[Giữ nguyên 30 TLTK tiếng Anh từ bài gốc \u2014 xem PDF trang 366\u2013367]', italic=True)

# BẢNG VIẾT TẮT
add_abbreviation_table(doc, [
    ('DASS-21', 'Depression Anxiety Stress Scale 21-item \u2014 Thang đo Trầm cảm Lo âu Căng thẳng 21 mục'),
    ('COVID-19', 'Coronavirus Disease 2019 \u2014 Bệnh virus corona 2019'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('VTN', 'Vị thành niên / Thanh thiếu niên'),
    ('THCS', 'Trung học cơ sở'),
    ('THPT', 'Trung học phổ thông'),
    ('VIF', 'Variance Inflation Factor \u2014 Hệ số phóng đại phương sai'),
    ('DW', 'Durbin-Watson \u2014 Kiểm định tự tương quan'),
    ('AJPR', 'American Journal of Psychiatric Rehabilitation'),
])

# PHẢN BIỆN
add_red_heading(doc, 'PHẢN BIỆN CHI TIẾT')

add_red(doc, 'Điểm mạnh:', bold=True)
add_red(doc, '1. Cỡ mẫu lớn nhất trong các NC SKTT VTN Việt Nam: 8.473 HS từ 6 tỉnh (3 miền Nam + 3 miền Bắc) \u2014 tính đại diện tương đối tốt.')
add_red(doc, '2. So sánh 2 thời điểm (trong vs sau COVID) \u2014 ít NC nào thực hiện so sánh này tại Việt Nam.')
add_red(doc, '3. Hồi quy đa biến với kiểm soát VIF, DW \u2014 phương pháp phân tích đáng tin cậy.')
add_red(doc, '4. Xác định 3 yếu tố mạnh nhất có thể can thiệp được (gia đình, điện tử, giấc ngủ).')

add_red(doc, 'Hạn chế:', bold=True)
add_red(doc, '1. Tạp chí AJPR không có impact factor rõ ràng, không lập chỉ mục PubMed/Scopus \u2014 giảm độ tin cậy học thuật.')
add_red(doc, '2. Hai đợt khảo sát trên 2 MẪU KHÁC NHAU (4.052 vs 4.337) \u2014 không phải nghiên cứu dọc thực sự, không thể theo dõi thay đổi ở cùng cá nhân.')
add_red(doc, '3. DASS-21 là sàng lọc, tỷ lệ 41,5\u201365,5% bao gồm cả mức nhẹ \u2014 so với V-NAMHS 2022 (DISC-5 chẩn đoán chỉ 2,3%), chênh lệch rất lớn.')
add_red(doc, '4. Lấy mẫu "convenient random sampling" \u2014 mâu thuẫn (thuận tiện \u2260 ngẫu nhiên), hạn chế tính đại diện.')
add_red(doc, '5. R\u00b2 = 0,190 \u2014 mô hình chỉ giải thích 19% phương sai, 81% còn lại do yếu tố chưa đo lường.')
add_red(doc, '6. Bảng 3 trong bài gốc ghi nhãn "Reckless" cho mức "Không lo âu" \u2014 có thể lỗi dịch từ tiếng Việt sang tiếng Anh.')
add_red(doc, '7. Không phân tích riêng giới tính trong tỷ lệ lo âu \u2014 chỉ có Beta = 0,053 cho giới tính trong hồi quy.')

add_red(doc, 'Khoảng trống nghiên cứu (Gap):', bold=True)
add_red(doc, '1. Nghiên cứu dọc thực sự theo dõi cùng nhóm HS từ trong COVID đến 2\u20133 năm sau.')
add_red(doc, '2. So sánh DASS-21 với DISC-5/DSM-5 trên cùng mẫu VTN Việt Nam.')
add_red(doc, '3. Phân tích riêng giới tính, vùng miền (Bắc vs Nam), nông thôn vs thành thị.')
add_red(doc, '4. Đánh giá hiệu quả can thiệp: đào tạo kỹ năng cha mẹ, quản lý thời gian điện tử tại trường.')
add_red(doc, '5. Xuất bản kết quả trên tạp chí có chỉ mục quốc tế (PubMed/Scopus) để tăng khả năng so sánh.')

fname = '14_HoangTrungHoc_2025_AJPR.docx'
doc.save(fname)
sys.stderr.write(f'{fname} saved OK\n')
