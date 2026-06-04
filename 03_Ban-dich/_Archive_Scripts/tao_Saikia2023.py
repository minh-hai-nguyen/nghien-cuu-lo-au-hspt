# -*- coding: utf-8 -*-
"""Dịch chi tiết Saikia et al. 2023 — format Jenkins
Lý do: Nam > Nữ về lo âu (trái y văn quốc tế)"""
import sys, io, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()

# 1. LINK + QR
add_link_and_qr(doc, 'https://doi.org/10.4103/ijcm.ijcm_1103_22', 'QR_Saikia2023.png')

# 2. TIÊU ĐỀ
add_heading(doc, 'Các bệnh lý sức khỏe tâm thần và các yếu tố liên quan ở thanh thiếu niên tại Kamrup (Metro), Assam: Một nghiên cứu tại trường học', 1)
add_heading(doc, 'Mental Health Morbidities and their Correlates among the Adolescents in Kamrup (Metro), Assam: A School-Based Study', 2)

# 3. THÔNG TIN THƯ MỤC
add_heading(doc, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(doc, [
    ('Tiêu đề gốc', 'Mental Health Morbidities and their Correlates among the Adolescents in Kamrup (Metro), Assam: A School-Based Study'),
    ('Tiêu đề dịch', 'Các bệnh lý sức khỏe tâm thần và các yếu tố liên quan ở thanh thiếu niên tại Kamrup (Metro), Assam: Một nghiên cứu tại trường học'),
    ('Tác giả', 'Anku M. Saikia, Hemen Das, Vinoth Rajendran'),
    ('Cơ quan', 'Khoa Y học Cộng đồng, Trường Y khoa và Bệnh viện Gauhati, Guwahati, Assam, Ấn Độ'),
    ('Tạp chí', 'Indian Journal of Community Medicine, 2023, Vol. 48(5), 733–738'),
    ('DOI', '10.4103/ijcm.ijcm_1103_22'),
    ('PMID', '37970140'),
])

# 4. TÓM TẮT
add_page_ref(doc, '733', 'Indian J Community Med', 'Vol. 48(5), 2023')

add_heading(doc, 'TÓM TẮT', 2)
add_p(doc, 'Bối cảnh: Thanh thiếu niên trải qua nhiều thay đổi về phát triển, cảm xúc, tâm lý và xã hội. Ở Ấn Độ, nhóm tuổi 10–19 chiếm khoảng 21,4% dân số. Dữ liệu về sức khỏe tâm thần thanh thiếu niên tại vùng Đông Bắc Ấn Độ còn rất hạn chế.')
add_p(doc, 'Mục tiêu: Xác định tỷ lệ hiện mắc của trầm cảm, lo âu và căng thẳng cùng các yếu tố liên quan ở thanh thiếu niên tại quận Kamrup (Metro), Assam.')

add_red(doc, 'Phương pháp: Nghiên cứu cắt ngang tại trường học được thực hiện từ tháng 4/2019 đến tháng 6/2020 trên 360 thanh thiếu niên 14–17 tuổi. Từ 120 trường trung học công lập, 10 trường được chọn ngẫu nhiên đơn giản, mỗi trường chọn 36 học sinh (12/lớp 8, 9, 10), đảm bảo nam nữ bằng nhau. Công cụ: DASS-21 phiên bản tiếng Assam + thang đo Kuppuswamy sửa đổi.')
add_red(doc, 'Kết quả: Tỷ lệ trầm cảm 22,2% (80/360), lo âu 24,4% (88/360), căng thẳng 6,9% (25/360). NAM có tỷ lệ lo âu CAO HƠN NỮ (30,0% so với 18,9%, P = 0,049) — TRÁI NGƯỢC y văn quốc tế. Cha mẹ đơn thân, sử dụng rượu của cha mẹ, chơi trò chơi điện tử, lưu ban có liên quan có ý nghĩa (P < 0,05).')

add_p(doc, 'Kết luận: Sức khỏe tâm thần thanh thiếu niên tại Đông Bắc Ấn Độ cần được chú ý. Phát hiện nam > nữ về lo âu gợi ý yếu tố văn hóa bộ lạc đặc thù khu vực có thể đảo ngược xu hướng giới tính thông thường.')

# TÓM TẮT ĐÁNH GIÁ NHANH
add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(doc, 'Điểm nổi bật:', bold=True)
add_p(doc, '• Nghiên cứu tiên phong tại vùng Đông Bắc Ấn Độ — khu vực có đặc điểm văn hóa bộ lạc khác biệt, hầu như chưa được khảo sát về SKTT thanh thiếu niên.')
add_p(doc, '• Phát hiện nam > nữ về lo âu (30,0% vs 18,9%) — trái ngược hoàn toàn với y văn quốc tế, chỉ phù hợp với 2 nghiên cứu khác (Wen 2020, Xu 2021).')
add_p(doc, '• Thiết kế lấy mẫu ngẫu nhiên tốt từ 10/120 trường, đảm bảo nam nữ bằng nhau.')

add_p(doc, 'Hạn chế:', bold=True)
add_p(doc, '• Chỉ sử dụng chi-square (đơn biến), không có hồi quy logistic đa biến — không thể đánh giá ảnh hưởng độc lập.')
add_p(doc, '• DASS-21 tiếng Assam thiếu báo cáo Cronbach alpha — không rõ độ tin cậy phiên bản dịch.')
add_p(doc, '• Thu thập 4/2019–6/2020 trùng COVID-19 nhưng không kiểm soát hiệu ứng đại dịch.')
add_p(doc, '• Tác giả không giải thích sâu cơ chế văn hóa cụ thể khiến nam > nữ.')

add_p(doc, 'Hướng cải thiện:', bold=True)
add_p(doc, '• Mở rộng cỡ mẫu với hồi quy đa biến. Khảo sát sâu yếu tố văn hóa bộ lạc. Xác thực DASS-21 tiếng Assam.')

# NỘI DUNG DỊCH ĐẦY ĐỦ
add_heading(doc, 'GIỚI THIỆU', 2)
add_page_ref(doc, '733', 'Indian J Community Med', 'Vol. 48(5), 2023')
add_p(doc, 'Thanh thiếu niên là giai đoạn chuyển tiếp từ thời thơ ấu sang tuổi trưởng thành, được đánh dấu bởi nhiều thay đổi về phát triển, cảm xúc, tâm lý và xã hội. Nhóm tuổi 10–19 tại Ấn Độ chiếm khoảng 21,4% dân số (khoảng 253 triệu người). Sức khỏe tâm thần trong giai đoạn này có ý nghĩa quan trọng cho sự phát triển lành mạnh. Trên toàn cầu, ước tính khoảng 10–20% trẻ em và thanh thiếu niên trải qua các rối loạn tâm thần.')
add_p(doc, 'Tại Ấn Độ, phần lớn các nghiên cứu về sức khỏe tâm thần thanh thiếu niên tập trung vào các bang miền Nam và miền Bắc. Vùng Đông Bắc Ấn Độ — nơi có đặc điểm văn hóa bộ lạc khác biệt và các yếu tố kinh tế xã hội đặc thù — hầu như chưa được khảo sát. Nghiên cứu này nhằm lấp đầy khoảng trống kiến thức đó bằng cách đánh giá tỷ lệ hiện mắc và các yếu tố liên quan của trầm cảm, lo âu và căng thẳng ở thanh thiếu niên tại quận Kamrup (Metro), Assam.')

add_heading(doc, 'PHƯƠNG PHÁP', 2)
add_page_ref(doc, '734', 'Indian J Community Med', 'Vol. 48(5), 2023')
add_p(doc, 'Thiết kế: Nghiên cứu cắt ngang tại trường học, thực hiện từ tháng 4/2019 đến tháng 6/2020.', bold=True)
add_p(doc, 'Cỡ mẫu và lấy mẫu: Từ tổng số 120 trường trung học công lập tại quận Kamrup (Metro), 10 trường được chọn bằng phương pháp ngẫu nhiên đơn giản. Từ mỗi trường, 36 học sinh được chọn ngẫu nhiên (12 từ mỗi lớp 8, 9, 10), đảm bảo số lượng nam và nữ bằng nhau. Tổng cỡ mẫu: 360 thanh thiếu niên từ 14 đến 17 tuổi.')
add_p(doc, 'Công cụ:', bold=True)
add_p(doc, '• Thang đo Trầm cảm Lo âu Căng thẳng 21 mục (DASS-21 — Depression Anxiety Stress Scale): phiên bản dịch sang tiếng Assam, dịch ngược bởi chuyên gia song ngữ, xác nhận bởi nhóm chuyên gia tâm thần học và y học cộng đồng. DASS-21 là phiên bản rút gọn của thang gốc 42 mục (Lovibond & Lovibond, 1995), sàng lọc đồng thời trầm cảm, lo âu và căng thẳng.')
add_p(doc, '• Thang đo Kuppuswamy sửa đổi (Modified Kuppuswamy Scale): đánh giá tình trạng kinh tế xã hội dựa trên học vấn trụ cột gia đình, nghề nghiệp và thu nhập.')
add_p(doc, '• Lịch phỏng vấn thiết kế sẵn: thu thập dữ liệu nhân khẩu xã hội (tuổi, giới tính, loại gia đình, tình trạng cha mẹ, sử dụng rượu, tiền sử lạm dụng, tiền sử bệnh tâm thần gia đình, chơi game, kết quả học tập).')
add_p(doc, 'Phân tích: SPSS — kiểm định chi-square (chi bình phương), P < 0,05 được coi là có ý nghĩa. Quyền riêng tư và bảo mật được duy trì nghiêm ngặt.')

add_heading(doc, 'KẾT QUẢ', 2)
add_page_ref(doc, '735', 'Indian J Community Med', 'Vol. 48(5), 2023')

add_p(doc, 'Tỷ lệ trầm cảm, lo âu và căng thẳng theo DASS-21 (N = 360):', bold=True)
add_table(doc,
    ['Tình trạng', 'Tỷ lệ', 'n/N'],
    [['Trầm cảm', '22,2%', '80/360'],
     ['Lo âu', '24,4%', '88/360'],
     ['Căng thẳng', '6,9%', '25/360']],
    widths=[5.0, 3.0, 3.0])

doc.add_paragraph()
add_p(doc, 'Phân bố lo âu theo giới tính:', bold=True)
add_table(doc,
    ['Giới tính', 'Có lo âu n (%)', 'Không lo âu n (%)', 'P'],
    [['Nam (n=180)', '54 (30,0%)', '126 (70,0%)', ''],
     ['Nữ (n=180)', '34 (18,9%)', '146 (81,1%)', '0,049']],
    widths=[3.0, 3.5, 3.5, 2.0])

doc.add_paragraph()
add_p(doc, '⚠️ PHÁT HIỆN ĐẶC BIỆT: Nam giới có tỷ lệ lo âu CAO HƠN nữ giới (30,0% so với 18,9%, P = 0,049) — TRÁI NGƯỢC với đại đa số y văn quốc tế, nơi nữ luôn có tỷ lệ cao hơn.', bold=True)

doc.add_paragraph()
add_p(doc, 'Các yếu tố liên quan có ý nghĩa thống kê:', bold=True)
add_table(doc,
    ['Yếu tố', 'P (lo âu)', 'P (trầm cảm)', 'Ghi chú'],
    [['Tình trạng cha mẹ (đơn thân)', '<0,001', '0,001', '52,6–63,2% nếu đơn thân'],
     ['Cha mẹ sử dụng rượu', '<0,001', '<0,001', '38,3–40,4% nếu có rượu'],
     ['Chơi trò chơi điện tử', '0,003', '<0,001', '27,9–30,8% nếu chơi game'],
     ['Kết quả học tập (lưu ban)', '<0,001', '<0,001', '51,7–58,6% nếu lưu ban'],
     ['Tầng lớp KT-XH thấp', '0,028', '0,839', 'Chỉ ý nghĩa với lo âu'],
     ['Tiền sử bị lạm dụng', '0,044', '0,076', 'Chỉ ý nghĩa với lo âu'],
     ['Tiền sử bệnh TT gia đình', '0,043', '0,471', 'Chỉ ý nghĩa với lo âu'],
     ['Giới tính (nam > nữ)', '0,049', '0,076', 'Trái y văn quốc tế']],
    widths=[4.0, 2.0, 2.5, 4.0])

add_page_ref(doc, '736', 'Indian J Community Med', 'Vol. 48(5), 2023')

add_heading(doc, 'THẢO LUẬN', 2)
add_p(doc, 'Tỷ lệ lo âu 24,4% trong nghiên cứu này phù hợp với các nghiên cứu trước tại Ấn Độ. Tuy nhiên, phát hiện đáng chú ý nhất là nam giới có tỷ lệ lo âu cao hơn nữ (30,0% so với 18,9%, P = 0,049). Phát hiện này trái ngược hoàn toàn với y văn quốc tế, nơi đại đa số nghiên cứu cho thấy nữ có tỷ lệ lo âu cao hơn nam (McLean và cộng sự, 2011).')
add_p(doc, 'Kết quả ngược này có thể được giải thích bởi đặc thù văn hóa bộ lạc vùng Đông Bắc Ấn Độ, nơi kỳ vọng về vai trò nam giới có thể khác biệt so với phần còn lại của Ấn Độ. Tuy nhiên, tác giả không đi sâu phân tích cơ chế văn hóa cụ thể.')
add_p(doc, 'Mối liên hệ mạnh giữa hoàn cảnh gia đình bất lợi (cha mẹ đơn thân, sử dụng rượu) và SKTT thanh thiếu niên là phù hợp với y văn. Thanh thiếu niên sống với cha/mẹ đơn thân có tỷ lệ lo âu 63,2% — gấp gần 3 lần so với nhóm sống với cả hai bố mẹ (22,3%).')
add_p(doc, 'Thói quen chơi trò chơi điện tử cũng liên quan có ý nghĩa đến cả lo âu và trầm cảm, tương tự phát hiện của Chen và cộng sự (2023) về rối loạn chơi game trực tuyến (OR = 5,00) tại Trung Quốc.')

add_heading(doc, 'KẾT LUẬN', 2)
add_page_ref(doc, '737', 'Indian J Community Med', 'Vol. 48(5), 2023')
add_p(doc, 'Tỷ lệ đáng kể trầm cảm (22,2%), lo âu (24,4%) và căng thẳng (6,9%) ở thanh thiếu niên trường học vùng Đông Bắc Ấn Độ — khu vực hầu như chưa được nghiên cứu trước đây — cho thấy nhu cầu chú ý đến sức khỏe tâm thần tại đây. Phát hiện nam > nữ về lo âu là trái ngược y văn và cần được nghiên cứu thêm. Các yếu tố gia đình (cha mẹ đơn thân, sử dụng rượu) và lối sống (chơi game) cần được giải quyết trong thiết kế can thiệp.')

# TÀI LIỆU THAM KHẢO
add_heading(doc, 'TÀI LIỆU THAM KHẢO', 2)
add_p(doc, '[Giữ nguyên danh mục tài liệu tham khảo tiếng Anh từ bài gốc — xem bản PDF gốc]', italic=True)

# BẢNG VIẾT TẮT
add_abbreviation_table(doc, [
    ('DASS-21', 'Depression Anxiety Stress Scale 21-item — Thang đo Trầm cảm Lo âu Căng thẳng 21 mục'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('KT-XH', 'Kinh tế-xã hội'),
    ('THCS', 'Trung học cơ sở'),
    ('IJCM', 'Indian Journal of Community Medicine'),
])

# PHẢN BIỆN
add_red_heading(doc, 'PHẢN BIỆN CHI TIẾT')

add_red(doc, 'Điểm mạnh:', bold=True)
add_red(doc, '1. Nghiên cứu tiên phong tại vùng Đông Bắc Ấn Độ — lấp đầy khoảng trống kiến thức quan trọng.')
add_red(doc, '2. Thiết kế lấy mẫu ngẫu nhiên tốt: 10/120 trường, đảm bảo nam nữ bằng nhau.')
add_red(doc, '3. DASS-21 được dịch, dịch ngược và xác nhận bởi chuyên gia — quy trình chuẩn.')

add_red(doc, 'Hạn chế:', bold=True)
add_red(doc, '1. Chỉ sử dụng chi-square (đơn biến) — không thể đánh giá ảnh hưởng độc lập. Chuẩn tối thiểu là hồi quy logistic đa biến (Nakie và cộng sự, 2022).')
add_red(doc, '2. DASS-21 tiếng Assam thiếu báo cáo Cronbach alpha — không rõ độ tin cậy.')
add_red(doc, '3. Thu thập 4/2019–6/2020 trùng giai đoạn đầu COVID-19 nhưng không đề cập tác động (Racine và cộng sự, 2021).')
add_red(doc, '4. Nam > nữ về lo âu (P = 0,049) — giá trị P rất sát ngưỡng 0,05, có thể thay đổi nếu cỡ mẫu lớn hơn.')
add_red(doc, '5. Tác giả không giải thích sâu cơ chế văn hóa bộ lạc Đông Bắc khiến nam lo âu hơn nữ.')

add_red(doc, 'Khoảng trống nghiên cứu (Gap):', bold=True)
add_red(doc, '1. Cần nghiên cứu quy mô lớn hơn tại nhiều quận Đông Bắc Ấn Độ với hồi quy đa biến.')
add_red(doc, '2. Khảo sát sâu yếu tố văn hóa bộ lạc đặc thù giải thích tại sao nam > nữ.')
add_red(doc, '3. So sánh thành thị-nông thôn trong cùng khu vực.')
add_red(doc, '4. Xác thực DASS-21 tiếng Assam với Cronbach alpha đầy đủ.')
add_red(doc, '5. Nghiên cứu can thiệp nhắm vào giảm chơi game và hỗ trợ gia đình đơn thân.')

fname = '11_Saikia_2023_IJCM.docx'
doc.save(fname)
sys.stderr.write(f'{fname} saved OK\n')
