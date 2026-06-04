# -*- coding: utf-8 -*-
"""
Tạo bản dịch DOCX cho V-NAMHS 2022 Report (51 trang)
Sử dụng template Jenkins et al.
"""
import sys, os, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Add parent dir for template import
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import (
    create_doc, add_link_and_qr, add_heading, add_page_ref,
    add_p, add_red, add_red_heading, add_table, add_info_table,
    add_image, add_abbreviation_table
)

JOURNAL = 'V-NAMHS Report'
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '06_VNAMHS_2022.docx')

doc = create_doc()

# ============================================================
# LINK + QR
# ============================================================
add_link_and_qr(doc, 'https://qcmhr.org/docman/reports/15-vnamhs-report-eng-15-feb-2023/file', 'QR_VNAMHS2022.png')

# ============================================================
# TITLE
# ============================================================
add_heading(doc, 'BÁO CÁO: Khảo sát Sức khỏe Tâm thần Vị thành niên Việt Nam (V-NAMHS)', 1)
add_p(doc, 'Báo cáo về Các Kết quả Chính -- Tháng 11/2022', size=12, bold=True)
add_p(doc, 'Bản dịch tiếng Việt đầy đủ (51 trang)', size=11, italic=True)

# ============================================================
# INFO TABLE
# ============================================================
add_heading(doc, 'THÔNG TIN BÀI BÁO', 2)
add_info_table(doc, [
    ('Tiêu đề gốc', 'Viet Nam Adolescent Mental Health Survey (V-NAMHS): Report on Main Findings'),
    ('Tác giả', 'Vũ Mạnh Lợi, Nguyễn Đức Vinh (Viện Xã hội học); Đào Thị Khánh Hòa (Hội Xã hội học VN); '
     'Holly E. Erskine, Cartiah McGrath, Krystina Wallis, Sarah J. Blondell, Harvey A. Whiteford, James G. Scott (ĐH Queensland); '
     'Robert Blum, Shoshanna Fine, Mengmeng Li, Astha Ramaiya (ĐH Johns Hopkins)'),
    ('Tổ chức', 'Viện Xã hội học (IOS/VASS), Đại học Queensland (UQ), Trường Y tế Công cộng Bloomberg - ĐH Johns Hopkins (JHSPH)'),
    ('Loại', 'Báo cáo khảo sát quốc gia (National survey report)'),
    ('Năm', '2022'),
    ('Mẫu', 'N = 5,996 cặp phụ huynh-vị thành niên (10-17 tuổi)'),
    ('Phạm vi', '38 tỉnh/thành phố trên toàn Việt Nam, đại diện quốc gia'),
    ('Công cụ chẩn đoán', 'DISC-5 (theo DSM-5) cho 6 rối loạn: ám sợ xã hội, lo âu lan tỏa, MDD, PTSD, rối loạn hành vi, ADHD'),
    ('Thu thập dữ liệu', '21/9/2021 -- 16/12/2021, 127 phỏng vấn viên'),
    ('Tỷ lệ phản hồi', '81.1%'),
])

# ============================================================
# ABBREVIATION TABLE
# ============================================================
add_abbreviation_table(doc, [
    ('V-NAMHS', 'Khảo sát Sức khỏe Tâm thần Vị thành niên Việt Nam (Viet Nam Adolescent Mental Health Survey)'),
    ('DISC-5', 'Bảng Phỏng vấn Chẩn đoán cho Trẻ em, Phiên bản 5 (Diagnostic Interview Schedule for Children, Version 5)'),
    ('DSM-5', 'Cẩm nang Chẩn đoán và Thống kê các Rối loạn Tâm thần, Ấn bản thứ 5'),
    ('SDQ', 'Bảng hỏi Điểm mạnh và Khó khăn (Strengths and Difficulties Questionnaire)'),
    ('CBCL', 'Bảng kiểm Hành vi Trẻ em (Child Behaviour Checklist)'),
    ('YSR', 'Bản Tự báo cáo của Thanh thiếu niên (Youth Self-Report)'),
    ('ADHD', 'Rối loạn tăng động giảm chú ý (Attention-deficit/hyperactivity disorder)'),
    ('PTSD', 'Rối loạn căng thẳng sau sang chấn (Posttraumatic stress disorder)'),
    ('MDD', 'Rối loạn trầm cảm chủ yếu (Major depressive disorder)'),
    ('ACEs', 'Trải nghiệm bất lợi thời thơ ấu (Adverse Childhood Experiences)'),
    ('GBD', 'Nghiên cứu Gánh nặng Bệnh tật Toàn cầu (Global Burden of Disease Study)'),
    ('GSHS', 'Khảo sát Sức khỏe Học sinh Toàn cầu dựa trên Trường học'),
    ('GEAS', 'Nghiên cứu Vị thành niên Sớm Toàn cầu (Global Early Adolescent Study)'),
    ('IOS', 'Viện Xã hội học (Institute of Sociology)'),
    ('VASS', 'Viện Hàn lâm Khoa học Xã hội Việt Nam'),
    ('UQ', 'Đại học Queensland'),
    ('JHSPH', 'Trường Y tế Công cộng Bloomberg, ĐH Johns Hopkins'),
    ('GOPFP', 'Tổng cục Dân số và Kế hoạch hóa Gia đình'),
    ('GSO', 'Tổng cục Thống kê Việt Nam'),
    ('HSPI', 'Viện Chiến lược và Chính sách Y tế'),
    ('LMICs', 'Các quốc gia thu nhập thấp và trung bình'),
    ('HICs', 'Các quốc gia thu nhập cao'),
    ('WHO', 'Tổ chức Y tế Thế giới'),
    ('MOH', 'Bộ Y tế'),
    ('mhGAP', 'Chương trình Hành động Khoảng trống Sức khỏe Tâm thần (Mental Health Gap Action Programme)'),
    ('EA', 'Khu vực điều tra (Enumeration Area)'),
    ('POPFP', 'Chi cục Dân số và Kế hoạch hóa Gia đình cấp tỉnh'),
    ('PSC-17', 'Bảng kiểm triệu chứng Nhi khoa-17 (Pediatric Symptom Checklist)'),
    ('PHQ-9', 'Bảng hỏi Sức khỏe Bệnh nhân-9 (Patient Health Questionnaire)'),
    ('GAD-7', 'Rối loạn Lo âu Tổng quát-7 (Generalised Anxiety Disorder-7)'),
])

# ============================================================
# TÓM TẮT TỔNG QUAN (Executive Summary)
# ============================================================
add_heading(doc, 'TÓM TẮT TỔNG QUAN', 1)
add_page_ref(doc, 'i-ii', JOURNAL)

add_p(doc, 'Tỷ lệ hiện mắc (prevalence) của các rối loạn tâm thần (mental disorders) ở vị thành niên tại Việt Nam phần lớn chưa được biết đến. Dữ liệu tỷ lệ hiện mắc chính xác rất quan trọng cho công tác phòng ngừa hiệu quả, lập kế hoạch dịch vụ và ưu tiên chính sách sức khỏe tâm thần. Báo cáo này trình bày các kết quả từ V-NAMHS, khảo sát đầu tiên đo lường tỷ lệ hiện mắc các rối loạn tâm thần trên một mẫu đại diện quốc gia (nationally representative sample) của các hộ gia đình trên toàn Việt Nam.')

add_p(doc, 'Các kết quả chính:', size=12, bold=True)
add_p(doc, '- Trong 12 tháng qua, cứ 5 vị thành niên thì có 1 người (21,7%) có vấn đề sức khỏe tâm thần; cứ 30 người thì có 1 người đáp ứng tiêu chí cho rối loạn tâm thần (3,3%).')
add_p(doc, '- Lo âu (anxiety) là vấn đề sức khỏe tâm thần phổ biến nhất (18,6%), tiếp theo là trầm cảm (4,3%).')
add_p(doc, '- Rối loạn lo âu theo DSM-5 = 2,3% (KHÔNG phải 3,3%); rối loạn trầm cảm chủ yếu (MDD) = 0,9%; bất kỳ rối loạn nào = 3,3%.')
add_p(doc, '- Chỉ 8,4% vị thành niên có vấn đề SKTT đã tiếp cận dịch vụ hỗ trợ trong 12 tháng qua.')
add_p(doc, '- Chỉ 5,1% phụ huynh nhận ra con cần giúp đỡ, dù 21,7% vị thành niên có vấn đề SKTT.')
add_p(doc, '- 7,7% vị thành niên thường xuyên trải qua vấn đề cảm xúc/hành vi nhiều hơn bình thường trong đại dịch COVID-19.')

add_p(doc, 'Hàm ý:', size=12, bold=True)
add_p(doc, '- SKTT là vấn đề y tế công cộng cần sự quan tâm của nhà hoạch định chính sách. Cần các Kế hoạch SKTT Quốc gia xem xét nhu cầu cụ thể của vị thành niên.')
add_p(doc, '- Phụ huynh và gia đình là nguồn hỗ trợ chính. Cần cải thiện hiểu biết về SKTT, giảm kỳ thị, nâng cao nhận thức về dịch vụ.')
add_p(doc, '- Cần tích hợp sàng lọc SKTT vào dịch vụ y tế tổng quát, đào tạo nhân viên y tế về lộ trình chuyển tuyến.')
add_p(doc, '- Cần đưa SKTT vào kế hoạch ứng phó các sự kiện quy mô dân số trong tương lai (đại dịch, thiên tai, xung đột).')

# ============================================================
# GIỚI THIỆU
# ============================================================
add_heading(doc, 'GIỚI THIỆU', 1)
add_page_ref(doc, '1-4', JOURNAL)
add_image(doc, 'VNAMHS2022_p1_img1.jpeg', 14, 'Trang bìa báo cáo V-NAMHS')

add_heading(doc, 'Bối cảnh', 2)
add_p(doc, 'Vị thành niên chiếm khoảng 14,5% dân số Việt Nam (~14 triệu người 10-19 tuổi). Rối loạn tâm thần là nguyên nhân hàng đầu gây khuyết tật ở vị thành niên toàn cầu (Erskine et al. 2015). Cho đến nay, Việt Nam chưa có chính sách sức khỏe tâm thần theo định nghĩa WHO. Quản trị SKTT bị phân mảnh, chỉ được các luật và quyết định hiện hành bao phủ một phần.')

add_p(doc, 'Các quyết định liên quan:')
add_p(doc, '- Luật Người khuyết tật (51/2010/QH12): xác định khuyết tật tâm thần là một nhóm phụ.')
add_p(doc, '- QĐ 1215/QĐ-TTg (2011): hỗ trợ xã hội dựa vào cộng đồng cho người rối loạn tâm thần.')
add_p(doc, '- QĐ 1364/QĐ-LĐTBXH (2012): đề án mạng lưới cơ sở bảo trợ xã hội.')
add_p(doc, '- QĐ 155/QĐ-TTg (29/1/2022): Kế hoạch Quốc gia Phòng chống Bệnh Không lây nhiễm và Rối loạn SKTT 2022-2025.')
add_p(doc, 'Tất cả tập trung vào người lớn có bệnh tâm thần nặng, không đề cập cụ thể đến vị thành niên.')

add_p(doc, 'V-NAMHS được phát triển với mục tiêu cụ thể: đưa ra ước tính đại diện quốc gia về tỷ lệ hiện mắc các rối loạn tâm thần ở vị thành niên Việt Nam 10-17 tuổi.')

add_heading(doc, 'V-NAMHS là gì?', 2)
add_p(doc, 'V-NAMHS là khảo sát hộ gia đình đại diện quốc gia trên vị thành niên và người chăm sóc chính. Ba mục tiêu cốt lõi:')
add_p(doc, '1. Xác định tỷ lệ hiện mắc các rối loạn tâm thần ở vị thành niên 10-17 tuổi.')
add_p(doc, '2. Đo lường các yếu tố nguy cơ và yếu tố bảo vệ liên quan.')
add_p(doc, '3. Thiết lập mô hình sử dụng dịch vụ, rào cản tiếp cận, nhu cầu nhận thức được.')

add_heading(doc, 'Ai thực hiện V-NAMHS?', 2)
add_p(doc, 'Viện Xã hội học (IOS/VASS) chịu trách nhiệm thực hiện. V-NAMHS là một phần của NAMHS (Kenya, Indonesia, Việt Nam) do ĐH Queensland dẫn đầu với sự hỗ trợ từ JHSPH. Các cộng tác viên: HSPI/MOH, SESD/GSO, GOPFP/MOH.')
add_p(doc, 'Phê duyệt đạo đức: ĐH Y tế Công cộng Hà Nội (499/2019/YTCC-HD3) và UQ HREC (2019001268).')

add_heading(doc, 'Ai tham gia V-NAMHS?', 2)
add_page_ref(doc, '3-4', JOURNAL)
add_p(doc, 'Khảo sát hộ gia đình tại 38 tỉnh. Vị thành niên đủ điều kiện: 10-17 tuổi, sống với người chăm sóc chính >50% thời gian. Vị thành niên 18-19 tuổi bị loại trừ (tỷ lệ cao sống xa gia đình, công cụ chẩn đoán không phù hợp).')

add_heading(doc, 'Bảng 1. Tuyển chọn mẫu cho V-NAMHS', 3)
add_table(doc,
    ['Hạng mục', 'Số lượng'],
    [
        ['Tổng số hộ gia đình trong phạm vi tiếp cận', '7.599'],
        ['Hộ gia đình không được tiếp cận/không có mặt', '502'],
        ['Hộ gia đình từ chối tham gia', '911'],
        ['Hộ gia đình không đủ điều kiện', '138'],
        ['Hộ gia đình tham gia, đủ điều kiện', '6.048'],
        ['Hộ gia đình có dữ liệu không đầy đủ', '52'],
        ['Hộ gia đình có dữ liệu đầy đủ (mẫu cuối cùng)', '5.996'],
    ],
    widths=[10, 4]
)

add_heading(doc, 'Công cụ đo lường', 2)
add_p(doc, 'Rối loạn tâm thần được đánh giá bằng DISC-5 (Bitsko et al. 2019, Shaffer et al. 2000), công cụ chẩn đoán chuẩn hóa cho trẻ em/vị thành niên. Sáu rối loạn được đo: ám sợ xã hội, rối loạn lo âu lan tỏa, MDD, PTSD, rối loạn hành vi, ADHD. Ám sợ xã hội + lo âu lan tỏa = "rối loạn lo âu".')
add_p(doc, 'Thang đo bổ sung: hành vi tự tử, tự gây thương tích, yếu tố nguy cơ/bảo vệ (bắt nạt, trường học, gia đình, hành vi tình dục, sử dụng chất, ACEs, bệnh tâm thần cha mẹ, lòng tự trọng), sử dụng dịch vụ, COVID-19.')

add_heading(doc, 'COVID-19 và V-NAMHS', 2)
add_p(doc, 'COVID-19 đã ảnh hưởng đến tiến trình V-NAMHS. Thu thập dữ liệu ban đầu dự kiến 6/2020, bị hoãn đến 6/2021 rồi 9/2021. Hạn chế di chuyển -> phỏng vấn viên phải là người địa phương (nhân viên POPFP), đào tạo trực tuyến, quy trình giám sát chất lượng toàn diện.')

add_heading(doc, 'Phạm vi báo cáo', 2)
add_p(doc, 'Ba chương chính: Sức khỏe tâm thần (bao gồm hành vi tự tử/tự gây thương tích), Sử dụng dịch vụ, và COVID-19. Tất cả kết quả đã được trọng số hóa (weighted) để đại diện cho phân bố tuổi-giới tính và thành thị-nông thôn.')

# ============================================================
# ĐẶC ĐIỂM MẪU KHẢO SÁT
# ============================================================
add_heading(doc, 'ĐẶC ĐIỂM MẪU KHẢO SÁT', 1)
add_page_ref(doc, '6-8', JOURNAL)

add_heading(doc, 'Bảng 2. Mẫu vị thành niên theo giới tính và nhóm tuổi', 3)
add_table(doc,
    ['Tuổi', 'Nam %', 'Nam n', 'Nữ %', 'Nữ n', 'Tổng %', 'Tổng n'],
    [
        ['10-13', '28,5', '1.709', '25,7', '1.539', '54,2', '3.248'],
        ['14-17', '24,1', '1.442', '21,8', '1.306', '45,8', '2.748'],
        ['10-17', '52,6', '3.151', '47,5', '2.845', '100,0', '5.996'],
    ],
    widths=[2.5, 2, 2, 2, 2, 2, 2]
)

add_heading(doc, 'Bảng 3. Mẫu vị thành niên theo tình trạng học vấn và việc làm', 3)
add_table(doc,
    ['Chỉ số', '%', 'n'],
    [
        ['Đang đi học', '94,5', '5.666'],
        ['Hiện không đi học nhưng đã đi học trong 12 tháng qua', '1,0', '62'],
        ['Đã đi học nhưng không trong 12 tháng qua', '4,4', '263'],
        ['Chưa bao giờ đi học', '0,1', '5'],
        ['Đang có việc làm', '3,0', '179'],
        ['Hiện không đi học VÀ không có việc làm', '3,2', '194'],
    ],
    widths=[10, 2, 2]
)

add_heading(doc, 'Bảng 4. Mẫu phụ huynh theo nhân khẩu học', 3)
add_table(doc,
    ['Chỉ số', '%', 'n'],
    [
        ['Tuổi trung bình (năm)', '44,2', '--'],
        ['Nữ', '71,7', '4.301'],
        ['Mẹ/mẹ kế', '63,6', '3.814'],
        ['Cha/cha kế', '25,0', '1.497'],
        ['Ông bà', '10,3', '616'],
        ['Đã kết hôn', '89,6', '5.374'],
        ['Trung học cơ sở', '32,4', '1.945'],
        ['Giáo dục đại học', '19,2', '1.152'],
        ['Toàn thời gian', '54,1', '3.246'],
        ['Tỷ lệ đồng thời là chủ hộ', '53,7', '3.220'],
    ],
    widths=[10, 2, 2]
)

# ============================================================
# SỨC KHỎE TÂM THẦN
# ============================================================
add_heading(doc, 'SỨC KHỎE TÂM THẦN', 1)
add_page_ref(doc, '9-16', JOURNAL)

add_heading(doc, 'Tổng quan', 2)
add_p(doc, 'Vị thành niên 10-19 tuổi chiếm 14,3% dân số Việt Nam. Còn tương đối ít thông tin về tỷ lệ hiện mắc các rối loạn tâm thần. Tổng quan tài liệu trước đây tìm thấy ước tính rất khác nhau: 8% đến 29% (Samuels et al. 2018). Các nghiên cứu hiện có bị hạn chế bởi: cỡ mẫu nhỏ, phạm vi địa lý hẹp, sử dụng thang triệu chứng thay vì công cụ chẩn đoán.')

add_p(doc, 'Ví dụ: Weiss et al. (2014) sử dụng SDQ trên 1.314 phụ huynh/591 vị thành niên: 10,7-11,9% đạt ngưỡng "ranh giới" cho vấn đề SKTT. Samuels et al. (2018) sử dụng SDQ trên 402 vị thành niên: 13,4% nằm trong vùng "bất thường". Tuy nhiên đây là thang triệu chứng, không phải công cụ chẩn đoán.')

add_heading(doc, 'Đo lường', 2)
add_p(doc, 'DISC-5 dùng để đánh giá 6 rối loạn tâm thần: ám sợ xã hội, lo âu lan tỏa (gộp = "rối loạn lo âu"), MDD, PTSD, rối loạn hành vi, ADHD. Ngoại trừ ADHD (hỏi phụ huynh), tất cả mô-đun được thực hiện với vị thành niên.')

add_p(doc, 'Phân biệt giữa "vấn đề sức khỏe tâm thần" (mental health problems) và "rối loạn tâm thần" (mental disorders): vấn đề SKTT = đáp ứng >= 1/2 triệu chứng; rối loạn = đáp ứng đầy đủ tiêu chí DSM-5 + suy giảm chức năng.')

# Table 5 - Key definitions
add_heading(doc, 'Bảng 5. Định nghĩa về vấn đề SKTT và rối loạn tâm thần', 3)
add_page_ref(doc, '11', JOURNAL)
add_table(doc,
    ['', 'Vấn đề SKTT (Mental health problem)', 'Rối loạn tâm thần (Mental disorder)'],
    [
        ['Định nghĩa chung',
         'Ảnh hưởng cách suy nghĩ, cảm nhận, hành xử ở mức độ thấp hơn rối loạn. Có thể tạm thời hoặc phản ứng cấp tính.',
         'Hội chứng có ý nghĩa lâm sàng, liên quan đến đau khổ, suy giảm chức năng, nguy cơ tử vong/khuyết tật tăng.'],
        ['Định nghĩa V-NAMHS',
         '>= 1/2 triệu chứng DSM-5 được xác nhận (triệu chứng dưới ngưỡng - subthreshold). Bao gồm cả những người đáp ứng tiêu chí rối loạn.',
         'Đáp ứng tất cả triệu chứng DSM-5 (đủ ngưỡng) + xác nhận suy giảm chức năng. Theo thuật toán chấm điểm DISC-5.'],
        ['Thuật ngữ VD',
         'Trầm cảm, Lo âu, Căng thẳng sau sang chấn, Vấn đề hành vi, Vấn đề thiếu chú ý/tăng động',
         'MDD, Rối loạn lo âu, PTSD, Rối loạn hành vi, ADHD'],
    ],
    widths=[3, 5.5, 5.5]
)

add_image(doc, 'VNAMHS2022_p19_img1.jpeg', 14, 'Hình 2. Tỷ lệ hiện mắc 12 tháng theo ngưỡng triệu chứng và suy giảm chức năng (Trang 13 báo cáo gốc)')

# Table 6 - KEY: 21.7% mental health problems
add_heading(doc, 'Kết quả: Vấn đề sức khỏe tâm thần', 2)
add_heading(doc, 'Bảng 6. Tỷ lệ hiện mắc 12 tháng của vấn đề SKTT ở nhóm 10-17 tuổi', 3)
add_page_ref(doc, '12', JOURNAL)
add_table(doc,
    ['Vấn đề SKTT', '10-13 %', '10-13 n', '14-17 %', '14-17 n', '10-17 %', '10-17 n'],
    [
        ['Nam', '21,4', '357', '20,2', '292', '20,8', '649'],
        ['Nữ', '20,2', '302', '25,3', '350', '22,6', '651'],
        ['Tổng', '20,8', '659', '22,7', '462', '21,7', '1.301'],
    ],
    widths=[3, 1.5, 1.5, 1.5, 1.5, 1.5, 2]
)
add_p(doc, 'Một phần năm vị thành niên (21,7%) có vấn đề SKTT trong 12 tháng qua. Không có sự khác biệt có ý nghĩa thống kê giữa nam/nữ hoặc giữa nhóm tuổi.', size=11)

add_p(doc, 'Phân bố theo loại (Hình 1):', size=12, bold=True)
add_p(doc, '- Lo âu (Anxiety): 18,6% -- phổ biến nhất')
add_p(doc, '- Trầm cảm (Depression): 4,3%')
add_p(doc, '- Vấn đề thiếu chú ý/tăng động: 2,8%')
add_p(doc, '- Căng thẳng sau sang chấn: 1,0%')
add_p(doc, '- Vấn đề hành vi: 0,7%')

add_image(doc, 'VNAMHS2022_p16_img1.jpeg', 12, 'Hình 1. Tỷ lệ hiện mắc 12 tháng của vấn đề SKTT theo loại (Trang 12)')

add_p(doc, 'Phân nhóm 21,7% theo ngưỡng triệu chứng (Hình 2):', size=12, bold=True)
add_p(doc, '- Triệu chứng đủ ngưỡng + suy giảm chức năng (= rối loạn tâm thần): 3,3%')
add_p(doc, '- Triệu chứng dưới ngưỡng + suy giảm chức năng: 10,3%')
add_p(doc, '- Triệu chứng đủ ngưỡng, không suy giảm: 0,6%')
add_p(doc, '- Triệu chứng dưới ngưỡng, không suy giảm: 7,5%')

add_heading(doc, 'Bảng 7. Suy giảm chức năng theo lĩnh vực', 3)
add_table(doc,
    ['Lĩnh vực suy giảm', '%', 'n'],
    [
        ['Gia đình', '67,0', '549'],
        ['Bạn bè', '47,0', '384'],
        ['Trường học/công việc', '45,4', '371'],
        ['Đau khổ cá nhân', '34,6', '284'],
    ],
    widths=[8, 3, 3]
)
add_p(doc, 'N có trọng số = 819', size=9, italic=True)

# Table 8 - KEY: disorders by type
add_heading(doc, 'Kết quả: Rối loạn tâm thần (Mental disorders)', 2)
add_heading(doc, 'Bảng 8. Tỷ lệ hiện mắc 12 tháng của rối loạn tâm thần ở nhóm 10-17 tuổi theo loại', 3)
add_page_ref(doc, '14', JOURNAL)
add_table(doc,
    ['Rối loạn tâm thần', '%', 'n'],
    [
        ['Rối loạn lo âu (Anxiety disorders)*', '2,3', '135'],
        ['Rối loạn trầm cảm chủ yếu (MDD)', '0,9', '51'],
        ['ADHD', '0,5', '29'],
        ['PTSD', '0,3', '19'],
        ['Rối loạn hành vi (Conduct disorder)', '0,2', '12'],
        ['BẤT KỲ RỐI LOẠN NÀO', '3,3', '200'],
    ],
    widths=[8, 3, 3]
)
add_p(doc, '* Rối loạn lo âu = ám sợ xã hội + lo âu lan tỏa. Tỷ lệ cao hơn đáng kể so với tất cả các rối loạn khác.', size=10, italic=True)
add_p(doc, 'QUAN TRỌNG: Rối loạn lo âu DSM-5 = 2,3% (KHÔNG phải 3,3%). Con số 3,3% là tỷ lệ BẤT KỲ rối loạn nào. MDD = 0,9%.', size=11, bold=True)
add_p(doc, 'Không có sự khác biệt giữa nam (3,3%) và nữ (3,4%), hoặc nhóm nhỏ tuổi (2,9%) và lớn tuổi (3,9%). Dưới 1% (0,6%) có 2+ rối loạn.')

add_heading(doc, 'Bảng 9. Rối loạn tâm thần theo lĩnh vực suy giảm', 3)
add_table(doc,
    ['Lĩnh vực suy giảm', '%', 'n'],
    [
        ['Gia đình', '74,2', '149'],
        ['Trường học/công việc', '64,1', '128'],
        ['Bạn bè', '63,3', '127'],
        ['Đau khổ cá nhân', '51,8', '104'],
    ],
    widths=[8, 3, 3]
)
add_p(doc, 'N có trọng số = 200', size=9, italic=True)

# Suicidal behaviours
add_heading(doc, 'Hành vi tự tử và tự gây thương tích', 2)
add_page_ref(doc, '14-15', JOURNAL)

add_heading(doc, 'Bảng 10. Hành vi tự tử ở nhóm 10-17 tuổi', 3)
add_table(doc,
    ['', 'Ý tưởng tự tử 12th (%)', 'Kế hoạch 12th (%)', 'Mưu toan 12th (%)', 'Từng mưu toan (%)'],
    [
        ['Tỷ lệ chung', '1,4', '0,4', '0,2', '1,6'],
        ['Có vấn đề SKTT', '73,5', '83,7', '77,4', '73,5'],
        ['Có rối loạn TT', '30,8', '46,6', '61,3', '28,1'],
    ],
    widths=[4, 3, 3, 3, 3]
)

add_heading(doc, 'Bảng 11. Tự gây thương tích ở nhóm 10-17 tuổi', 3)
add_table(doc,
    ['', 'Trong 12 tháng qua (%)', 'Từng (%)'],
    [
        ['Tỷ lệ chung', '~1,1', '4,7'],
        ['Có vấn đề SKTT', '76,3', '64,5'],
        ['Có rối loạn TT', '29,0', '20,9'],
    ],
    widths=[6, 4, 4]
)
add_p(doc, 'Hơn 70% những người báo cáo hành vi tự tử và 76% tự gây thương tích trong 12 tháng qua có vấn đề SKTT.')

# Considerations - Mental Health
add_heading(doc, 'Diễn giải', 2)
add_page_ref(doc, '15-16', JOURNAL)
add_p(doc, 'V-NAMHS phát hiện SKTT kém là vấn đề phổ biến: 1/5 vị thành niên (21,7%) có vấn đề SKTT, 1/30 (3,3%) có rối loạn tâm thần. Rất ít nghiên cứu dùng công cụ chẩn đoán để đo lường ở Việt Nam. So sánh trực tiếp khó khăn do khác biệt phương pháp.')

add_p(doc, 'Tỷ lệ hành vi tự tử trong V-NAMHS thấp hơn GSHS 2019 (ý tưởng tự tử 1,4% vs 15,6%). Có thể do: (1) kỳ thị khi được hỏi trực tiếp bởi phỏng vấn viên vs tự điền; (2) khác biệt phương pháp và định nghĩa câu hỏi.')

add_heading(doc, 'Hạn chế', 2)
add_p(doc, '- Phỏng vấn trực tiếp: kỳ thị và thiếu nhận thức SKTT có thể ảnh hưởng sự sẵn lòng tiết lộ.')
add_p(doc, '- DSM-5 dựa trên tiêu chí phương Tây, có thể không tính đến khác biệt văn hóa VN (VD: triệu chứng cơ thể dự đoán lo âu/trầm cảm ở VTN Việt Nam, Kim et al. 2019).')
add_p(doc, '- Nỗ lực điều chỉnh: dịch/dịch ngược DISC-5, đánh giá bởi bác sĩ lâm sàng VN, nghiên cứu thí điểm, sửa đổi liên tục 2019-2020.')

add_heading(doc, 'Hàm ý', 2)
add_p(doc, '- SKTT là vấn đề y tế công cộng cần chính sách. Lo âu phổ biến nhất -> chiến dịch nâng cao nhận thức về lo âu.')
add_p(doc, '- 94,5% VTN đang đi học -> triển khai sàng lọc và quản lý SKTT trong trường học.')
add_p(doc, '- Cần chính sách SKTT quốc gia xem xét VTN: QĐ 1442 và 2138 của Bộ GD-ĐT (2022) là bước đầu nhưng cần khung chính sách toàn diện hơn bao gồm cả y tế.')
add_p(doc, '- Hành động cho SKTT VTN mang lại lợi ích hiện tại, dài hạn, và cho thế hệ tiếp theo (Patton et al. 2016).')

# ============================================================
# SỬ DỤNG DỊCH VỤ
# ============================================================
add_heading(doc, 'SỬ DỤNG DỊCH VỤ', 1)
add_page_ref(doc, '17-23', JOURNAL)

add_heading(doc, 'Tổng quan', 2)
add_p(doc, 'Khởi phát triệu chứng rối loạn tâm thần có xu hướng xảy ra trong giai đoạn VTN, mang lại cơ hội can thiệp sớm. Tại Việt Nam, hầu hết dịch vụ SKTT phục vụ người lớn có bệnh tâm thần nặng (tâm thần phân liệt). Thiếu hụt đáng kể dịch vụ cho VTN, đặc biệt khu vực nông thôn.')

add_heading(doc, 'Đo lường', 2)
add_p(doc, 'Thuật ngữ "các vấn đề về cảm xúc và hành vi" được sử dụng thay vì "sức khỏe tâm thần" để giảm kỳ thị. Phạm vi nhà cung cấp rộng: y tế, giáo dục, tôn giáo/truyền thống. Câu hỏi hỏi phụ huynh (trừ hỗ trợ không chính thức và tự giúp -> hỏi VTN).')

add_heading(doc, 'Kết quả', 2)

add_heading(doc, 'Bảng 12. Tần suất tiếp cận dịch vụ theo giới tính', 3)
add_page_ref(doc, '18', JOURNAL)
add_p(doc, 'Trong số VTN có vấn đề SKTT (n=1.301), CHỈ 8,4% đã sử dụng dịch vụ. Tổng thể: 6,5% (n=389) VTN sử dụng dịch vụ. Hầu hết phụ huynh (80%) cho biết dịch vụ hữu ích.', size=11, bold=True)
add_table(doc,
    ['Giới tính', '1 lần %', '1 lần n', '2-4 lần %', '2-4 lần n', '>=5 lần %', '>=5 lần n'],
    [
        ['Nam', '47,2', '109', '26,3', '61', '4,5', '10'],
        ['Nữ', '56,0', '89', '26,2', '41', '1,7', '3'],
        ['Tổng', '50,8', '198', '26,2', '102', '3,4', '13'],
    ],
    widths=[2.5, 2, 2, 2, 2, 2, 2]
)
add_p(doc, '50,8% chỉ sử dụng dịch vụ MỘT LẦN trong 12 tháng qua.')

add_heading(doc, 'Bảng 13. Nhà cung cấp dịch vụ được sử dụng nhiều nhất', 3)
add_table(doc,
    ['Loại nhà cung cấp', '%', 'n'],
    [
        ['Bác sĩ hoặc y tá', '56,2', '219'],
        ['Nhân viên y tế cộng đồng', '10,7', '42'],
        ['Nhân viên trường học', '5,5', '22'],
        ['Lãnh đạo tôn giáo/tín ngưỡng', '4,5', '17'],
        ['Chuyên gia (bác sĩ tâm thần)', '1,4', '5'],
        ['Thầy thuốc truyền thống', '0,1', '0'],
    ],
    widths=[8, 3, 3]
)

add_image(doc, 'VNAMHS2022_p27_img1.jpeg', 12, 'Hình 3. Ngưỡng triệu chứng ở VTN có vấn đề SKTT sử dụng dịch vụ (Trang 19)')

add_heading(doc, 'Nhu cầu nhận thức được và rào cản', 2)
add_p(doc, 'CHỈ 5,1% (n=307) phụ huynh xác định con cần giúp đỡ, mặc dù 21,7% VTN có vấn đề SKTT.', size=11, bold=True)
add_p(doc, 'Trong số phụ huynh nhận ra nhu cầu: 37,7% cho biết con đã nhận đủ giúp đỡ.')

add_heading(doc, 'Bảng 14. Rào cản tiếp cận dịch vụ', 3)
add_page_ref(doc, '20', JOURNAL)
add_table(doc,
    ['Rào cản (phụ huynh báo cáo)', '%', 'n'],
    [
        ['Thích tự xử lý hoặc với gia đình', '20,4', '63'],
        ['Không chắc con có cần giúp đỡ không', '10,7', '33'],
        ['Không chắc nên tìm ở đâu', '10,0', '30'],
        ['Nghĩ vấn đề sẽ tự khỏi', '5,8', '18'],
        ['Không có nơi nhận giúp đỡ', '3,8', '12'],
        ['Chi phí quá cao', '2,7', '8'],
        ['Khó đến nơi cung cấp dịch vụ', '2,0', '6'],
        ['Lo suy nghĩ của người khác', '1,3', '4'],
        ['Không muốn nói với người lạ', '1,3', '4'],
        ['Con từ chối giúp đỡ', '0,7', '2'],
    ],
    widths=[8, 3, 3]
)

add_heading(doc, 'Bảng 15. Hỗ trợ không chính thức -- Người VTN nói chuyện khi lo lắng', 3)
add_table(doc,
    ['Người được nói chuyện', '%', 'n'],
    [
        ['Thành viên gia đình', '73,9', '4.429'],
        ['Bạn bè', '38,2', '2.288'],
        ['Giáo viên', '2,9', '176'],
        ['Bạn trai/bạn gái', '2,3', '137'],
        ['Thành viên cộng đồng', '1,6', '94'],
        ['Bác sĩ', '0,5', '31'],
        ['Lãnh đạo tôn giáo', '0,4', '22'],
        ['Không nói với ai', '9,5', '--'],
    ],
    widths=[8, 3, 3]
)

add_heading(doc, 'Bảng 16. Chiến lược tự giúp', 3)
add_table(doc,
    ['Chiến lược', '%', 'n'],
    [
        ['Tập thể dục nhiều hơn', '30,7', '1.839'],
        ['Làm nhiều hơn những điều bản thân thích', '24,8', '1.490'],
        ['Tìm hỗ trợ từ gia đình', '19,2', '1.149'],
        ['Tìm hỗ trợ từ bạn bè', '17,7', '1.060'],
        ['Cải thiện chế độ ăn', '14,7', '884'],
        ['Tìm thông tin sách/TV', '9,0', '538'],
        ['Mạng xã hội', '6,9', '416'],
        ['Thiền/thư giãn', '1,6', '97'],
        ['Không sử dụng chiến lược nào', '11,7', '--'],
    ],
    widths=[8, 3, 3]
)

add_heading(doc, 'Diễn giải -- Sử dụng dịch vụ', 2)
add_p(doc, 'Nhu cầu chưa được đáp ứng lớn: trong số 21,7% VTN có vấn đề SKTT, chỉ 8,4% sử dụng dịch vụ. Phù hợp với các khảo sát quốc tế cho thấy hầu hết người dân ở LMICs có rối loạn TT không được điều trị (Demyttenaere et al. 2004).')
add_p(doc, 'Vấn đề hiểu biết SKTT: chỉ 5,1% phụ huynh nhận ra nhu cầu. 20,4% thích tự xử lý. Kỳ thị ảnh hưởng hành vi tìm kiếm giúp đỡ (Kamimura et al. 2018, Mai Do et al. 2014).')
add_p(doc, 'Phát hiện đáng chú ý: 71,9% những người tiếp cận dịch vụ KHÔNG có vấn đề SKTT theo DISC-5 -> có thể đã được điều trị thành công, hoặc DISC-5/DSM-5 cần xem xét thêm cho bối cảnh VN.')

add_heading(doc, 'Hàm ý -- Sử dụng dịch vụ', 2)
add_p(doc, '- Cải thiện mental health literacy cho phụ huynh/gia đình: nhận biết dấu hiệu, biết tìm giúp đỡ ở đâu.')
add_p(doc, '- 56,2% sử dụng bác sĩ/y tá -> cơ hội tích hợp sàng lọc SKTT vào dịch vụ y tế tổng quát.')
add_p(doc, '- Áp dụng hướng dẫn mhGAP (WHO 2016) để tích hợp SKTT vào y tế ban đầu tại VN.')
add_p(doc, '- Cần chính sách công nhận rối loạn TT, hướng dẫn phân bổ nguồn lực cho SKTT VTN.')

# ============================================================
# COVID-19
# ============================================================
add_heading(doc, 'COVID-19', 1)
add_page_ref(doc, '24-27', JOURNAL)

add_heading(doc, 'Tổng quan', 2)
add_p(doc, 'Ca COVID-19 đầu tiên tại VN: tháng 1/2020. Phong tỏa toàn quốc từ 4/2020. Đỉnh dịch 9/2021 (>10.000 ca/ngày). Năm 2022 nới lỏng biện pháp. COVID-19 tác động đến VTN qua: cô lập xã hội, đóng cửa trường học, thiếu thói quen, căng thẳng gia đình.')

add_heading(doc, 'Đo lường', 2)
add_p(doc, 'Câu hỏi tập trung vào yếu tố liên quan SKTT: tiếp xúc COVID-19, cách ly, kỳ thị, tác động kinh tế, thay đổi sử dụng rượu/chất gây nghiện, vấn đề cảm xúc/hành vi tăng trong đại dịch, ngừng đi học.')

add_heading(doc, 'Kết quả', 2)
add_heading(doc, 'Bảng 17. Tỷ lệ "rất đồng ý" thường xuyên gặp vấn đề nhiều hơn trong COVID-19', 3)
add_page_ref(doc, '25', JOURNAL)
add_table(doc,
    ['Vấn đề', 'Nam %', 'Nam n', 'Nữ %', 'Nữ n', 'Tổng %', 'Tổng n'],
    [
        ['Lo âu/căng thẳng hơn', '5,2', '162', '4,9', '141', '5,1', '304'],
        ['Buồn/trầm cảm hơn', '3,6', '112', '3,7', '106', '3,6', '218'],
        ['Khó tập trung hơn', '2,7', '85', '1,9', '55', '2,3', '140'],
        ['Cô đơn/cô lập hơn', '2,2', '70', '1,5', '42', '1,9', '112'],
        ['TỔNG (bất kỳ vấn đề nào)', '8,2', '254', '7,3', '209', '7,7', '463'],
    ],
    widths=[3.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5]
)
add_p(doc, '7,7% VTN "rất đồng ý" thường xuyên gặp >= 1 vấn đề nhiều hơn bình thường trong COVID-19. Nếu tính cả "đồng ý": lên tới 67%.', size=11, bold=True)

add_p(doc, 'Các phát hiện khác:', size=12, bold=True)
add_p(doc, '- 7,1% (n=427) phụ huynh báo cáo con cần giúp đỡ trong COVID-19; 80,3% trong số đó không tiếp cận dịch vụ, chủ yếu do sợ nhiễm COVID (69,2%).')
add_p(doc, '- 71,5% phụ huynh báo cáo thu nhập hộ gia đình giảm trong COVID-19.')
add_p(doc, '- 12,3% thường xuyên không có đủ tiền cho nhu cầu thiết yếu.')
add_p(doc, '- <1% VTN báo cáo tăng sử dụng rượu/chất gây nghiện của người lớn trong hộ gia đình.')

add_image(doc, 'VNAMHS2022_p34_img1.jpeg', 12, 'Hình 4. Trải nghiệm trong đại dịch COVID-19 ở nhóm 10-17 tuổi (Trang 26)')

add_heading(doc, 'Diễn giải', 2)
add_p(doc, '1/13 VTN (7,7%) báo cáo tăng vấn đề cảm xúc/hành vi. Nhất quán với UNICEF (2020a). Tuy nhiên rất ít nghiên cứu định lượng quy mô lớn tại VN. 80,3% phụ huynh nhận ra nhu cầu nhưng không tiếp cận dịch vụ, chủ yếu do sợ COVID.')

add_heading(doc, 'Hạn chế', 2)
add_p(doc, '- Câu hỏi COVID-19 không nhằm đo lường toàn diện mọi khía cạnh đại dịch, chỉ cung cấp "bức tranh tổng quát".')
add_p(doc, '- Tỷ lệ phản hồi 81,1% cho thấy nỗi sợ COVID-19 không ảnh hưởng đáng kể đến sự tham gia.')

add_heading(doc, 'Hàm ý', 2)
add_p(doc, '- Cần cải thiện chính sách/dịch vụ hỗ trợ VTN hiện tại và chuẩn bị cho khủng hoảng tương lai.')
add_p(doc, '- Dịch vụ dễ tiếp cận (đường dây nóng, trò chuyện trực tuyến) cho chăm sóc khủng hoảng.')
add_p(doc, '- Nâng cao sức khỏe có mục tiêu: giảm cô đơn, trầm cảm liên quan đóng cửa trường học.')
add_p(doc, '- Đưa SKTT vào kế hoạch ứng phó đại dịch, thiên tai, xung đột trong tương lai.')

# ============================================================
# PHỤ LỤC (tóm tắt)
# ============================================================
add_heading(doc, 'PHỤ LỤC (Tóm tắt)', 1)
add_page_ref(doc, '28-44', JOURNAL)

add_heading(doc, 'Phụ lục 1: Công cụ đo lường', 2)
add_p(doc, 'Phụ huynh: Nhân khẩu học, Bệnh mạn tính, PSC-17, PHQ-9, GAD-7, DISC-5 (giới thiệu + ADHD), Sử dụng dịch vụ, COVID-19.')
add_p(doc, 'Vị thành niên: DISC-5 (giới thiệu + 5 rối loạn), Tự gây thương tích, Hỗ trợ/tự giúp, Sức khỏe tự đánh giá, Hoạt động thể chất, Rosenberg Self-Esteem, Bắt nạt, Trường học, Bạn bè/cô đơn, GEAS Family Connectedness, Tín ngưỡng, An toàn, Hành vi tình dục*, ACEs*, Sử dụng chất*, COVID-19.')
add_p(doc, '* Các phần do VTN tự thực hiện.', size=10, italic=True)

add_heading(doc, 'Phụ lục 2: Phương pháp nghiên cứu', 2)
add_p(doc, 'Khung mẫu: 63 tỉnh/thành -> 38 tỉnh được chọn, chia 4 vùng. 200 EA (50 EA/vùng: 25 thành thị + 25 nông thôn). 38 hộ/EA -> 7.600 hộ -> ước tính 6.000 hoàn thành (tỷ lệ phản hồi ước tính 79%).')

add_heading(doc, 'Bảng 18. Dân số VN 2019 theo vùng', 3)
add_table(doc,
    ['Vùng', 'Dân số', '% tổng', '% thành thị'],
    [
        ['V1: Trung du miền núi phía Bắc + Tây Nguyên', '18.375.547', '19,1', '21,5'],
        ['V2: Đồng bằng sông Hồng', '22.543.607', '23,4', '34,9'],
        ['V3: Bắc Trung Bộ + Duyên hải miền Trung', '20.187.293', '21,0', '28,3'],
        ['V4: Đông Nam Bộ + ĐBSCL', '35.102.537', '36,5', '44,2'],
        ['Tổng', '96.208.984', '100,0', '34,4'],
    ],
    widths=[6, 3, 2, 3]
)

add_p(doc, '38 tỉnh được chọn:', size=11, bold=True)
add_p(doc, 'Vùng 1: Hà Giang, Cao Bằng, Bắc Kạn, Tuyên Quang, Điện Biên, Yên Bái, Thái Nguyên, Bắc Giang, Kon Tum, Đắk Lắk, Lâm Đồng.', size=10)
add_p(doc, 'Vùng 2: Hà Nội, Quảng Ninh, Vĩnh Phúc, Hải Dương, Hải Phòng, Hưng Yên, Thái Bình, Nam Định.', size=10)
add_p(doc, 'Vùng 3: Thanh Hoá, Nghệ An, Quảng Bình, Đà Nẵng, Quảng Ngãi, Phú Yên, Ninh Thuận.', size=10)
add_p(doc, 'Vùng 4: Bình Phước, Bình Dương, Đồng Nai, Bà Rịa-Vũng Tàu, TP.HCM, Tiền Giang, Trà Vinh, Đồng Tháp, An Giang, Kiên Giang, Hậu Giang, Bạc Liêu.', size=10)

add_image(doc, 'VNAMHS2022_p41_img1.jpeg', 10, 'Hình 5. Bản đồ Việt Nam và 38 tỉnh/thành phố được chọn (Trang 31)')

add_p(doc, 'Nghiên cứu thí điểm: tại Hà Nội (25 VTN) và Đồng Nai (25 VTN). Kiểm tra nhất quán câu hỏi, thời gian phỏng vấn, tính phù hợp ngôn ngữ.')
add_p(doc, 'Công tác thực địa: Bắt đầu 21/9/2021 từ miền Bắc đến miền Nam. 127 điều tra viên (nhân viên POPFP địa phương). Thu thập dữ liệu chỉ khi EA phân loại "xanh" hoặc "vàng" về COVID.')

add_heading(doc, 'Phụ lục 3: Bảng thuật ngữ (Tóm tắt)', 2)
add_p(doc, 'Các thuật ngữ chính đã được giải thích trong nội dung chính. Xem bảng viết tắt ở đầu tài liệu.', size=11)

add_heading(doc, 'Phụ lục 4: Nhóm nghiên cứu', 2)
add_table(doc,
    ['Nhóm', 'Thành viên chính', 'Vai trò'],
    [
        ['V-NAMHS (IOS)', 'PGS. Vũ Mạnh Lợi', 'Nghiên cứu viên chính'],
        ['V-NAMHS (IOS)', 'PGS. Nguyễn Đức Vinh', 'Giám đốc Dự án'],
        ['V-NAMHS (IOS)', 'Đào Thị Khánh Hòa', 'Nghiên cứu viên cao cấp'],
        ['GOPFP', 'TS. Phạm Vũ Hoàng', 'Trưởng nhóm GOPFP'],
        ['UQ', 'TS. Holly Erskine', 'Nghiên cứu viên chính NAMHS'],
        ['UQ', 'GS. Harvey Whiteford', 'Cố vấn cao cấp'],
        ['UQ', 'GS. James Scott', 'Cố vấn lâm sàng'],
        ['JHSPH', 'GS. Robert Blum', 'Trưởng dự án NAMHS JHSPH'],
        ['JHSPH', 'TS. Shoshanna Fine', 'Nhà khoa học trợ lý'],
    ],
    widths=[3, 5, 6]
)

# ============================================================
# ĐÁNH GIÁ PHẢN BIỆN (RED CRITIQUE)
# ============================================================
add_heading(doc, 'ĐÁNH GIÁ PHẢN BIỆN', 1)

add_red_heading(doc, 'ĐIỂM MẠNH')
add_red(doc, '1. Mẫu đại diện quốc gia lớn nhất: N = 5.996, 38 tỉnh, tỷ lệ phản hồi cao (81,1%).')
add_red(doc, '2. Sử dụng công cụ chẩn đoán chuẩn hóa (DISC-5/DSM-5) thay vì thang triệu chứng -- khảo sát ĐẦU TIÊN làm điều này trên mẫu đại diện quốc gia tại Việt Nam.')
add_red(doc, '3. Phân biệt rõ ràng "vấn đề SKTT" (subthreshold) vs "rối loạn tâm thần" (full threshold) -- quan trọng cho chính sách.')
add_red(doc, '4. Đo lường toàn diện: SKTT + sử dụng dịch vụ + rào cản + nhu cầu nhận thức + COVID-19 + yếu tố nguy cơ/bảo vệ.')
add_red(doc, '5. Phỏng vấn riêng biệt phụ huynh và VTN trong môi trường riêng tư.')
add_red(doc, '6. Nỗ lực điều chỉnh văn hóa đáng kể: dịch/dịch ngược, đánh giá bác sĩ lâm sàng VN, nghiên cứu thí điểm.')
add_red(doc, '7. So sánh quốc tế: NAMHS triển khai đồng thời tại Kenya, Indonesia, VN -> cho phép so sánh xuyên quốc gia.')
add_red(doc, '8. Cung cấp dữ liệu nền tảng (baseline) cho GBD và các ước tính tỷ lệ hiện mắc toàn cầu.')

add_red_heading(doc, 'HẠN CHẾ')
add_red(doc, '1. Phỏng vấn trực tiếp (face-to-face): kỳ thị có thể ảnh hưởng sự tiết lộ, đặc biệt với hành vi tự tử/tự gây thương tích -> tỷ lệ thấp hơn so với GSHS (tự điền).')
add_red(doc, '2. DSM-5 dựa trên tiêu chí phương Tây: có thể không tính đến triệu chứng cơ thể (somatic symptoms) đặc trưng văn hóa VN (Kim et al. 2019).')
add_red(doc, '3. Chỉ đo 6 rối loạn: thiếu rối loạn ăn uống, OCD, rối loạn phân ly, rối loạn nhân cách.')
add_red(doc, '4. Tuổi 10-17: loại trừ 18-19 tuổi (WHO định nghĩa VTN = 10-19), bỏ sót giai đoạn chuyển tiếp quan trọng.')
add_red(doc, '5. Câu hỏi sử dụng dịch vụ chủ yếu hỏi phụ huynh, không phải VTN -> thiếu quan điểm trực tiếp của người trẻ về rào cản.')
add_red(doc, '6. Thiết kế cắt ngang (cross-sectional): không thể xác định quan hệ nhân quả hoặc thay đổi theo thời gian.')
add_red(doc, '7. COVID-19 module là add-on, không được thiết kế ban đầu: chỉ cung cấp "bức tranh tổng quát", không toàn diện.')
add_red(doc, '8. Thu thập dữ liệu trùng với đỉnh dịch COVID-19 (9-12/2021): có thể ảnh hưởng đến kết quả SKTT nhưng không thể tách biệt tác động.')
add_red(doc, '9. 71,9% người tiếp cận dịch vụ không có vấn đề SKTT theo DISC-5 -> câu hỏi về tính giá trị của thuật toán chấm điểm DSM-5 trong bối cảnh VN.')

add_red_heading(doc, 'KHOẢNG TRỐNG NGHIÊN CỨU (GAPS)')
add_red(doc, '1. Thiếu dữ liệu dọc (longitudinal): cần theo dõi mẫu V-NAMHS để đánh giá thay đổi theo thời gian và hậu quả dài hạn.')
add_red(doc, '2. Thiếu so sánh nhóm dân tộc thiểu số: VN có 54 dân tộc, báo cáo không phân tích theo dân tộc.')
add_red(doc, '3. Thiếu phân tích theo khu vực thành thị-nông thôn: mặc dù mẫu bao gồm cả hai, kết quả không phân tích riêng.')
add_red(doc, '4. Thiếu phân tích yếu tố nguy cơ/bảo vệ chi tiết: ACEs, bắt nạt, sử dụng chất, lòng tự trọng đã đo nhưng không trình bày trong báo cáo chính.')
add_red(doc, '5. Cần phát triển và chuẩn hóa thuật toán chấm điểm DISC-5 cho bối cảnh VN: xem xét triệu chứng cơ thể, khác biệt văn hóa.')
add_red(doc, '6. Thiếu can thiệp: V-NAMHS cung cấp dữ liệu mô tả nhưng chưa có chương trình can thiệp dựa trên kết quả.')
add_red(doc, '7. Cần nghiên cứu về hiểu biết SKTT (mental health literacy) và kỳ thị ở cả VTN lẫn phụ huynh VN.')
add_red(doc, '8. Thiếu dữ liệu về chi phí kinh tế của rối loạn TT ở VTN VN để hỗ trợ vận động chính sách.')
add_red(doc, '9. Cần lặp lại khảo sát sau COVID-19 để đánh giá tác động dài hạn của đại dịch lên SKTT VTN.')

# ============================================================
# Cuối
# ============================================================
add_image(doc, 'VNAMHS2022_p51_img1.jpeg', 10, 'Trang cuối báo cáo V-NAMHS')

add_p(doc, '')
add_p(doc, '--- Hết bản dịch ---', size=11, italic=True)

# Save
doc.save(OUT)
print(f'DONE: {OUT}')
