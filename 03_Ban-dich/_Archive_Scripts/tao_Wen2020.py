# -*- coding: utf-8 -*-
"""Dịch chi tiết Wen et al. 2020 — format Jenkins
Lý do: Nam > Nữ về lo âu (trái y văn quốc tế)"""
import sys, io, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()

# 1. LINK + QR
add_link_and_qr(doc, 'https://doi.org/10.3390/ijerph17114079', 'QR_Wen2020.png')

# 2. TIÊU ĐỀ
add_heading(doc, 'Phân tích hồ sơ tiềm ẩn về lo âu ở học sinh trung học cơ sở tại các vùng nông thôn kém phát triển của Trung Quốc', 1)
add_heading(doc, 'A Latent Profile Analysis of Anxiety among Junior High School Students in Less Developed Rural Regions of China', 2)

# 3. THÔNG TIN THƯ MỤC
add_heading(doc, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(doc, [
    ('Tiêu đề gốc', 'A Latent Profile Analysis of Anxiety among Junior High School Students in Less Developed Rural Regions of China'),
    ('Tiêu đề dịch', 'Phân tích hồ sơ tiềm ẩn về lo âu ở học sinh trung học cơ sở tại các vùng nông thôn kém phát triển của Trung Quốc'),
    ('Tác giả', 'Xiaotong Wen, Yixiang Lin, Yuchen Liu, Katie Starcevich, Fang Yuan, Xiuzhu Wang, Xiaoxu Xie, Zhaokang Yuan'),
    ('Cơ quan', 'Trường Y tế Công cộng, Đại học Nam Xương, Giang Tây; ĐH Carnegie Mellon; ĐH Nevada, Reno; ĐH Hawaii tại Mānoa; ĐH Y khoa Phúc Kiến'),
    ('Tạp chí', 'International Journal of Environmental Research and Public Health, 2020, 17(11):4079'),
    ('DOI', '10.3390/ijerph17114079'),
    ('PMID', '32521646 | PMCID: PMC7312008'),
])

# 4. TÓM TẮT
add_page_ref(doc, '1/14', 'Int J Environ Res Public Health', '17(11):4079, 2020')

add_heading(doc, 'TÓM TẮT', 2)
add_p(doc, 'Mục đích: Tìm hiểu các loại lo âu tiềm ẩn ở học sinh trung học cơ sở (THCS) bằng cách phân tích tình trạng lo âu hiện tại và các yếu tố ảnh hưởng tại vùng nông thôn kém phát triển tỉnh Giang Tây, Trung Quốc.')

add_red(doc, 'Phương pháp: Lấy mẫu cụm ngẫu nhiên phân tầng nhiều giai đoạn. 900 học sinh THCS từ lớp 9 đến lớp 12 tại 6 huyện nông thôn. Công cụ: Trắc nghiệm Sức khỏe Tâm thần (MHT — 100 mục, 8 thang nội dung). Phân tích hồ sơ tiềm ẩn (LPA) bằng Mplus 7.4.')
add_red(doc, 'Kết quả: LPA xác định 3 hồ sơ: lo âu nhẹ 19,22%, lo âu trung bình 56,00%, lo âu nặng 24,78%. NAM có nhiều khả năng phát triển lo âu trung bình và nặng HƠN NỮ (OR lo âu trung bình = 0,649 cho nam, tức nữ cao hơn; OR lo âu nặng = 0,262 cho nam, tức NỮ cao gấp gần 4 lần — ĐÂY LÀ TRÁI NGƯỢC: thực tế nữ > nam trong nghiên cứu này). Áp lực học tập rất cao: OR = 11,579 cho lo âu nặng. Hỗ trợ SKTT tại trường đầy đủ: OR = 0,562 (bảo vệ).')

add_p(doc, 'Kết luận: Cần thiết lập dịch vụ tư vấn SKTT tại trường học và giảm áp lực học tập để phòng ngừa lo âu nghiêm trọng.')

# TÓM TẮT ĐÁNH GIÁ NHANH
add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(doc, 'Điểm nổi bật:', bold=True)
add_p(doc, '• Phương pháp LPA (Latent Profile Analysis) độc đáo — tiếp cận tập trung vào cá nhân (person-centered) thay vì biến.')
add_p(doc, '• Áp lực học tập rất cao là yếu tố nguy cơ mạnh nhất (OR = 11,579 cho lo âu nặng) — phản ánh đặc thù hệ thống giáo dục Trung Quốc.')
add_p(doc, '• Hỗ trợ SKTT tại trường là yếu tố bảo vệ (OR = 0,562) — bằng chứng cho can thiệp dựa trên trường học.')

add_p(doc, 'Hạn chế:', bold=True)
add_p(doc, '• Cỡ mẫu 900 — khiêm tốn, chỉ nông thôn 1 tỉnh.')
add_p(doc, '• MHT là thang đo đặc thù Trung Quốc, khó so sánh quốc tế (không dùng GAD-7/DASS-21).')
add_p(doc, '• Thiết kế cắt ngang — không thể suy luận nhân quả.')

add_p(doc, 'Lưu ý về giới tính:', bold=True)
add_p(doc, '• Trong nghiên cứu này, NỮ thực tế lo âu hơn nam (nữ lo âu nặng 63,23% so với nam 36,77%). OR = 0,262 cho nam nghĩa là nam có nguy cơ lo âu nặng chỉ bằng 26,2% so với nữ — tức NỮ cao gấp gần 4 lần. Đây KHÔNG phải nghiên cứu "nam > nữ" như ban đầu tưởng. Cần phân biệt: trong nhóm lo âu nặng, nữ chiếm 63,23% nhưng trong tổng mẫu, tỷ lệ phần trăm nam có lo âu cũng cao.')

# NỘI DUNG DỊCH ĐẦY ĐỦ
add_heading(doc, '1. GIỚI THIỆU', 2)
add_page_ref(doc, '1–2/14', 'Int J Environ Res Public Health', '17(11):4079, 2020')
add_p(doc, 'Trung học cơ sở là giai đoạn quan trọng cho sự trưởng thành và phát triển cá nhân, đồng thời cũng là giai đoạn hình thành nhân cách và giá trị. Tuy nhiên, gia đình và nhà trường chủ yếu tập trung vào phát triển thể chất, thành tích học tập và hành vi, trong khi sức khỏe tâm thần và cảm xúc của học sinh bị bỏ qua nghiêm trọng.')
add_p(doc, 'Hệ thống giáo dục Trung Quốc bị chỉ trích rộng rãi vì tập trung quá mức vào học thuật, tạo ra mức căng thẳng độc hại, bỏ qua sức khỏe cảm xúc. Hệ quả: tỷ lệ lo âu cao ở học sinh. Một tổng quan hệ thống tại 44 quốc gia ước tính tỷ lệ lo âu nặng toàn cầu là 7,3%. Tại Mỹ: 31%, châu Âu: 14%.')
add_p(doc, 'Hầu hết nghiên cứu trong nước tập trung vào khu vực thành thị. Thanh thiếu niên nông thôn cần được quan tâm nhiều hơn. Nghiên cứu này khảo sát học sinh THCS nông thôn tỉnh Giang Tây.')
add_p(doc, 'Bốn giả thuyết: (1) 3 phân loại tiềm ẩn; (2) SKTT trường đầy đủ → lo âu thấp; (3) Thành tích cao → lo âu cao hơn; (4) Áp lực học tập cao → lo âu cao hơn.')

add_heading(doc, '2. PHƯƠNG PHÁP', 2)
add_page_ref(doc, '3–6/14', 'Int J Environ Res Public Health', '17(11):4079, 2020')
add_p(doc, '2.1. Đối tượng và lấy mẫu:', bold=True)
add_p(doc, '900 học sinh THCS tại 6 huyện nông thôn tỉnh Giang Tây: Vu Đô, Thượng Nhiêu, Đô Xương, Phong Thành, Đông Hương, Toại Xuyên. Tuổi trung bình: 14,14 ± 1,32. Lấy mẫu cụm ngẫu nhiên phân tầng nhiều giai đoạn: Ủy ban Y tế tỉnh chọn huyện → chọn xã → chọn trường → chọn ngẫu nhiên lớp từ mỗi khối.')
add_p(doc, '2.2. Công cụ:', bold=True)
add_p(doc, 'Trắc nghiệm Sức khỏe Tâm thần (MHT — Mental Health Test): 100 câu hỏi, 8 thang nội dung + 1 thang hiệu lực. Thang đo lo âu tiêu chuẩn hóa cho THCS tại Trung Quốc, chuyển thể từ bản Nhật Bản (Kiyoshi Suzuki) bởi Giáo sư Bucheng Zhou (ĐH Sư phạm Hoa Đông, 1991). Cronbach alpha: 0,878. Độ tin cậy: 0,84–0,88. Kiểm tra lại: 0,78–0,86.')
add_p(doc, '8 thang nội dung: (1) Lo âu học tập, (2) Lo âu quan hệ, (3) Xu hướng cô đơn, (4) Xu hướng tự trách, (5) Xu hướng dị ứng, (6) Triệu chứng thể chất, (7) Xu hướng sợ hãi, (8) Xu hướng bốc đồng.')
add_p(doc, '2.3. Phân tích:', bold=True)
add_p(doc, 'Phân tích hồ sơ tiềm ẩn (LPA) bằng Mplus 7.4. Đánh giá: AIC, BIC, ABIC, entropy, LMRT, VLMR, BLRT. Hồi quy logistic đa biến bằng SPSS 24.0.')

add_heading(doc, '3. KẾT QUẢ', 2)
add_page_ref(doc, '7–10/14', 'Int J Environ Res Public Health', '17(11):4079, 2020')

add_p(doc, '3.1. Phân tích hồ sơ tiềm ẩn (LPA):', bold=True)
add_p(doc, 'Giải pháp 3 lớp phù hợp nhất (entropy = 0,816, LMRT/VLMR/BLRT P < 0,001):')
add_table(doc,
    ['Hồ sơ', 'n', 'Tỷ lệ', 'Đặc điểm'],
    [['Lớp 1: Lo âu nhẹ', '173', '19,22%', 'Điểm thấp nhất trên 8 yếu tố'],
     ['Lớp 2: Lo âu trung bình', '504', '56,00%', 'Điểm trung bình'],
     ['Lớp 3: Lo âu nặng', '223', '24,78%', 'Điểm cao nhất trên 8 yếu tố']],
    widths=[3.5, 1.5, 2.0, 5.0])

doc.add_paragraph()
add_p(doc, '3.2. Phân bố giới tính trong nhóm lo âu:', bold=True)
add_table(doc,
    ['Giới tính', 'Lo âu nhẹ', 'Lo âu trung bình', 'Lo âu nặng'],
    [['Nam (n=467, 51,89%)', '110 (63,58%)', '275 (54,56%)', '82 (36,77%)'],
     ['Nữ (n=433, 48,11%)', '63 (36,42%)', '229 (45,44%)', '141 (63,23%)']],
    widths=[4.0, 3.0, 3.0, 3.0])

doc.add_paragraph()
add_p(doc, '⚠️ Nữ chiếm 63,23% nhóm lo âu nặng (so với nam 36,77%). Giới tính có ý nghĩa thống kê (χ² = 31,337, P < 0,001).', bold=True)

doc.add_paragraph()
add_p(doc, '3.3. Hồi quy logistic đa biến (nhóm tham chiếu: lo âu nhẹ):', bold=True)
add_table(doc,
    ['Biến', 'Lo âu TB (OR)', 'Lo âu nặng (OR)'],
    [['Giới tính nam (vs nữ)', '0,649*', '0,262* (nam = bảo vệ)'],
     ['Áp lực học tập rất cao (vs rất thấp)', '6,523*', '11,579*'],
     ['Áp lực học tập cao (vs rất thấp)', '6,122*', '5,894*'],
     ['SKTT trường đầy đủ (vs không chắc)', '—', '0,562* (bảo vệ)'],
     ['Thành tích xuất sắc (vs không đạt)', '0,377*', '—'],
     ['Thành tích tốt (vs không đạt)', '—', '0,362*'],
     ['Tuổi 15 (vs 12)', '—', '0,461*']],
    widths=[5.5, 3.0, 3.5])
add_p(doc, '* P < 0,05', italic=True, size=10)

add_heading(doc, '4. THẢO LUẬN', 2)
add_page_ref(doc, '11–12/14', 'Int J Environ Res Public Health', '17(11):4079, 2020')
add_p(doc, 'Lo âu nặng chiếm 24,78% — thấp hơn một báo cáo khác (35%) nhưng vẫn đáng lo ngại. Phân tích hồi quy cho thấy nam có nguy cơ thấp hơn nữ (OR lo âu nặng = 0,262), có thể do nữ nhạy cảm hơn với môi trường và biến động cảm xúc, vai trò hormone sinh dục trong điều hòa cảm xúc.')
add_p(doc, 'Áp lực học tập rất cao là yếu tố nguy cơ mạnh nhất (OR = 11,579 cho lo âu nặng so với áp lực rất thấp). Đây là hệ quả của hệ thống giáo dục Trung Quốc tập trung vào thi cử.')
add_p(doc, 'Hỗ trợ SKTT tại trường đầy đủ giảm nguy cơ lo âu nặng (OR = 0,562). Khuyến nghị: thiết lập phòng tư vấn tâm lý chuyên nghiệp tại trường.')
add_p(doc, 'Hạn chế: Thiết kế cắt ngang, tự báo cáo, chỉ thanh thiếu niên đang đi học.')

add_heading(doc, 'TÀI LIỆU THAM KHẢO', 2)
add_p(doc, '[Giữ nguyên 53 tài liệu tham khảo tiếng Anh từ bài gốc — xem bản PDF gốc]', italic=True)

# BẢNG VIẾT TẮT
add_abbreviation_table(doc, [
    ('MHT', 'Mental Health Test — Trắc nghiệm Sức khỏe Tâm thần'),
    ('LPA', 'Latent Profile Analysis — Phân tích Hồ sơ Tiềm ẩn'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('THCS', 'Trung học cơ sở'),
    ('OR', 'Odds Ratio — Tỷ số chênh'),
    ('AIC', 'Akaike Information Criterion — Tiêu chí thông tin Akaike'),
    ('BIC', 'Bayesian Information Criterion — Tiêu chí thông tin Bayes'),
    ('LBC', 'Left-Behind Children — Trẻ em bị bỏ lại'),
    ('IJERPH', 'International Journal of Environmental Research and Public Health'),
])

# PHẢN BIỆN
add_red_heading(doc, 'PHẢN BIỆN CHI TIẾT')

add_red(doc, 'Điểm mạnh:', bold=True)
add_red(doc, '1. Phương pháp LPA độc đáo, tập trung vào cá nhân — ít nghiên cứu lo âu thanh thiếu niên sử dụng.')
add_red(doc, '2. Xác định 3 hồ sơ lo âu khách quan bằng tiêu chuẩn thống kê (BIC, entropy, BLRT).')
add_red(doc, '3. Phát hiện áp lực học tập (OR = 11,579) — dữ liệu mạnh cho chính sách giáo dục.')

add_red(doc, 'Hạn chế:', bold=True)
add_red(doc, '1. Cỡ mẫu 900 khiêm tốn, chỉ nông thôn 1 tỉnh — hạn chế tính khái quát.')
add_red(doc, '2. MHT là thang đo đặc thù Trung Quốc (1991), khó so sánh quốc tế. Không sử dụng GAD-7 hay DASS-21.')
add_red(doc, '3. Không có phân tích đa biến kiểm soát đồng thời tất cả biến (chỉ hồi quy logistic với biến giả).')
add_red(doc, '4. Tiêu đề "lớp 9 đến lớp 12" nhưng THCS Trung Quốc chỉ đến lớp 9 — không nhất quán.')
add_red(doc, '5. Tạp chí IJERPH đã bị loại khỏi SCIE vào năm 2023 do lo ngại chất lượng.')

add_red(doc, 'Khoảng trống nghiên cứu (Gap):', bold=True)
add_red(doc, '1. Nghiên cứu dọc theo dõi lo âu từ THCS đến THPT.')
add_red(doc, '2. So sánh nông thôn-thành thị cùng tỉnh bằng cùng công cụ.')
add_red(doc, '3. Can thiệp giảm áp lực học tập tại trường nông thôn và đánh giá hiệu quả.')
add_red(doc, '4. Sử dụng GAD-7 song song MHT để cho phép so sánh quốc tế.')

fname = '12_Wen_2020_IJERPH.docx'
doc.save(fname)
sys.stderr.write(f'{fname} saved OK\n')
