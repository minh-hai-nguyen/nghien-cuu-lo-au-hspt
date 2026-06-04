# -*- coding: utf-8 -*-
"""Dịch toàn văn QT21 - Norway 2025 - Brunborg et al. - Social Science & Medicine"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()

# ========== 1. LINK + QR ==========
add_p(doc, 'Link bài báo gốc: https://doi.org/10.1016/j.socscimed.2025.118528', size=10)

# ========== 2. TIÊU ĐỀ ==========
add_heading(doc, 'Giải thích khả thi cho xu hướng tăng căng thẳng tâm thần ở thanh thiếu niên Na Uy từ 2011 đến 2024', 1)
h2 = doc.add_paragraph()
r = h2.add_run('Possible Explanations for the Upward Trend in Mental Distress among Adolescents in Norway from 2011 to 2024')
r.font.name = 'Times New Roman'; r.font.size = Pt(13); r.italic = True

# ========== 3. THÔNG TIN THƯ MỤC ==========
add_heading(doc, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(doc, [
    ('Tiêu đề gốc', 'Possible Explanations for the Upward Trend in Mental Distress among Adolescents in Norway from 2011 to 2024'),
    ('Tiêu đề dịch', 'Giải thích khả thi cho xu hướng tăng căng thẳng tâm thần ở thanh thiếu niên Na Uy từ 2011 đến 2024'),
    ('Tác giả', 'Geir Scott Brunborg, Sondre Aasen Nilsen, Jens Christoffer Skogen, Lasse Bang'),
    ('Cơ quan', 'Viện Y tế Công cộng Na Uy (NIPH), Oslo; Viện Karolinska, Stockholm; NORCE, Na Uy'),
    ('Tạp chí', 'Social Science & Medicine (Q1, IF = 5,4)'),
    ('Thông tin xuất bản', 'Vol. 384 (2025), Article 118528, 10 trang'),
    ('DOI', '10.1016/j.socscimed.2025.118528'),
    ('Loại nghiên cứu', 'Phân tích xu hướng — cắt ngang lặp lại (repeated cross-sectional), decomposition'),
    ('Mẫu', '979.043 thanh thiếu niên Na Uy, 1.417 khảo sát cấp thành phố (2011–2024)'),
])

# ========== 4. PAGE REF ==========
add_page_ref(doc, '1–10', 'Social Science & Medicine', 'Vol. 384, 2025')

# ========== 5. TÓM TẮT (Abstract) ==========
add_heading(doc, 'TÓM TẮT', 2)

add_p(doc, 'Bối cảnh: Xu hướng tăng căng thẳng tâm thần (mental distress) tự báo cáo ở thanh thiếu niên đã được ghi nhận tại Na Uy và nhiều quốc gia khác, nhưng nguyên nhân vẫn chưa rõ. Nghiên cứu này nhằm xác định các giải thích tiềm năng cho xu hướng này bằng cách kiểm tra các yếu tố giả thuyết sử dụng dữ liệu cắt ngang lặp lại.')

p = doc.add_paragraph()
r = p.add_run('Phương pháp: Phân tích phản hồi từ 979.043 thanh thiếu niên Na Uy, thu thập qua 1.417 khảo sát cấp thành phố từ 2011 đến 2024. Tám yếu tố giải thích có bằng chứng liên quan đến căng thẳng tâm thần được kiểm tra: khó khăn tài chính gia đình, tối ở nhà, thiếu vận động, bắt nạt, bất mãn với cha mẹ, sử dụng mạng xã hội, sử dụng cần sa, và bất mãn với trường học. Theo khung dịch tễ học đã công bố, đánh giá thay đổi về tỷ lệ, độc lực (virulence — sức mạnh liên quan), và mức độ mỗi yếu tố giải thích xu hướng.')
r.font.name = 'Times New Roman'; r.font.size = Pt(12)
r.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

p = doc.add_paragraph()
r = p.add_run('Kết quả: Căng thẳng tâm thần tăng đáng kể ở cả hai giới theo thời gian. Bất mãn với trường học tăng chiếm phần lớn xu hướng tăng ở cả nam và nữ. Thời gian dùng mạng xã hội tăng cũng giải thích một phần đáng kể, đặc biệt ở nữ. Tối ở nhà tăng cũng đóng góp nhưng nhỏ hơn. Sử dụng cần sa tăng chỉ đóng góp nhỏ ở nữ. Khó khăn tài chính, thiếu vận động, bắt nạt và bất mãn với cha mẹ không giải thích xu hướng tăng.')
r.font.name = 'Times New Roman'; r.font.size = Pt(12)
r.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

add_p(doc, 'Kết luận: Bất mãn với trường học tăng, sử dụng mạng xã hội tăng, tối ở nhà nhiều hơn, và sử dụng cần sa được xác định là yếu tố giải thích tiềm năng, mặc dù tác động ước tính khác nhau. Kết quả nhấn mạnh tính phức tạp của xu hướng sức khỏe tâm thần và tầm quan trọng của việc xem xét nhiều giải thích.')

add_p(doc, 'Từ khóa: Căng thẳng tâm thần, sức khỏe tâm thần, xu hướng, tăng, giải thích, thanh thiếu niên, giới trẻ.')

# ========== 6. TÓM TẮT ĐÁNH GIÁ NHANH ==========
add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
bullets = [
    'Xu hướng 13 năm (2011–2024) với mẫu cực lớn (n = 979.043): căng thẳng tâm thần tăng ở CẢ nam và nữ VTN Na Uy.',
    'Bất mãn với trường học là yếu tố giải thích LỚN NHẤT — giữ cố định yếu tố này loại bỏ hoàn toàn xu hướng tăng ở nam, và giảm đáng kể ở nữ.',
    'Thời gian dùng mạng xã hội cũng giải thích một phần đáng kể, phù hợp với Nature 2025 và Chen 2023.',
    'Sử dụng khung dịch tễ học Keyes & Platt (2024) — phương pháp phân tích decomposition nâng cao.',
    'Social Science & Medicine Q1 IF = 5,4 — tạp chí uy tín cao trong lĩnh vực y tế công cộng.',
]
for b in bullets:
    add_p(doc, f'• {b}')

add_p(doc, 'Hạn chế:', bold=True)
limits = [
    'Chỉ dữ liệu Na Uy — không khái quát trực tiếp cho châu Á/Việt Nam.',
    'Đo "mental distress" chung (6 mục Hopkins), không tách lo âu riêng bằng GAD-7.',
    'Giả định chiều hướng nhân quả một chiều (yếu tố → distress) — hạn chế lớn nhất.',
    'Mạng xã hội chỉ đo thời gian (top: ≥3 giờ/ngày) — không đo nội dung, tương tác.',
]
for b in limits:
    add_p(doc, f'• {b}')

add_p(doc, 'Hướng cải thiện:', bold=True)
improve = [
    'Áp dụng phân tích decomposition cho dữ liệu VN (Hoàng Trung Học 2025 có 2 thời điểm).',
    'So sánh xu hướng Na Uy vs VN, châu Á — kiểm tra tính khái quát.',
    'Nghiên cứu dọc (longitudinal) theo dõi cùng cá nhân để xác nhận chiều nhân quả.',
]
for b in improve:
    add_p(doc, f'• {b}')

# ========== 7. GIỚI THIỆU ==========
add_page_ref(doc, '1–2', 'Social Science & Medicine', 'Vol. 384, 2025')
add_heading(doc, '1. GIỚI THIỆU', 2)

add_p(doc, 'Xu hướng tăng căng thẳng tâm thần tự báo cáo ở thanh thiếu niên đã được ghi nhận ở Na Uy và nhiều quốc gia khác kể từ khoảng năm 2010. Những xu hướng này đã trở thành chủ đề gây lo ngại trong chính sách y tế công cộng, và nhiều giải thích đã được đề xuất nhưng chưa được kiểm tra thực nghiệm đầy đủ.')

add_p(doc, 'Một số giả thuyết đã được đưa ra. Haidt (2024) cho rằng mạng xã hội đóng vai trò nhân quả, trong khi Twenge (2017) nhấn mạnh smartphone và thời gian trực tuyến. Theo "giả thuyết suy giảm phúc lợi trẻ em" (Eckersley, 2011), các yếu tố xã hội rộng hơn như thương mại hóa thời thơ ấu và áp lực thành tích đã ảnh hưởng tiêu cực. "Giả thuyết thế hệ bông" (Haidt & Lukianoff, 2018) cho rằng trẻ em được bảo vệ quá mức dẫn đến ít kiên cường hơn. Thêm vào đó, "giả thuyết lạm phát tỷ lệ" (Foulkes & Andrews, 2023) đề xuất rằng nhận thức tăng về sức khỏe tâm thần có thể dẫn đến tăng báo cáo triệu chứng do diễn giải quá mức trải nghiệm tâm lý thông thường.')

add_p(doc, 'Mặc dù được chú ý rộng rãi trên truyền thông, đáng ngạc nhiên là rất ít nghiên cứu thực nghiệm đã cố gắng giải thích trực tiếp xu hướng tăng căng thẳng tâm thần (Armitage et al., 2024). Xác định nguyên nhân xu hướng là thách thức lớn — thiết kế thực nghiệm thường không khả thi vì lý do thực tế và đạo đức.')

add_p(doc, 'Keyes và Platt (2024) gần đây đề xuất khung dịch tễ học với ba tiêu chí hữu ích để xác định giải thích hợp lý cho xu hướng tăng căng thẳng tâm thần:')
add_p(doc, '(1) Liên quan ở cấp cá nhân, lý tưởng là quan hệ nhân quả, giữa yếu tố giải thích và căng thẳng tâm thần.')
add_p(doc, '(2) Tỷ lệ của yếu tố giải thích phải tăng theo thời gian.')
add_p(doc, '(3) Sức mạnh liên quan (độc lực — virulence) phải tăng theo thời gian, nghĩa là cùng "liều lượng" yếu tố trở nên có hại hơn.')
add_p(doc, 'Do căng thẳng tâm thần tăng mạnh hơn ở nữ so với nam tại Mỹ, Keyes và Platt cũng thêm rằng tăng độc lực hoặc tỷ lệ nên rõ rệt hơn ở nữ.')

add_p(doc, 'Mục tiêu nghiên cứu này là kiểm tra xem xu hướng tăng căng thẳng tâm thần ở VTN Na Uy có thể quy cho các giải thích giả thuyết cụ thể hay không, và từ đó thu hẹp phạm vi giải thích hợp lý. Nghiên cứu phân tích phản hồi từ 979.043 VTN Na Uy, thu thập qua 1.417 khảo sát cấp thành phố từ 2011 đến 2024.')

# ========== 8. PHƯƠNG PHÁP ==========
add_page_ref(doc, '2–3', 'Social Science & Medicine', 'Vol. 384, 2025')
add_heading(doc, '2. PHƯƠNG PHÁP', 2)

add_p(doc, '2.1. Mẫu và quy trình', bold=True)
add_p(doc, 'Dữ liệu từ khảo sát Ungdata (Ungdata surveys) — hệ thống khảo sát VTN quốc gia Na Uy, do NOVA (Trung tâm Nghiên cứu Phúc lợi và Lao động Na Uy) quản lý. Khảo sát dựa trên web, thực hiện tại trường trong giờ học, ẩn danh. Tổng cộng 979.043 VTN từ 1.417 khảo sát cấp thành phố (2011–2024). Độ tuổi: chủ yếu lớp 8–10 (THCS) và năm 1 THPT (khoảng 13–16 tuổi).')

add_p(doc, '2.2. Thang đo', bold=True)
add_p(doc, 'Căng thẳng tâm thần đo bằng 6 mục rút gọn từ Thang Triệu chứng Hopkins (HSCL — Hopkins Symptom Checklist): "Cảm thấy mọi thứ là gánh nặng", "Khó ngủ", "Cảm thấy không hạnh phúc, buồn bã, hoặc trầm cảm", "Cảm thấy tuyệt vọng về tương lai", "Cảm thấy cứng đờ hoặc căng thẳng", "Lo lắng quá nhiều về mọi thứ". Trả lời: 1 = Hoàn toàn không, 2 = Một chút, 3 = Khá nhiều, 4 = Rất nhiều. Điểm trung bình 6 mục, dao động 1–4.')

add_p(doc, 'Tám yếu tố giải thích tiềm năng:', bold=True)
factors = [
    'Khó khăn tài chính: "Gia đình bạn có gặp khó khăn tài chính không?" (1–5: Không bao giờ → Rất thường xuyên).',
    'Tối ở nhà: "Tuần qua, bạn đã dành bao nhiêu buổi tối hoàn toàn ở nhà?" (1–4: Không → ≥6 lần).',
    'Hoạt động thể chất: 3 mục (CLB thể thao, phòng gym, hoạt động tổ chức khác), 6 mức (Không bao giờ → ≥5 lần/tuần).',
    'Bắt nạt: "Bạn có bị bắt nạt, đe dọa hoặc cô lập?" (1–6: Không bao giờ → Nhiều lần/tuần).',
    'Hài lòng với cha mẹ: "Bạn hạnh phúc hay không hạnh phúc với cha mẹ?" (1–5: Rất không → Rất hạnh phúc).',
    'Mạng xã hội: "Ngày bình thường, bạn dành bao lâu cho mạng xã hội (Facebook, Instagram...)?" (1–6: Không → >3 giờ). Có dữ liệu từ 2014.',
    'Sử dụng cần sa: "Bạn đã dùng cần sa bao nhiêu lần?" (1–6: Không bao giờ → >50 lần).',
    'Bất mãn với trường học: "Bạn cảm thấy thế nào về trường học?" (1–5: Rất thích → Rất không thích). Có dữ liệu từ 2014.',
]
for f in factors:
    add_p(doc, f'• {f}')

add_p(doc, '2.3. Phân tích', bold=True)
add_p(doc, 'Khảo sát được thực hiện trên Stata phiên bản 18 (StataCorp LLC, 2023). Kiểm tra thay đổi điểm trung bình theo thời gian bằng hồi quy tuyến tính, với năm là biến liên tục. Điều chỉnh cho khác biệt về lớp học và tính trung tâm địa lý. Ước tính mô hình hồi quy tuyến tính (linear regression) riêng cho nam và nữ.')

add_p(doc, 'Phân tích theo ba bước chính:')
add_p(doc, '(1) Xu hướng thay đổi tỷ lệ (prevalence) của các yếu tố giải thích theo thời gian.')
add_p(doc, '(2) Thay đổi độc lực (virulence) — sức mạnh liên quan giữa mỗi yếu tố và căng thẳng tâm thần theo thời gian (tương tác yếu tố × năm).')
add_p(doc, '(3) Decomposition — giữ cố định (holding constant) tác động của từng yếu tố theo thời gian, kiểm tra xem xu hướng tăng có bị giảm không.')
add_p(doc, 'Ngoài ra, mô hình hồi quy logistic cũng được chạy minh họa với biến nhị phân.')

# ========== Bảng 1: Thống kê mô tả ==========
add_page_ref(doc, '3', 'Social Science & Medicine', 'Vol. 384, 2025')
add_heading(doc, 'Bảng 1. Thống kê mô tả', 3)
add_table(doc,
    ['Biến số', 'Khoảng', 'Nam', 'Nữ', 'Giai đoạn dữ liệu'],
    [
        ['Căng thẳng tâm thần', '1–4', '1,64 (0,71) 443.676', '1,96 (0,80) 461.247', '2011–2024 excl. 2020'],
        ['Khó khăn tài chính', '1–5', '1,79 (0,80) 470.406', '1,93 (0,91) 474.709', '2011–2024 excl. 2020'],
        ['Tối ở nhà', '1–4', '3,09 (0,89) 482.228', '3,14 (0,82) 476.255', '2011–2024 excl. 2020'],
        ['Hoạt động thể chất', '1–6', '4,30 (1,55) 404.221', '4,13 (1,43) 407.078', '2011–2024 excl. 2020'],
        ['Bắt nạt', '1–6', '1,64 (1,18) 460.685', '1,63 (1,13) 477.215', '2011–2024 excl. 2020'],
        ['Hài lòng với cha mẹ', '1–5', '4,46 (0,83) 461.652', '4,37 (0,84) 462.536', '2011–2024 excl. 2020'],
        ['Mạng xã hội', '1–6', '3,63 (1,67) 459.946', '4,37 (1,34) 460.861', '2014–2024 excl. 2020'],
        ['Sử dụng cần sa', '1–6', '1,23 (0,76) 401.343', '1,12 (0,53) 474.291', '2011–2024 excl. 2020'],
        ['Bất mãn trường học', '1–5', '1,25 (0,74) 413.606', '1,09 (0,43) 413.868', '2014–2024 excl. 2020'],
    ],
    widths=[3.5, 1.5, 4.0, 4.0, 3.0]
)
add_p(doc, 'Ghi chú: Giá trị = Trung bình (Độ lệch chuẩn) Cỡ mẫu. Điều chỉnh cho lớp học và vùng địa lý. Excl. 2020 = không bao gồm 2020 (COVID, không khảo sát).', size=9, italic=True)

# ========== 9. KẾT QUẢ ==========
add_page_ref(doc, '3–6', 'Social Science & Medicine', 'Vol. 384, 2025')
add_heading(doc, '3. KẾT QUẢ', 2)

add_p(doc, '2.1. Xu hướng tổng thể căng thẳng tâm thần', bold=True)
add_p(doc, 'Điểm trung bình căng thẳng tâm thần tăng đáng kể theo thời gian ở cả nam và nữ. Hình 1 cho thấy xu hướng tăng liên tục từ 2011 đến 2024. Với cả hai giới, ước tính thay đổi tuyến tính hàng năm có ý nghĩa thống kê (p < 0,001). Đáng chú ý, tất cả 6 mục riêng lẻ đều cho thấy xu hướng tăng tương tự — tăng mạnh nhất ở mục "lo lắng quá nhiều về mọi thứ" (worried too much about things).')

# Insert Figure 1
fig1_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Charts', 'Norway_Fig1_mental_distress_trends.png')
if os.path.exists(fig1_path):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(fig1_path, width=Cm(15))
    pc = doc.add_paragraph()
    pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = pc.add_run('Hình 1. Điểm trung bình căng thẳng tâm thần từ 2011 đến 2024 cho thanh thiếu niên Na Uy. Ước tính điều chỉnh cho khác biệt về lớp học và tính trung tâm địa lý. Khoảng tin cậy dựa trên sai số chuẩn điều chỉnh cho cụm thành phố–năm. (Figure 1. Mean mental distress scores from 2011 to 2024 for Norwegian adolescents.)')
    r.font.name = 'Times New Roman'; r.font.size = Pt(9); r.italic = True

# Bảng 2: Thay đổi trung bình hàng năm của căng thẳng tâm thần
add_heading(doc, 'Bảng 2. Thay đổi tuyến tính trung bình hàng năm của căng thẳng tâm thần (2011–2024)', 3)
add_table(doc,
    ['Biến số', 'Giai đoạn', 'Nam b', 'Nữ b', 'Ý nghĩa'],
    [
        ['Tất cả là gánh nặng', '2011–2024', '0,007', '0,010', 'p < 0,001'],
        ['Khó ngủ', '2011–2024', '0,017', '0,018', 'p < 0,001'],
        ['Không hạnh phúc, buồn', '2011–2024', '0,007', '0,012', 'p < 0,001'],
        ['Tuyệt vọng tương lai', '2011–2024', '0,023', '0,033', 'p < 0,001'],
        ['Cứng đờ/căng thẳng', '2011–2024', '0,005', '0,010', 'p < 0,001'],
        ['Lo lắng quá nhiều', '2011–2024', '0,021', '0,020', 'p < 0,001'],
        ['TỔNG (mean 6 mục)', '2011–2024', '0,013', '0,017', 'p < 0,001'],
    ],
    widths=[4.5, 2.5, 2.0, 2.0, 2.5]
)
add_p(doc, 'Ghi chú: b = hệ số hồi quy tuyến tính (thay đổi trung bình mỗi năm). Điều chỉnh cho lớp học và vùng. In đậm = p < 0,001.', size=9, italic=True)

add_p(doc, '2.2. Xu hướng thay đổi các yếu tố giải thích tiềm năng', bold=True)
add_p(doc, 'Đối với cả hai giới, có xu hướng tăng tuyến tính trong tối ở nhà, thời gian dùng mạng xã hội, sử dụng cần sa, và bất mãn với trường học. Tăng nhiều nhất ở mạng xã hội và bất mãn trường học. Tăng cần sa khá nhỏ. Bắt nạt ổn định theo thời gian, trong khi khó khăn tài chính GIẢM và hoạt động thể chất TĂNG. Chỉ ở nam, hài lòng với cha mẹ tăng nhẹ.')

# Insert Figure 2
fig2_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Charts', 'Norway_Fig2_explanatory_factors.png')
if os.path.exists(fig2_path):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(fig2_path, width=Cm(15))
    pc = doc.add_paragraph()
    pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = pc.add_run('Hình 2. Điểm z trung bình ước tính cho các yếu tố giải thích tiềm năng qua giai đoạn nghiên cứu. Ước tính điều chỉnh cho lớp học và tính trung tâm. KTC dựa trên sai số chuẩn điều chỉnh cho cụm thành phố–năm. (Figure 2. Estimated mean z-scores for potential explanatory factors across the study period.)')
    r.font.name = 'Times New Roman'; r.font.size = Pt(9); r.italic = True

# Bảng 3: Thay đổi tuyến tính hàng năm cho các yếu tố
add_heading(doc, 'Bảng 3. Thay đổi tuyến tính trung bình hàng năm cho các yếu tố giải thích', 3)
add_table(doc,
    ['Yếu tố', 'Giai đoạn', 'Nam b', 'Nam β', 'Nữ b', 'Nữ β'],
    [
        ['Khó khăn tài chính', '2011–2024', '−0,006***', '−0,006', '−0,011***', '−0,012'],
        ['Tối ở nhà', '2011–2024', '0,022***', '0,026', '0,020***', '0,023'],
        ['Hoạt động thể chất', '2011–2024', '0,045***', '0,030', '0,016***', '0,011'],
        ['Bắt nạt', '2011–2024', '0,003', '0,002', '−0,003', '−0,003'],
        ['Hài lòng với cha mẹ', '2011–2024', '0,015***', '0,013', '0,002', '0,002'],
        ['Mạng xã hội', '2014–2024', '0,084***', '0,056', '0,079***', '0,053'],
        ['Sử dụng cần sa', '2011–2024', '0,008***', '0,012', '0,006***', '0,009'],
        ['Bất mãn trường học', '2014–2024', '0,037***', '0,048', '0,041***', '0,053'],
    ],
    widths=[3.5, 2.0, 2.0, 1.5, 2.0, 1.5]
)
add_p(doc, 'Ghi chú: b = hệ số chưa chuẩn hóa. β = chuẩn hóa theo y. *** p < 0,001. In đậm = có ý nghĩa thống kê.', size=9, italic=True)

add_p(doc, '2.3. Thay đổi độc lực (virulence) của các yếu tố theo thời gian', bold=True)
add_p(doc, 'Tất cả 8 yếu tố đều có liên quan cắt ngang với căng thẳng tâm thần theo chiều dự kiến. Bất mãn trường học có liên quan mạnh nhất (β = 0,40 nam; β = 0,44 nữ), hoạt động thể chất yếu nhất.')
add_p(doc, 'Ở nam: độc lực TĂNG cho khó khăn tài chính, tối ở nhà, và bất mãn trường học.')
add_p(doc, 'Ở nữ: độc lực TĂNG cho khó khăn tài chính và tối ở nhà; nhưng GIẢM cho bắt nạt, cần sa, mạng xã hội, và bất mãn trường học.')

# Bảng 4: Thay đổi độc lực
add_heading(doc, 'Bảng 4. Thay đổi độc lực theo thời gian cho từng yếu tố giải thích', 3)
add_table(doc,
    ['Yếu tố', 'Nam — Tác động chính β', 'Nam — Thay đổi/năm', 'Nữ — Tác động chính β', 'Nữ — Thay đổi/năm'],
    [
        ['Khó khăn tài chính', '0,23', '0,006***', '0,24', '0,003***'],
        ['Tối ở nhà', '0,11', '0,004***', '0,11', '0,004***'],
        ['Hoạt động thể chất', '−0,06', '−0,000', '−0,08', '−0,001'],
        ['Bắt nạt', '0,30', '−0,001', '0,31', '−0,003***'],
        ['Hài lòng với cha mẹ', '−0,20', '−0,001', '−0,23', '−0,001'],
        ['Mạng xã hội', '0,13', '−0,001', '0,18', '−0,002***'],
        ['Sử dụng cần sa', '0,16', '0,002', '0,12', '−0,005***'],
        ['Bất mãn trường học', '0,40', '0,004***', '0,44', '−0,007***'],
    ],
    widths=[3.5, 3.0, 2.5, 3.0, 2.5]
)
add_p(doc, 'Ghi chú: β = hệ số chuẩn hóa theo y cho tác động chính. Thay đổi/năm = tương tác với thời gian. *** p < 0,001.', size=9, italic=True)

add_p(doc, '2.4. Thay đổi xu hướng khi giữ cố định tác động của các yếu tố (Decomposition)', bold=True)
add_p(doc, 'Chuỗi mô hình hồi quy tuyến tính được chạy để kiểm tra xem giữ cố định tác động của mỗi yếu tố có giảm xu hướng tăng ước tính hay không.')

add_p(doc, 'Ở NAM:', bold=True)
add_p(doc, '• Giữ cố định tối ở nhà: b giảm nhẹ (0,013 → 0,012), có ý nghĩa nhưng nhỏ.')
add_p(doc, '• Giữ cố định mạng xã hội: b giảm đáng kể (0,016 → 0,011; χ² = 911,50, p < 0,001).')
add_p(doc, '• Giữ cố định bất mãn trường học: b giảm RẤT MẠNH (0,016 → −0,004; χ² = 843,17, p < 0,001) — KHÔNG còn khác biệt đáng kể so với 0. Nghĩa là: bất mãn trường học giải thích TOÀN BỘ xu hướng tăng ở nam.')
add_p(doc, '• Giữ cố định các yếu tố khác (khó khăn tài chính, thể chất, bắt nạt, cha mẹ): KHÔNG giảm xu hướng.')

add_p(doc, 'Ở NỮ:', bold=True)
add_p(doc, '• Giữ cố định tối ở nhà: b giảm nhẹ (0,016 → 0,015).')
add_p(doc, '• Giữ cố định mạng xã hội: b giảm ĐÁNG KỂ (0,015 → 0,006; χ² = 764,04, p < 0,001).')
add_p(doc, '• Giữ cố định bất mãn trường học: b giảm RẤT MẠNH (0,015 → −0,004; χ² = 843,17, p < 0,001) — KHÔNG còn khác biệt đáng kể so với 0.')
add_p(doc, '• Giữ cố định các yếu tố khác: KHÔNG giảm xu hướng.')

# Bảng 5: Decomposition results
add_heading(doc, 'Bảng 5. Thay đổi trung bình hàng năm khi giữ cố định từng yếu tố', 3)
add_table(doc,
    ['Yếu tố giữ cố định', 'Giai đoạn', 'Nam b (gốc → sau)', 'Nữ b (gốc → sau)', 'Kết luận'],
    [
        ['Không giữ (baseline)', '2011–2024', '0,013', '0,017', 'Xu hướng tăng'],
        ['Khó khăn tài chính', '2011–2024', '0,013 → 0,013', '0,017 → 0,016', 'Không thay đổi'],
        ['Tối ở nhà', '2011–2024', '0,013 → 0,012', '0,017 → 0,015', 'Giảm nhẹ'],
        ['Hoạt động thể chất', '2011–2024', '0,013 → 0,013', '0,017 → 0,017', 'Không thay đổi'],
        ['Bắt nạt', '2011–2024', '0,013 → 0,013', '0,017 → 0,017', 'Không thay đổi'],
        ['Hài lòng cha mẹ', '2011–2024', '0,013 → 0,013', '0,017 → 0,017', 'Không thay đổi'],
        ['Mạng xã hội', '2014–2024', '0,016 → 0,011', '0,015 → 0,006', 'GIẢM ĐÁNG KỂ'],
        ['Sử dụng cần sa', '2011–2024', '0,013 → 0,013', '0,017 → 0,015', 'Nhỏ (chỉ nữ)'],
        ['Bất mãn trường học', '2014–2024', '0,016 → −0,004', '0,015 → −0,004', 'LOẠI BỎ XU HƯỚNG'],
    ],
    widths=[3.5, 2.0, 3.0, 3.0, 3.0]
)
add_p(doc, 'Ghi chú: b = hệ số tuyến tính hàng năm. Gốc = mô hình không giữ cố định. Sau = mô hình giữ cố định yếu tố đó. In đậm = thay đổi có ý nghĩa thống kê (p < 0,001).', size=9, italic=True)

# Bảng 6: Tóm tắt kết quả
add_heading(doc, 'Bảng 6. Tóm tắt kết quả nghiên cứu', 3)
add_table(doc,
    ['Yếu tố', 'Liên quan với distress?', 'Tỷ lệ tăng?', 'Độc lực tăng?', 'Giữ cố định giảm xu hướng?', 'Kết luận'],
    [
        ['Khó khăn tài chính', 'Có', 'Không (giảm)', 'Có, nhưng nhỏ', 'Không', 'Không giải thích'],
        ['Tối ở nhà', 'Có', 'Có', 'Có', 'Có, nhỏ', 'Đóng góp nhỏ'],
        ['Hoạt động thể chất', 'Có', 'Không (tăng)', 'Không', 'Không', 'Không giải thích'],
        ['Bắt nạt', 'Có', 'Không (ổn định)', 'Không', 'Không', 'Không giải thích'],
        ['Hài lòng cha mẹ', 'Có', 'Không', 'Không', 'Không', 'Không giải thích'],
        ['Mạng xã hội', 'Có', 'Có', 'Có (chỉ nam)', 'Có, đáng kể', 'GIẢI THÍCH MỘT PHẦN'],
        ['Sử dụng cần sa', 'Có', 'Có (nhỏ)', 'Không', 'Có, nhỏ (chỉ nữ)', 'Nhỏ, chỉ nữ'],
        ['Bất mãn trường học', 'Có', 'Có', 'Có (chỉ nam)', 'Có, RẤT LỚN', 'GIẢI THÍCH CHÍNH'],
    ],
    widths=[3.0, 2.0, 2.0, 2.0, 3.0, 2.5]
)

# ========== 10. THẢO LUẬN ==========
add_page_ref(doc, '6–8', 'Social Science & Medicine', 'Vol. 384, 2025')
add_heading(doc, '3. THẢO LUẬN', 2)

add_p(doc, 'Nghiên cứu này cung cấp cái nhìn sâu về các giải thích tiềm năng cho xu hướng tăng căng thẳng tâm thần VTN Na Uy từ 2011 đến 2024. Sử dụng khung dịch tễ học của Keyes và Platt (2024), cho thấy ba trong tám yếu tố được xem xét nổi lên là đóng góp hợp lý cho xu hướng tăng ở cả nữ và nam: bất mãn trường học tăng, thời gian dùng mạng xã hội tăng, và tối ở nhà tăng. Cần sa tăng cũng là đóng góp hợp lý nhưng nhỏ chỉ ở nữ. Bốn yếu tố còn lại không hợp lý: khó khăn tài chính, thiếu vận động, bắt nạt và bất mãn với cha mẹ.')

add_p(doc, 'Bất mãn trường học.', bold=True)
add_p(doc, 'Bất mãn trường học nổi lên là giải thích hợp lý nhất. Tỷ lệ bất mãn tăng theo thời gian, và giữ cố định yếu tố này loại bỏ hoàn toàn xu hướng tăng ở nam, phần lớn ở nữ. Phát hiện phù hợp với Cosma et al. (2020) ở 36 quốc gia — bất mãn trường học là yếu tố mạnh nhất liên quan đến căng thẳng tâm thần. Cơ chế tiềm năng bao gồm: áp lực học tập tăng, kỳ vọng thành tích cao, thay đổi thực hành giáo dục (nhiều kiểm tra hơn, ít thời gian tự do).')

add_p(doc, 'Mạng xã hội.', bold=True)
add_p(doc, 'Thời gian dùng mạng xã hội tăng cũng nổi lên là giải thích hợp lý. Phát hiện phù hợp với các NC trước về liên quan giữa mạng xã hội cao và tăng căng thẳng (Appel et al., 2020; Colasante et al., 2024; Fitzpatrick et al., 2023). Tuy nhiên, bằng chứng vẫn bị tranh cãi — một số nghiên cứu không tìm thấy mối quan hệ tiến triển (Beeres et al., 2021; Coyne et al., 2020; Ferguson et al., 2025). Cơ chế đề xuất: so sánh xã hội tăng, bắt nạt trực tuyến, giảm tương tác mặt đối mặt. Cần lưu ý mối quan hệ có thể hai chiều — VTN căng thẳng có thể tăng dùng MXH để phân tâm.')

add_p(doc, 'Tối ở nhà và cần sa.', bold=True)
add_p(doc, 'Tăng tối ở nhà và tăng cần sa cũng đóng góp nhưng nhỏ hơn. Tối ở nhà tăng có thể phản ánh thay đổi lối sống — giảm hoạt động xã hội trực tiếp, tăng thời gian trực tuyến. Cần sa chỉ đóng góp nhỏ và chỉ ở nữ.')

add_p(doc, 'Các yếu tố không giải thích.', bold=True)
add_p(doc, 'Bốn yếu tố không giải thích: khó khăn tài chính (GIẢM theo thời gian), hoạt động thể chất (TĂNG), bắt nạt (ổn định), bất mãn với cha mẹ (ổn định hoặc tăng nhẹ). Kết quả khá rõ ràng — những yếu tố này không đóng góp vào xu hướng tăng.')

# ========== 11. HẠN CHẾ ==========
add_p(doc, 'Hạn chế', bold=True)
add_p(doc, '• Giả định lớn nhất: chiều hướng nhân quả một chiều (yếu tố → distress). Mục đích không phải ước tính tác động nhân quả, mà thu hẹp phạm vi giải thích hợp lý.')
add_p(doc, '• Tất cả phân tích với dữ liệu Na Uy — cần phân tích ở nước khác để đánh giá tính khái quát.')
add_p(doc, '• Tất cả biến đo bằng tự báo cáo với mục đơn giản — đo lường không hoàn hảo. Đặc biệt mạng xã hội chỉ đo thời gian (top: ≥3 giờ/ngày) — không đo nội dung.')
add_p(doc, '• Mô hình tuyến tính — phương pháp linh hoạt hơn có thể cải thiện.')
add_p(doc, '• Chỉ kiểm tra yếu tố có trong bộ dữ liệu — có thể thiếu yếu tố khác.')

# ========== 12. KẾT LUẬN ==========
add_heading(doc, '4. KẾT LUẬN', 2)
add_p(doc, 'Nghiên cứu này cung cấp cái nhìn sâu về các yếu tố có thể đã đóng góp vào xu hướng tăng căng thẳng tâm thần VTN tại Na Uy. Bất mãn trường học tăng, sử dụng mạng xã hội tăng, tối ở nhà nhiều hơn, và sử dụng cần sa được xác định là yếu tố giải thích tiềm năng, mặc dù tác động ước tính khác nhau. Kết quả nhấn mạnh tính phức tạp của xu hướng sức khỏe tâm thần và tầm quan trọng của việc xem xét nhiều giải thích tiềm năng. Chiều hướng nhân quả vẫn chưa chắc chắn — nghiên cứu đặt câu hỏi cho NC tương lai, đặc biệt về tương tác giữa mạng xã hội và sức khỏe VTN.')

add_p(doc, 'Chúng tôi có thể kết luận rõ hơn rằng thay đổi trong tình hình tài chính gia đình, hoạt động thể chất, bắt nạt, và hài lòng với cha mẹ là những giải thích KHÔNG HỢP LÝ cho xu hướng tăng căng thẳng tâm thần. Phát hiện của chúng tôi có thể làm nền tảng cho NC tương lai nhằm cung cấp bằng chứng mạnh hơn.')

# ========== 13. TÀI LIỆU THAM KHẢO ==========
add_heading(doc, 'TÀI LIỆU THAM KHẢO', 2)
refs = [
    'Appel, M., Marker, C., Gnambs, T. (2020). Are social media ruining our lives? A review of meta-analytic evidence. Review of General Psychology, 24(1), 60–74.',
    'Armitage, J.M., et al. (2024). Identifying the causes and consequences of adolescent mental health difficulties. Nature Reviews Psychology.',
    'Brunborg, G.S., Nilsen, S.A., Skogen, J.C., Bang, L. (2025). Possible explanations for the upward trend in mental distress among adolescents in Norway from 2011 to 2024. Social Science & Medicine, 384, 118528.',
    'Colasante, T., et al. (2024). Any time spent on social media is associated with negative affect in youth. Nature Mental Health.',
    'Cosma, A., et al. (2020). Cross-national time trends in adolescent mental well-being from 2002 to 2018 and the explanatory role of schoolwork pressure. JAMA, 66(S), S50–S58.',
    'Eckersley, R. (2011). A new narrative of young people\'s health and well-being. Journal of Youth Studies, 14(5), 627–638.',
    'Foulkes, L., Andrews, J.L. (2023). Are mental health awareness efforts contributing to the rise in reported mental health problems? New Ideas in Psychology, 69, 101010.',
    'Haidt, J. (2024). The Anxious Generation. Penguin Press.',
    'Keyes, K.M., Platt, J.M. (2024). An epidemiological framework for explaining the upward trend in adolescent mental distress.',
    'Twenge, J.M. (2017). iGen: Why Today\'s Super-connected Kids Are Growing Up Less Rebellious. Simon and Schuster.',
]
for ref in refs:
    add_p(doc, ref, size=10)
add_p(doc, '(Xem danh mục đầy đủ trong bài gốc — 60+ tài liệu tham khảo)', size=10, italic=True)

# ========== 14. BẢNG VIẾT TẮT ==========
add_abbreviation_table(doc, [
    ('HSCL', 'Hopkins Symptom Checklist — Thang Triệu chứng Hopkins'),
    ('VTN', 'Vị thành niên (thanh thiếu niên)'),
    ('MXH', 'Mạng xã hội'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('NIPH', 'Norwegian Institute of Public Health — Viện Y tế Công cộng Na Uy'),
    ('NOVA', 'Norwegian Social Research — Trung tâm Nghiên cứu Xã hội Na Uy'),
    ('GAD-7', 'Generalized Anxiety Disorder 7-item Scale — Thang Rối loạn Lo âu Tổng quát 7 mục'),
    ('DASS-21', 'Depression Anxiety Stress Scale 21 — Thang Trầm cảm Lo âu Căng thẳng 21 mục'),
    ('IF', 'Impact Factor — Hệ số tác động'),
    ('β', 'Hệ số hồi quy chuẩn hóa'),
    ('b', 'Hệ số hồi quy chưa chuẩn hóa'),
    ('χ²', 'Chi bình phương'),
])

# ========== 15. PHẢN BIỆN ==========
add_red_heading(doc, 'QUAN ĐIỂM PHẢN BIỆN / CRITICAL REVIEW')

add_red(doc, 'Điểm mạnh:', bold=True)
strengths = [
    'Tạp chí Q1 IF = 5,4 — Social Science & Medicine, uy tín cao trong y tế công cộng.',
    'Mẫu CỰC LỚN: 979.043 VTN — một trong những nghiên cứu lớn nhất về xu hướng SKTT VTN.',
    'Xu hướng 13 năm (2011–2024) — dài nhất trong các NC cắt ngang lặp lại của Đề tài (cùng Korea 16 năm).',
    'Phương pháp decomposition nâng cao — không chỉ mô tả xu hướng mà xác định yếu tố giải thích.',
    'Sử dụng khung dịch tễ học Keyes & Platt (2024) — phương pháp luận chặt chẽ, có hệ thống.',
]
for s in strengths:
    add_red(doc, f'• {s}')

add_red(doc, 'Hạn chế chi tiết:', bold=True)
limits = [
    'Chỉ dữ liệu Na Uy — quốc gia Bắc Âu giàu, hệ thống phúc lợi mạnh. Khác biệt văn hóa rất lớn với Việt Nam và châu Á.',
    'Đo "mental distress" chung (6 mục HSCL) — KHÔNG tách lo âu riêng. Không dùng GAD-7 hay DASS-21 như các bài VN.',
    'Giả định chiều nhân quả một chiều (yếu tố → distress) — HẠN CHẾ LỚN NHẤT. Có thể distress → tăng dùng MXH, không chỉ ngược lại.',
    'Thiết kế cắt ngang lặp lại — KHÔNG theo dõi cùng cá nhân theo thời gian. Không phải cohort.',
    'Mạng xã hội chỉ đo thời gian tổng (top: ≥3 giờ/ngày) — không phân biệt loại nội dung, loại tương tác.',
    'Dữ liệu MXH chỉ từ 2014 — thiếu 3 năm đầu (2011–2013).',
    'Tự báo cáo — có thể bị ảnh hưởng bởi nhận thức tăng về SKTT (prevalence inflation hypothesis mà chính tác giả đề cập).',
]
for s in limits:
    add_red(doc, f'• {s}')

add_red(doc, 'Khoảng trống nghiên cứu / Research Gap:', bold=True)
gaps = [
    'CHƯA CÓ phân tích decomposition tương tự cho dữ liệu Việt Nam hoặc châu Á. Hoàng Trung Học 2025 có 2 thời điểm nhưng chưa dùng phương pháp này.',
    'Cần so sánh xu hướng Na Uy vs Việt Nam — vai trò bất mãn trường học và MXH ở VN có giống không?',
    'Thiếu NC về tác động NỘI DUNG mạng xã hội (không chỉ thời gian) lên SKTT VTN.',
    'Cần RCT hoặc nghiên cứu can thiệp tại trường để kiểm tra: giảm bất mãn trường học có giảm căng thẳng tâm thần không?',
    'Liên hệ VN: áp lực học tập tại VN (thi cử, học thêm) có thể còn mạnh hơn Na Uy — nhưng chưa có dữ liệu decomposition.',
]
for s in gaps:
    add_red(doc, f'• {s}')

# ========== SAVE ==========
outpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), '21_Norway_2025_SocSciMed.docx')
doc.save(outpath)
print(f'Saved: {outpath}')
