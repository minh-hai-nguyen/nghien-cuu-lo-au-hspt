# -*- coding: utf-8 -*-
"""Dịch toàn văn QT24 - Tarasenko et al. 2025 - Lancet Regional Health Europe - WHO Europe"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()

# ========== 1. LINK ==========
add_p(doc, 'Link bài báo gốc: https://doi.org/10.1016/j.lanepe.2025.101262', size=10)

# ========== 2. TIÊU ĐỀ ==========
add_heading(doc, 'Sức khỏe tâm thần của trẻ em và thanh niên trong khu vực WHO châu Âu', 1)
h2 = doc.add_paragraph()
r = h2.add_run('Mental Health of Children and Young People in the WHO Europe Region')
r.font.name = 'Times New Roman'; r.font.size = Pt(13); r.italic = True

# ========== 3. THÔNG TIN THƯ MỤC ==========
add_heading(doc, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(doc, [
    ('Tiêu đề gốc', 'Mental Health of Children and Young People in the WHO Europe Region'),
    ('Tiêu đề dịch', 'Sức khỏe tâm thần của trẻ em và thanh niên trong khu vực WHO châu Âu'),
    ('Tác giả', 'Anna Tarasenko, Gisel Josy, Hellen Minnis, Jennifer Hall, Andrea Danese, Jennifer Y.F. Lau, Samuele Cortese, Argyris Stringaris, Cassie Redlich, Dennis Ougrin'),
    ('Cơ quan', 'UNODC Kyiv; Đại học Glasgow; WHO Athens & Copenhagen; King\'s College London; Queen Mary & UCL London; Đại học Southampton'),
    ('Tạp chí', 'The Lancet Regional Health — Europe (Q1, IF ≈ 15,0)'),
    ('Thông tin xuất bản', '2025, 13 trang, Series: Transforming Mental Health in Europe'),
    ('DOI', '10.1016/j.lanepe.2025.101262'),
    ('Loại nghiên cứu', 'Tổng quan chính sách (policy review / Series paper) — KHÔNG phải nghiên cứu gốc'),
    ('Phạm vi', 'Khu vực WHO châu Âu (53 quốc gia thành viên, bao gồm cả Trung Á)'),
])

add_page_ref(doc, '1–13', 'Lancet Regional Health Europe', '2025')

# ========== TÓM TẮT ==========
add_heading(doc, 'TÓM TẮT', 2)
p = doc.add_paragraph()
r = p.add_run('Phần lớn rối loạn sức khỏe tâm thần bắt đầu trước tuổi trưởng thành. Chúng có tỷ lệ cao, gây tàn phế và thường có thể điều trị được. Bài Series này thảo luận các vấn đề hiện tại đóng góp vào cuộc khủng hoảng SKTT trẻ em, VTN và thanh niên ngày càng tăng tại châu Âu. Bao gồm: tác động đại dịch COVID-19, xung đột quân sự leo thang, khủng hoảng khí hậu, và môi trường số không được kiểm soát. Các vấn đề SKTT ở giới trẻ càng phức tạp hơn bởi sự thay đổi trong cách trẻ em tương tác với thế giới, chăm sóc sức khỏe và các dịch vụ khác.')
r.font.name = 'Times New Roman'; r.font.size = Pt(12)
r.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

add_p(doc, 'Bài viết đưa ra danh sách toàn diện các khuyến nghị để giải quyết vấn đề SKTT thông qua: tích hợp phòng ngừa sáng tạo và phương pháp điều trị với sự hỗ trợ của dịch vụ cộng đồng và hệ thống hỗ trợ, cùng với chiến lược nghiên cứu và triển khai mạnh mẽ để đảm bảo chăm sóc dựa trên bằng chứng, hiệu quả chi phí.')

add_p(doc, 'Từ khóa: Sức khỏe tâm thần toàn cầu; Trẻ em và thanh niên; Tổng quan.')

# ========== TÓM TẮT ĐÁNH GIÁ NHANH ==========
add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
for b in [
    'Lancet Regional Health Europe Q1 IF ≈ 15 — tạp chí uy tín rất cao.',
    '9 triệu VTN châu Âu (10–19 tuổi) sống với rối loạn SKTT.',
    'Lo âu + trầm cảm chiếm >50% tổng ca rối loạn ở VTN châu Âu.',
    'Bài Series chính sách — tổng hợp toàn diện thách thức + khuyến nghị cho 53 quốc gia.',
    'Đặc biệt nhấn mạnh: tích hợp SKTT vào giáo dục, dịch vụ cộng đồng, và kiểm soát MXH.',
]:
    add_p(doc, f'• {b}')

add_p(doc, 'Hạn chế:', bold=True)
for b in [
    'Tổng quan chính sách — KHÔNG phải nghiên cứu gốc, thiếu dữ liệu tỷ lệ cụ thể từng nước.',
    'Bối cảnh châu Âu — hệ thống dịch vụ SKTT phát triển hơn nhiều so với ASEAN/VN.',
    'Khuyến nghị rộng — chưa đánh giá cụ thể hiệu quả/chi phí cho từng can thiệp.',
]:
    add_p(doc, f'• {b}')

add_p(doc, 'Hướng cải thiện:', bold=True)
for b in [
    'Tổng quan tương tự cho khu vực ASEAN — đánh giá thực trạng dịch vụ SKTT VTN.',
    'So sánh mô hình châu Âu vs VN — khả năng áp dụng.',
    'Đánh giá chi phí–hiệu quả can thiệp SKTT trường học cho bối cảnh VN.',
]:
    add_p(doc, f'• {b}')

# ========== GIỚI THIỆU ==========
add_page_ref(doc, '1–2', 'Lancet Regional Health Europe', '2025')
add_heading(doc, '1. GIỚI THIỆU', 2)

add_p(doc, 'Trên toàn cầu, đã có lời kêu gọi khẩn cấp giải quyết cuộc khủng hoảng SKTT trẻ em, VTN và thanh niên ngày càng tăng. Tại cuộc họp đầu tiên của Liên minh SKTT Toàn châu Âu của WHO, cải thiện chất lượng sức khỏe tâm thần của trẻ em và thanh niên được xác định là ưu tiên hàng đầu.')

add_p(doc, 'Khoảng 9 triệu VTN (10–19 tuổi) trong khu vực WHO châu Âu sống với một tình trạng SKTT. Lo âu và trầm cảm chiếm hơn 50% tổng số ca rối loạn — biến chúng thành vấn đề SKTT hàng đầu. Tự tử là nguyên nhân tử vong hàng đầu thứ hai ở nhóm 15–19 tuổi tại nhiều nước.')

add_p(doc, 'Các yếu tố sinh học, tâm lý, xã hội và môi trường đóng góp vào SKTT kém ở VTN. Nghèo đói, phân biệt đối xử, bạo lực, và khuyết tật là yếu tố nguy cơ. Những thay đổi gần đây làm tình hình nghiêm trọng hơn: đại dịch COVID-19, xung đột quân sự, khủng hoảng khí hậu, và môi trường số không kiểm soát.')

# ========== THÁCH THỨC HIỆN TẠI ==========
add_page_ref(doc, '2–4', 'Lancet Regional Health Europe', '2025')
add_heading(doc, '2. THÁCH THỨC HIỆN TẠI', 2)

add_p(doc, '2.1. Tác động đại dịch COVID-19', bold=True)
add_p(doc, 'Đại dịch gây gián đoạn nghiêm trọng cuộc sống VTN — ảnh hưởng xã hội hóa, giáo dục, SKTT. Phong tỏa, đóng cửa trường, giãn cách xã hội dẫn đến tăng cảm giác cô lập, cô đơn, lo âu và bất an. Số ngày đóng cửa trường khác nhau: 0 ngày (Phần Lan, Thụy Điển) đến 341 ngày (Ý). Nghiên cứu cho thấy tăng chuyển tuyến rối loạn tâm thần sau phong tỏa, bao gồm tăng đáng kể tự gây hại tại khoa cấp cứu.')

add_p(doc, '2.2. Mạng xã hội và công nghệ số', bold=True)
add_p(doc, 'Trong khi công nghệ số và MXH cung cấp kênh kết nối, lạm dụng chúng đã được xác lập là yếu tố nguy cơ có thể cho kết quả SKTT kém ở VTN. Tuy nhiên, bằng chứng vẫn lẫn lộn. Lo ngại bao gồm: tiếp xúc quá mức với hình ảnh chiến tranh và nội dung tình dục không phù hợp. Khảo sát HBSC cho thấy 11% VTN 11–15 tuổi sử dụng MXH "có vấn đề" — nhưng định nghĩa "có vấn đề" còn tranh cãi.')

add_p(doc, '2.3. Xung đột quân sự và di cư', bold=True)
add_p(doc, 'Xung đột vũ trang tại châu Âu (Ukraine, cùng nhiều vùng khác) gây ảnh hưởng sâu sắc lên SKTT trẻ em — bao gồm PTSD, trầm cảm, lo âu, tang tóc. Trẻ em tị nạn và di cư đặc biệt dễ bị tổn thương do mất ổn định, tách biệt gia đình.')

add_p(doc, '2.4. Tiếp cận dịch vụ SKTT', bold=True)
add_p(doc, 'Nhu cầu SKTT VTN tăng nhanh hơn khả năng đáp ứng của hệ thống dịch vụ — dẫn đến thời gian chờ dài hơn, kiệt sức nhân viên, khó tuyển dụng lâm sàng. Ngân sách cho SKTT trẻ em/VTN vẫn eo hẹp trên toàn khu vực. Trong khảo sát 27 nước châu Âu 2017: 16 nước xếp dịch vụ SKTT cộng đồng cho trẻ em/VTN là "thiếu hoặc không đủ". Chỉ 6/28 nước báo cáo có giám sát kết quả điều trị định kỳ.')

add_p(doc, '2.5. Nghịch cảnh thời thơ ấu', bold=True)
add_p(doc, 'Lạm dụng thời thơ ấu (thể chất, tình dục, cảm xúc; bỏ mặc) là yếu tố nguy cơ xuyên chẩn đoán (trans-diagnostic) cho rối loạn tâm thần, với bằng chứng nhân quả rõ ràng từ các NC bán thực nghiệm. Lạm dụng không chỉ tăng nguy cơ khởi phát mà còn liên quan đến diễn biến bệnh xấu hơn và đáp ứng điều trị kém hơn.')

# ========== GIẢI PHÁP ĐỀ XUẤT ==========
add_page_ref(doc, '4–8', 'Lancet Regional Health Europe', '2025')
add_heading(doc, '3. GIẢI PHÁP ĐỀ XUẤT', 2)

add_p(doc, '3.1. Chiến lược phòng ngừa', bold=True)
add_p(doc, 'Phòng ngừa trong SKTT VTN có tiềm năng giảm đáng kể gánh nặng bằng cách giải quyết yếu tố nguy cơ, tăng cường yếu tố bảo vệ, và can thiệp sớm. Ba cấp độ:')
add_p(doc, '• Phổ quát (universal): Nhắm toàn bộ dân số — ví dụ: chương trình toàn trường phòng tự tử.')
add_p(doc, '• Chọn lọc (selective): Nhắm nhóm nguy cơ cao — ví dụ: trẻ gia đình thu nhập thấp, trẻ có nghịch cảnh.')
add_p(doc, '• Chỉ định (indicated): Nhắm cá nhân có dấu hiệu sớm nhưng chưa đạt ngưỡng chẩn đoán.')

add_p(doc, '3.2. Tích hợp SKTT vào chương trình giáo dục', bold=True)
add_p(doc, 'Trường học đóng vai trò then chốt trong can thiệp sớm và phòng ngừa. Triển khai chương trình học tập cảm xúc–xã hội phổ quát tại trường có thể giảm kỳ thị, tăng khả năng phục hồi, cung cấp chiến lược ứng phó. Các chương trình dựa trên bằng chứng: MindMatters, SEAL, SPARK Resilience. Tuy nhiên, mức độ ưu tiên SKTT rất khác nhau: 72,8% trường ở Ba Lan vs chỉ 17,8% ở Pháp. Chỉ 15,7% trường ở Pháp có chính sách SKTT riêng so với 78,3% ở Hà Lan.')

add_p(doc, '3.3. Phát triển dịch vụ cộng đồng', bold=True)
add_p(doc, 'Cần phát triển dịch vụ SKTT cộng đồng chất lượng cao, bao gồm chăm sóc cộng đồng tập trung thay vì nhập viện. Mô hình "kê đơn xã hội" (social prescribing): giới thiệu trẻ em/thanh niên đến tổ chức địa phương tổ chức hoạt động có lợi cho sức khỏe (nghệ thuật, thể thao, thiên nhiên). Mô hình này giúp trẻ có quyền chủ động, ít kỳ thị hơn điều trị truyền thống.')

add_p(doc, '3.4. Phát triển điều trị dược lý', bold=True)
add_p(doc, 'Dữ liệu dịch tễ dược lý cho thấy tăng sử dụng thuốc tâm thần cho trẻ em/VTN trong 20 năm qua. Bằng chứng ngày càng tăng từ RCT và NC quan sát:')

# Bảng 1: Hiệu quả thuốc
add_heading(doc, 'Bảng 1. Hiệu quả điều trị dược lý cho các rối loạn tâm thần trẻ em/VTN', 3)
add_table(doc,
    ['Rối loạn', 'Thuốc hiệu quả', 'Mức hiệu quả', 'Ghi chú'],
    [
        ['ADHD', 'Thuốc kích thích (stimulants)', 'Cao nhất', 'Triệu chứng cốt lõi'],
        ['ASD — kích ứng', 'Aripiprazole, Risperidone', 'Cao', ''],
        ['Trầm cảm', 'Fluoxetine', 'Trung bình', 'Hiệu ứng placebo cao ở trẻ em'],
        ['OCD', 'SSRI (Fluoxetine)', 'Khá cao', 'Ngoại lệ với placebo'],
        ['Lo âu', 'SSRI', 'Cao ban đầu', 'Thiếu dữ liệu dài hạn'],
        ['Rối loạn Tic/Tourette', 'Aripiprazole, Haloperidol', 'Trung bình–Cao', ''],
        ['Tâm thần phân liệt', 'Risperidone, Olanzapine', 'Trung bình–Cao', 'Dữ liệu hạn chế ở trẻ em'],
    ],
    widths=[3.0, 4.0, 2.5, 4.0]
)
add_p(doc, 'Ghi chú: Tóm tắt từ Bảng 1 gốc. Hiệu quả dựa trên RCT ngắn hạn. Thiếu dữ liệu an toàn dài hạn cho hầu hết thuốc.', size=9, italic=True)

add_p(doc, '3.5. Chiến lược nghiên cứu và triển khai mạnh mẽ', bold=True)
add_p(doc, 'Một số can thiệp phòng ngừa có cơ sở lý thuyết mạnh nhưng tỏ ra không hiệu quả hoặc gây hại. Ví dụ: RCT lớn về chương trình chánh niệm (mindfulness) toàn trường không cho thấy hiệu quả tổng thể và gợi ý có thể gây hại cho một số kết quả. Cần khung đánh giá hại/lợi cho bất kỳ can thiệp nào.')

add_p(doc, '3.6. Nghiên cứu có sự tham gia của người dùng dịch vụ', bold=True)
add_p(doc, 'Sự tham gia có ý nghĩa của trẻ em và thanh niên trong nghiên cứu và thiết kế dịch vụ là thiết yếu. Tuy nhiên, thực hành này vẫn chưa phổ biến và thường chỉ mang tính hình thức.')

# ========== KHUYẾN NGHỊ ==========
add_page_ref(doc, '8–9', 'Lancet Regional Health Europe', '2025')
add_heading(doc, '4. KHUYẾN NGHỊ CHÍNH SÁCH', 2)

# Bảng 2: Tổng hợp khuyến nghị
add_heading(doc, 'Bảng 2. Tổng hợp khuyến nghị chính cho khu vực WHO châu Âu', 3)
add_table(doc,
    ['Lĩnh vực', 'Khuyến nghị', 'Ý nghĩa cho VN'],
    [
        ['Phòng ngừa', 'Triển khai chương trình phòng ngừa phổ quát, chọn lọc và chỉ định', 'VN cần phát triển chương trình tương tự tại trường'],
        ['Trường học', 'Tích hợp SKTT vào chương trình giáo dục, đào tạo giáo viên nhận biết dấu hiệu', 'Phù hợp với Wen 2020 (hỗ trợ SKTT trường OR = 0,562)'],
        ['Dịch vụ cộng đồng', 'Phát triển dịch vụ cộng đồng thay vì nhập viện', 'VN mới có 8,4% tiếp cận — cần mở rộng'],
        ['MXH/số', 'Kiểm soát nội dung, quy định nền tảng, giáo dục kỹ năng số', 'Phù hợp với phát hiện Norway 2025, Chen 2023'],
        ['Nghiên cứu', 'Đánh giá cả hại lẫn lợi, nghiên cứu triển khai', 'VN cần NC can thiệp + đánh giá hiệu quả'],
        ['Tham gia', 'Trẻ em/thanh niên tham gia thiết kế dịch vụ', 'Chưa phổ biến ở VN'],
        ['Dược lý', 'RCT dài hạn, giảm hiệu ứng placebo, tiếp cận chính xác', 'VN thiếu NC dược lý SKTT VTN'],
        ['Bất bình đẳng', 'Chú ý nhóm dễ bị tổn thương: nghèo, tị nạn, DTTS', 'Phù hợp: VN có DTTS, vùng khó khăn'],
    ],
    widths=[2.5, 6.0, 5.0]
)

# ========== THẢO LUẬN ==========
add_page_ref(doc, '8–9', 'Lancet Regional Health Europe', '2025')
add_heading(doc, '5. THẢO LUẬN', 2)

add_p(doc, 'Phòng ngừa SKTT kém vẫn là ưu tiên; tuy nhiên, có sự khác biệt khu vực trong hệ thống triển khai. Phân tích 38 chính sách/kế hoạch SKTT quốc gia cho thấy sự khác biệt đáng kể: mô hình từ trên xuống phổ biến ở Đông Âu, mô hình từ dưới lên ở Bắc Âu. Chỉ 5 nước có tổ chức đa ngành cho triển khai SKTT.')

add_p(doc, 'Bài viết thừa nhận chủ yếu dựa trên tổng quan hệ thống và phân tích tổng hợp đã công bố — không có dữ liệu gốc mới. Tuy nhiên, không phải lúc nào cũng có bằng chứng hệ thống cho mọi khuyến nghị.')

add_p(doc, 'Cuộc khủng hoảng SKTT VTN tại châu Âu là đa yếu tố: dịch COVID-19, xung đột, biến đổi khí hậu, và yếu tố công nghệ. Giải quyết đòi hỏi cách tiếp cận toàn diện: phòng ngừa tại trường, tiếp cận công bằng dịch vụ SKTT, chấm dứt xung đột, và giải quyết bất bình đẳng kinh tế–xã hội gốc rễ.')

# ========== KẾT LUẬN ==========
add_heading(doc, '6. KẾT LUẬN', 2)
add_p(doc, 'Dữ liệu tổng hợp WHO châu Âu cho thấy 9 triệu VTN đang sống với rối loạn SKTT và lo âu chiếm >50% tổng ca. Giải quyết cuộc khủng hoảng đòi hỏi chiến lược toàn diện: phòng ngừa dựa trên bằng chứng, dịch vụ cộng đồng, kiểm soát MXH, và hợp tác xuyên quốc gia. Mặc dù bối cảnh châu Âu (dịch vụ phát triển) khác VN/ASEAN, các nguyên tắc cốt lõi — tích hợp SKTT vào giáo dục, phát triển dịch vụ cộng đồng, can thiệp sớm — hoàn toàn áp dụng được.')

# ========== TÀI LIỆU THAM KHẢO ==========
add_heading(doc, 'TÀI LIỆU THAM KHẢO', 2)
refs = [
    'Tarasenko, A., et al. (2025). Mental health of children and young people in the WHO Europe region. Lancet Regional Health — Europe.',
    'UNICEF. (2021). The State of the World\'s Children: On My Mind — promoting, protecting and caring for children\'s mental health.',
    'Cortese, S., et al. (2018). Comparative efficacy and tolerability of medications for ADHD in children, adolescents, and adults. Lancet Psychiatry, 5(9), 727–738.',
    'Danese, A. & Widom, C.S. (2020). Objective and subjective experiences of child maltreatment and their relationships with psychopathology. Nature Human Behaviour, 4, 811–818.',
    'Stringaris, A., et al. (2018). Editorial: In search of specificity: anxiety and depression in childhood and adolescence. JCPP, 59, 1263–1266.',
]
for ref in refs:
    add_p(doc, ref, size=10)
add_p(doc, '(Xem danh mục đầy đủ trong bài gốc — 110+ tài liệu tham khảo)', size=10, italic=True)

# ========== BẢNG VIẾT TẮT ==========
add_abbreviation_table(doc, [
    ('WHO', 'World Health Organization — Tổ chức Y tế Thế giới'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('VTN', 'Vị thành niên (thanh thiếu niên)'),
    ('MXH', 'Mạng xã hội'),
    ('SSRI', 'Selective Serotonin Reuptake Inhibitors — Thuốc ức chế tái hấp thu serotonin chọn lọc'),
    ('ADHD', 'Attention-Deficit/Hyperactivity Disorder — Rối loạn Tăng động Giảm chú ý'),
    ('ASD', 'Autism Spectrum Disorder — Rối loạn Phổ tự kỷ'),
    ('OCD', 'Obsessive-Compulsive Disorder — Rối loạn Ám ảnh Cưỡng chế'),
    ('PTSD', 'Post-Traumatic Stress Disorder — Rối loạn Stress Sau Sang chấn'),
    ('RCT', 'Randomized Controlled Trial — Thử nghiệm Ngẫu nhiên Có đối chứng'),
    ('ACEs', 'Adverse Childhood Experiences — Trải nghiệm Bất lợi Thời thơ ấu'),
    ('SEAL', 'Social and Emotional Aspects of Learning — Các khía cạnh Cảm xúc-Xã hội của Học tập'),
    ('HBSC', 'Health Behaviour in School-aged Children — Hành vi Sức khỏe ở Trẻ em Đi học'),
    ('DMDD', 'Disruptive Mood Dysregulation Disorder — Rối loạn Rối loạn Điều hòa Tâm trạng'),
])

# ========== PHẢN BIỆN ==========
add_red_heading(doc, 'QUAN ĐIỂM PHẢN BIỆN / CRITICAL REVIEW')

add_red(doc, 'Điểm mạnh:', bold=True)
for s in [
    'Lancet Regional Health Europe Q1 IF ≈ 15 — tạp chí uy tín rất cao.',
    'Phạm vi toàn diện: 53 quốc gia khu vực WHO châu Âu.',
    'Bài Series chính sách — tổng hợp thách thức + giải pháp + khuyến nghị trong một khuôn khổ nhất quán.',
    'Đa tác giả từ nhiều tổ chức hàng đầu (WHO, UCL, King\'s College, Glasgow).',
    'Bao gồm góc nhìn người dùng dịch vụ (Gisel Josy — đồng tác giả là service user).',
]:
    add_red(doc, f'• {s}')

add_red(doc, 'Hạn chế chi tiết:', bold=True)
for s in [
    'KHÔNG phải nghiên cứu gốc — tổng quan chính sách/narrative review. Thiếu phương pháp hệ thống (không PRISMA, không tìm kiếm có cấu trúc).',
    'Thiếu DỮ LIỆU TỶ LỆ CỤ THỂ từng nước — chỉ có con số tổng (9 triệu, >50%). Khó so sánh chi tiết giữa các nước.',
    'Bối cảnh châu Âu — hệ thống SKTT, giáo dục, phúc lợi phát triển rất khác ASEAN/VN. Nhiều khuyến nghị (social prescribing, community services) đòi hỏi hạ tầng mà VN chưa có.',
    'Khuyến nghị RẤT RỘNG — không đánh giá cụ thể hiệu quả, chi phí, tính khả thi cho từng bối cảnh.',
    'Thừa nhận bằng chứng vẫn lẫn lộn cho MXH — không đưa ra kết luận rõ ràng.',
    'Không thảo luận bất bình đẳng thu nhập chi tiết (Korea 2024 đã làm tốt hơn).',
]:
    add_red(doc, f'• {s}')

add_red(doc, 'Khoảng trống nghiên cứu / Research Gap:', bold=True)
for s in [
    'CHƯA CÓ bài tổng quan tương tự cho khu vực ASEAN hoặc Việt Nam — đây là GAP lớn.',
    'Cần đánh giá hệ thống dịch vụ SKTT VTN tại VN theo tiêu chuẩn WHO — chỉ 8,4% tiếp cận (V-NAMHS 2022).',
    'Can thiệp tại trường (school-based): bài này khuyến nghị mạnh nhưng VN chưa có NC đánh giá hiệu quả can thiệp SKTT trường.',
    'So sánh mô hình Đông Âu (top-down) vs Bắc Âu (bottom-up) — VN thuộc mô hình nào? Cần nghiên cứu.',
    'Kiểm soát MXH: EU đã có DSA (Digital Services Act) — VN cần đánh giá khung pháp lý bảo vệ VTN trực tuyến.',
]:
    add_red(doc, f'• {s}')

# ========== SAVE ==========
outpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), '24_WHO_Europe_2025_LancetRegional.docx')
doc.save(outpath)
print(f'Saved: {outpath}')
