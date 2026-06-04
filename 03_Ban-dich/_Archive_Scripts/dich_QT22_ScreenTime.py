# -*- coding: utf-8 -*-
"""Dịch toàn văn QT22 - Li et al. 2025 - BJCP - Screen Time & Adolescent Mental Health"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()

# ========== 1. LINK ==========
add_p(doc, 'Link bài báo gốc: https://doi.org/10.1111/bjc.12547', size=10)

# ========== 2. TIÊU ĐỀ ==========
add_heading(doc, 'Mối liên quan cắt ngang và dọc giữa thời gian sử dụng màn hình với trầm cảm và lo âu ở thanh thiếu niên', 1)
h2 = doc.add_paragraph()
r = h2.add_run('Cross-sectional and Longitudinal Associations of Screen Time with Adolescent Depression and Anxiety')
r.font.name = 'Times New Roman'; r.font.size = Pt(13); r.italic = True

# ========== 3. THÔNG TIN THƯ MỤC ==========
add_heading(doc, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(doc, [
    ('Tiêu đề gốc', 'Cross-sectional and Longitudinal Associations of Screen Time with Adolescent Depression and Anxiety'),
    ('Tiêu đề dịch', 'Mối liên quan cắt ngang và dọc giữa thời gian sử dụng màn hình với trầm cảm và lo âu ở thanh thiếu niên'),
    ('Tác giả', 'Sophie H. Li, Philip J. Batterham, Alexis E. Whitton, Kate Maston, Asaduzzaman Khan, Helen Christensen, Aliza Werner-Seidler'),
    ('Cơ quan', 'Black Dog Institute, Đại học New South Wales, Sydney, Úc; Đại học Quốc gia Úc, Canberra; Đại học Queensland'),
    ('Tạp chí', 'British Journal of Clinical Psychology (BJCP) (Q1, IF ≈ 3,0)'),
    ('Thông tin xuất bản', 'Vol. 64, pp. 873–887, 2025, 15 trang'),
    ('DOI', '10.1111/bjc.12547'),
    ('Loại nghiên cứu', 'Nghiên cứu thuần tập tiến cứu (prospective cohort) — cắt ngang + dọc 12 tháng'),
    ('Mẫu', '4.058 VTN Úc (tuổi TB = 13,9; SD = 0,58), 134 trường trung học'),
])

add_page_ref(doc, '873–887', 'British Journal of Clinical Psychology', 'Vol. 64, 2025')

# ========== TÓM TẮT ==========
add_heading(doc, 'TÓM TẮT', 2)
add_p(doc, 'Mục tiêu: Mối quan hệ giữa thời gian sử dụng màn hình (screen time) và sức khỏe tâm thần ở thanh thiếu niên đang bị tranh cãi trong tài liệu khoa học, với các nghiên cứu dọc còn thiếu. Nghiên cứu này kiểm tra mối liên quan cắt ngang và dọc giữa screen time với trầm cảm và lo âu, cùng ảnh hưởng của sử dụng mạng xã hội tiêu cực (maladaptive social media use) và giới tính.')

p = doc.add_paragraph()
r = p.add_run('Phương pháp: Phân tích mẫu 4.058 VTN (tuổi TB = 13,9) từ 134 trường trung học Úc, thuộc Nghiên cứu Future Proofing — nghiên cứu thuần tập tiến cứu 5 năm. Mô hình hỗn hợp tuyến tính (linear mixed models) sử dụng dữ liệu Thời điểm 1 và Thời điểm 2 (theo dõi 12 tháng) để kiểm tra mối liên quan cắt ngang và dọc của screen time với trầm cảm (PHQ-A) và lo âu (CAS-8), cùng ảnh hưởng của sử dụng MXH tiêu cực và giới tính.')
r.font.name = 'Times New Roman'; r.font.size = Pt(12)
r.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

p = doc.add_paragraph()
r = p.add_run('Kết quả: Screen time có liên quan đáng kể với trầm cảm và lo âu ở phân tích cắt ngang. Tuy nhiên, trong phân tích dọc, chỉ có mối liên quan YẾU với trầm cảm sau 12 tháng (b = 0,15 mỗi giờ tăng thêm, p = 0,007), và KHÔNG có liên quan đáng kể với lo âu sau 12 tháng. Giới tính và sử dụng MXH tiêu cực KHÔNG có ảnh hưởng đáng kể lên các mối liên quan này.')
r.font.name = 'Times New Roman'; r.font.size = Pt(12)
r.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

add_p(doc, 'Kết luận: Phát hiện cho thấy mối liên quan cắt ngang đáng kể nhưng mối liên quan dọc yếu hoặc không đáng kể giữa screen time với trầm cảm và lo âu ở VTN. Cần thận trọng khi kết luận screen time GÂY RA giảm SKTT — mối quan hệ có thể hai chiều.')

# ========== TÓM TẮT ĐÁNH GIÁ NHANH ==========
add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
bullets = [
    'Nghiên cứu DỌC (longitudinal) — hiếm trong lĩnh vực screen time, cho phép kiểm tra chiều hướng nhân quả.',
    'Phát hiện QUAN TRỌNG và NGƯỢC LẠI với giả định phổ biến: screen time chỉ liên quan YẾU với trầm cảm sau 12 tháng, KHÔNG liên quan với lo âu dọc.',
    'Mẫu lớn (n = 4.058), 134 trường, thiết kế thuần tập tiến cứu 5 năm — Future Proofing Study.',
    'PHQ-A (α = 0,88) và CAS-8 (α = 0,90) — công cụ đo có độ tin cậy cao.',
    'BJCP Q1 — tạp chí tâm lý lâm sàng uy tín Anh quốc.',
]
for b in bullets:
    add_p(doc, f'• {b}')

add_p(doc, 'Hạn chế:', bold=True)
limits = [
    'Chỉ dữ liệu Úc — cần xác nhận ở châu Á/Việt Nam.',
    'Screen time đo bằng 1 câu tự báo cáo — không phân biệt loại nội dung.',
    'Theo dõi chỉ 12 tháng — có thể cần thời gian dài hơn để phát hiện tác động.',
    'Mẫu gender diverse nhỏ (n = 88) — hạn chế phân tích nhóm này.',
]
for b in limits:
    add_p(doc, f'• {b}')

add_p(doc, 'Hướng cải thiện:', bold=True)
improve = [
    'Nghiên cứu dọc tương tự tại VN với thời gian theo dõi dài hơn (2–3 năm).',
    'Phân biệt loại screen time (MXH, game, học tập) — không gộp chung.',
    'RCT giảm screen time tại trường THCS/THPT VN để kiểm tra nhân quả.',
]
for b in improve:
    add_p(doc, f'• {b}')

# ========== GIỚI THIỆU ==========
add_page_ref(doc, '873–875', 'BJCP', 'Vol. 64, 2025')
add_heading(doc, '1. GIỚI THIỆU', 2)

add_p(doc, 'Sức khỏe tâm thần của thanh thiếu niên (10–19 tuổi, WHO, 2023) đang xấu đi. Một tổng quan gần đây cho thấy tỷ lệ triệu chứng trầm cảm ở VTN tăng trong 20 năm qua trên toàn cầu (Shorey et al., 2022). Một số nhà khoa học suy đoán rằng những thay đổi này liên quan đến việc tiếp cận và sử dụng công nghệ số tăng (Twenge & Campbell, 2019).')

add_p(doc, 'Mối liên quan tích cực giữa thời lượng sử dụng công nghệ số, thường gọi là "thời gian sử dụng màn hình" (screen time), và giảm sức khỏe tâm thần đã được ghi nhận. Hai tổng quan gần đây tìm thấy mối liên quan nhỏ và trung bình giữa screen time với trầm cảm và các triệu chứng SKTT khác (Sanders et al., 2024; Stiglic & Viner, 2019). Cơ chế đề xuất: screen time thay thế các hành vi thích ứng (hoạt động thể chất, ngủ, tương tác trực tiếp), giảm tự đánh giá bản thân qua so sánh xã hội tiêu cực, và tiếp xúc với nội dung không phù hợp và bắt nạt mạng.')

add_p(doc, 'Tuy nhiên, các nghiên cứu về tác động screen time lên SKTT đa phần là cắt ngang — không thể xác lập chiều hướng quan hệ hay nhân quả (Heffer et al., 2019). Trong tổng quan các NC dọc, Tang et al. (2021) tìm thấy mối quan hệ tích cực NHỎ giữa screen time tổng thể và trầm cảm sau, và ÍT bằng chứng hỗ trợ liên quan với lo âu.')

add_p(doc, 'Sử dụng mạng xã hội tiêu cực (maladaptive social media use) được định nghĩa là xu hướng dùng MXH để thực hiện so sánh xã hội tiêu cực, ảnh hưởng đến đánh giá giá trị bản thân, sự chấp nhận và hòa nhập (Smith et al., 2013). Đây được coi là cơ chế tiềm năng đằng sau mối liên quan screen time–SKTT. Tuy nhiên, cần phân biệt sử dụng MXH tiêu cực với các hình thức khác (kết nối xã hội, điều tiết cảm xúc, tìm kiếm trợ giúp) — những hình thức được liên kết với SKTT tích cực.')

add_p(doc, 'Nghiên cứu hiện tại nhằm: (1) kiểm tra mối liên quan cắt ngang và dọc giữa screen time với trầm cảm và lo âu, (2) đánh giá ảnh hưởng điều tiết của sử dụng MXH tiêu cực và giới tính, sử dụng dữ liệu từ nghiên cứu thuần tập lớn VTN Úc.')

# ========== PHƯƠNG PHÁP ==========
add_page_ref(doc, '876–877', 'BJCP', 'Vol. 64, 2025')
add_heading(doc, '2. PHƯƠNG PHÁP', 2)

add_p(doc, '2.1. Thiết kế, bối cảnh và đồng ý', bold=True)
add_p(doc, 'Dữ liệu từ Nghiên cứu Future Proofing — nghiên cứu thuần tập tiến cứu 5 năm (Werner-Seidler et al., 2023). Thu thập từ học sinh 134 trường trung học Úc tại thời điểm ban đầu (3 đợt: tháng 3/2019 – tháng 3/2022) và theo dõi 12 tháng (tháng 3/2021 – tháng 3/2023). Đồng ý bằng văn bản từ phụ huynh. Phê duyệt đạo đức từ Đại học NSW (HC180836).')

add_p(doc, '2.2. Người tham gia', bold=True)
add_p(doc, 'Tuyển chọn từ tháng 3/2019 đến tháng 3/2022. Tất cả trường trung học công lập, độc lập NSW và trường Công giáo đủ điều kiện được mời. VTN lớp 8 (Year 8) tại các trường tham gia được mời, chỉ cần smartphone iOS/Android và số điện thoại. Mẫu ban đầu: 6.388 → loại 559 (không hoàn thành T1), 109 (không chọn giới tính), 1.662 (không hoàn thành T2) → mẫu cuối: 4.058.')

add_p(doc, '2.3. Thang đo', bold=True)
add_p(doc, 'Trầm cảm: Bảng hỏi Sức khỏe Bệnh nhân dành cho VTN (PHQ-A — Patient Health Questionnaire for Adolescents; Johnson et al., 2002), phiên bản VTN của PHQ-9. Đo mức độ trầm cảm theo DSM-IV, 0–27 điểm. Cronbach α = 0,88.')
add_p(doc, 'Lo âu: Thang đo Lo âu Trẻ em Rút gọn (CAS-8 — Children\'s Anxiety Scale Short-Form), 8 mục đo lo âu tổng quát và lo âu xã hội, 0–24 điểm. Cronbach α = 0,90.')
add_p(doc, 'Screen time: 1 câu tự báo cáo: "Bạn dành bao lâu trên màn hình ngày bình thường (không tính bài tập)? Bao gồm điện thoại, máy tính, iPad, TV, game...". Trả lời: 1 (0–1 giờ) đến 6 (5+ giờ).')
add_p(doc, 'Sử dụng MXH tiêu cực: Chỉnh sửa từ Thang Sử dụng Facebook Tiêu cực (Smith et al., 2013), 7 mục đo xu hướng so sánh xã hội tiêu cực trên MXH. Tổng 7–49 điểm.')

add_p(doc, '2.4. Phân tích', bold=True)
add_p(doc, 'SPSS phiên bản 26. Mô hình hỗn hợp tuyến tính (linear mixed models) với trường học là tác động ngẫu nhiên. Ba mô hình cho mỗi biến kết quả:')
add_p(doc, '• Mô hình 1: Screen time (tác động cố định) → PHQ-A hoặc CAS-8')
add_p(doc, '• Mô hình 2: + sử dụng MXH tiêu cực + giới tính')
add_p(doc, '• Mô hình 3: + tương tác screen time × MXH tiêu cực + screen time × giới tính')
add_p(doc, 'Phân tích cắt ngang (T1→T1) và dọc (T1→T2, kiểm soát T1). Ngưỡng ý nghĩa: p < 0,01 (hiệu chỉnh Bonferroni).')

# ========== KẾT QUẢ ==========
add_page_ref(doc, '878–882', 'BJCP', 'Vol. 64, 2025')
add_heading(doc, '3. KẾT QUẢ', 2)

add_p(doc, '3.1. Đặc điểm người tham gia', bold=True)

# Bảng 1: Đặc điểm mẫu
add_heading(doc, 'Bảng 1. Đặc điểm người tham gia theo giới tính', 3)
add_table(doc,
    ['Đặc điểm', 'Nữ (n=2.121)', 'Nam (n=1.798)', 'Đa dạng giới (n=88)', 'Tổng (n=4.058)'],
    [
        ['Tuổi TB (SD)', '13,85 (0,59)', '13,97 (0,54)', '13,68 (0,51)', '13,91 (0,58)'],
        ['Sinh tại Úc', '92,5%', '92,3%', '92,1%', '92,4%'],
        ['Có chẩn đoán SKTT', '17,4%', '12,7%', '29,5%', '16,8%'],
        ['Chẩn đoán trầm cảm', '3,4%', '2,4%', '12,9%', '3,3%'],
        ['Screen time 0–1h', '2,9%', '3,4%', '2,2%', '3,1%'],
        ['Screen time 1–2h', '11,3%', '13,3%', '5,0%', '12,0%'],
        ['Screen time 2–3h', '18,6%', '22,8%', '9,4%', '20,1%'],
        ['Screen time 3–4h', '20,5%', '21,7%', '14,4%', '20,8%'],
        ['Screen time 4–5h', '18,4%', '15,2%', '17,3%', '19,9%'],
        ['Screen time 5+h', '28,4%', '23,5%', '51,8%', '27,1%'],
    ],
    widths=[3.5, 3.0, 3.0, 3.0, 3.0]
)

add_p(doc, '3.2. Mối liên quan CẮT NGANG (cùng thời điểm T1)', bold=True)

add_p(doc, 'TRẦM CẢM (PHQ-A):', bold=True)
add_p(doc, '• Mô hình 1: Mỗi giờ screen time tăng thêm → tăng 1,25 điểm PHQ-A (KTC 95%: 1,64–2,83, p < 0,001). Ý nghĩa lâm sàng: người dùng 5+ giờ/ngày có PHQ-A cao hơn ≈ 6,25 điểm so với người dùng 0–1 giờ.')
add_p(doc, '• Mô hình 2: Screen time vẫn có ý nghĩa (b = 1,05, p < 0,001). MXH tiêu cực cũng có ý nghĩa (b = 0,12, p < 0,001). Giới tính: đa dạng giới cao nhất (b = 6,73, p < 0,001), nam thấp hơn nữ (b = −2,58, p < 0,001).')
add_p(doc, '• Mô hình 3: Screen time × giới tính có ý nghĩa — nam có liên quan YẾU hơn nữ (b = −0,45, p < 0,001).')

add_p(doc, 'LO ÂU (CAS-8):', bold=True)
add_p(doc, '• Mô hình 1: Mỗi giờ screen time tăng → tăng 0,72 điểm CAS-8 (KTC 95%: 0,58–0,87, p < 0,001). Người 5+ giờ → cao hơn ≈ 3,6 điểm so với 0–1 giờ.')
add_p(doc, '• Mô hình 2: Screen time vẫn có ý nghĩa (b = 0,53, p < 0,001). Đa dạng giới cao hơn nữ (b = 3,26, p < 0,001), nam thấp hơn nữ (b = −3,87, p < 0,001).')
add_p(doc, '• Mô hình 3: Các tương tác KHÔNG có ý nghĩa.')

add_p(doc, '3.3. Mối liên quan DỌC (T1 screen time → T2 triệu chứng sau 12 tháng)', bold=True)

add_p(doc, 'TRẦM CẢM (PHQ-A) — kết quả then chốt:', bold=True)
add_p(doc, '• Mô hình 1: Mối liên quan nhỏ nhưng có ý nghĩa — mỗi giờ screen time T1 → tăng 0,15 điểm PHQ-A T2 (KTC 95%: 0,04–0,26, p = 0,007). Ý nghĩa lâm sàng: chênh lệch chỉ 0,75 điểm giữa người 0–1h và 5+h.')
add_p(doc, '• Mô hình 2: Vẫn có ý nghĩa (b = 0,14, p = 0,008). MXH tiêu cực KHÔNG có ý nghĩa (p = 0,490).')
add_p(doc, '• Mô hình 3: Screen time KHÔNG còn ý nghĩa (p = 0,028 > 0,01 sau hiệu chỉnh Bonferroni). Các tương tác KHÔNG có ý nghĩa.')

add_p(doc, 'LO ÂU (CAS-8):', bold=True)
add_p(doc, '• Mô hình 1: KHÔNG có ý nghĩa (p = 0,443). Screen time T1 KHÔNG dự báo lo âu T2.')
add_p(doc, '• Mô hình 2–3: Chỉ giới tính có ý nghĩa (nam thấp hơn nữ). Screen time và MXH tiêu cực KHÔNG có ý nghĩa.')

# Bảng 2: Tổng hợp kết quả chính
add_heading(doc, 'Bảng 2. Tổng hợp kết quả chính — Cắt ngang vs Dọc', 3)
add_table(doc,
    ['Phân tích', 'Biến kết quả', 'Screen time b', 'p', 'Ý nghĩa lâm sàng (5+h vs 0–1h)'],
    [
        ['Cắt ngang (T1→T1)', 'Trầm cảm PHQ-A', '1,25/giờ', '<0,001', '+6,25 điểm — ĐÁNG KỂ'],
        ['Cắt ngang (T1→T1)', 'Lo âu CAS-8', '0,72/giờ', '<0,001', '+3,60 điểm — ĐÁNG KỂ'],
        ['Dọc (T1→T2)', 'Trầm cảm PHQ-A', '0,15/giờ', '0,007', '+0,75 điểm — RẤT NHỎ'],
        ['Dọc (T1→T2)', 'Lo âu CAS-8', '—', '0,443', 'KHÔNG có ý nghĩa'],
    ],
    widths=[3.5, 3.0, 2.5, 1.5, 5.0]
)
add_p(doc, 'Ghi chú: b = hệ số hồi quy (Mô hình 1, điều chỉnh cho clustering trường). Dọc kiểm soát cho T1 scores. Ngưỡng p < 0,01 (Bonferroni).', size=9, italic=True)

# Insert Figure 1
fig1_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Charts', 'ScreenTime_Fig1_slopes.png')
if os.path.exists(fig1_path):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(fig1_path, width=Cm(14))
    pc = doc.add_paragraph()
    pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = pc.add_run('Hình 1. Đường hồi quy đơn giản phản ánh trầm cảm T1 (a), lo âu T1 (b), trầm cảm 12 tháng (c) và lo âu 12 tháng (d) theo screen time cho nam, nữ và đa dạng giới. (Figure 1. Simple slopes reflecting depression and anxiety as a function of screen time by gender.)')
    r.font.name = 'Times New Roman'; r.font.size = Pt(9); r.italic = True

# Bảng 3: Giới tính
add_heading(doc, 'Bảng 3. Ảnh hưởng của giới tính lên trầm cảm và lo âu', 3)
add_table(doc,
    ['Biến kết quả', 'Thời điểm', 'Đa dạng giới vs Nữ', 'Nam vs Nữ', 'Tương tác Screen×Giới'],
    [
        ['Trầm cảm PHQ-A', 'T1 (cắt ngang)', 'b = 6,73***', 'b = −2,58***', 'Nam YẾU hơn (b = −0,45***)'],
        ['Lo âu CAS-8', 'T1 (cắt ngang)', 'b = 3,26***', 'b = −3,87***', 'Không có ý nghĩa'],
        ['Trầm cảm PHQ-A', 'T2 (dọc)', 'b = 1,60***', 'b = −2,22***', 'Không có ý nghĩa'],
        ['Lo âu CAS-8', 'T2 (dọc)', 'b = 2,44***', 'b = −3,87***', 'Không có ý nghĩa'],
    ],
    widths=[3.0, 2.5, 3.5, 3.0, 4.0]
)

# Insert Figure 2
fig2_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Charts', 'ScreenTime_Fig2_SMU.png')
if os.path.exists(fig2_path):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(fig2_path, width=Cm(14))
    pc = doc.add_paragraph()
    pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = pc.add_run('Hình 2. Đường hồi quy đơn giản phản ánh trầm cảm và lo âu theo screen time cho mức sử dụng MXH tiêu cực thấp vs cao. Thấp = dưới trung vị 17, cao = trên trung vị. (Figure 2. Simple slopes for low vs high maladaptive social media use.)')
    r.font.name = 'Times New Roman'; r.font.size = Pt(9); r.italic = True

# ========== THẢO LUẬN ==========
add_page_ref(doc, '882–884', 'BJCP', 'Vol. 64, 2025')
add_heading(doc, '4. THẢO LUẬN', 2)

add_p(doc, 'Phát hiện then chốt: screen time có liên quan đáng kể với trầm cảm và lo âu ở phân tích cắt ngang, nhưng tác động KHÔNG duy trì sau 12 tháng. Cụ thể, mối liên quan dọc với trầm cảm RẤT NHỎ (mỗi giờ tăng chỉ 0,1–0,3 điểm PHQ-A — chưa đến 1 điểm chênh lệch giữa 0–1h và 5+h). KHÔNG có mối liên quan dọc với lo âu.')

add_p(doc, 'Mối liên quan dọc yếu hơn hoặc không đáng kể gợi ý hiệu ứng HAI CHIỀU: triệu chứng tăng ảnh hưởng đến hành vi dùng màn hình nhiều bằng chiều ngược lại. Ví dụ, VTN có triệu chứng rút lui xã hội và khó ngủ (APA, 2013) có thể tăng screen time. Nghiên cứu mới cho thấy người ta sử dụng công nghệ số có chủ đích để điều chỉnh tâm trạng (Wadley et al., 2020), bao gồm xoa dịu tâm trạng tiêu cực.')

add_p(doc, 'Trái với dự đoán, giới tính và sử dụng MXH tiêu cực KHÔNG có ảnh hưởng mạnh lên mối liên quan screen time–triệu chứng. Kết quả cho thấy MXH tiêu cực chỉ liên quan YẾU với trầm cảm T1, KHÔNG liên quan với trầm cảm T2, lo âu T1 hay T2. Điều này trái với tài liệu trước — có thể do VTN ngày nay sử dụng MXH ít có hại hơn (giao tiếp bạn bè thay vì so sánh xã hội), hoặc thang đo (dựa trên Facebook 2013) chưa nắm bắt được các hành vi mới (toxic positivity, doom scrolling).')

add_p(doc, 'Hạn chế:', bold=True)
add_p(doc, '• Screen time đo bằng 1 câu tự báo cáo — bỏ qua nhiều loại khác nhau (game, MXH, học tập) và dựa trên hồi tưởng.')
add_p(doc, '• Tập trung hẹp vào yếu tố điều tiết — chưa xem xét bắt nạt, kết nối xã hội, rối loạn giấc ngủ, kiểm soát của cha mẹ.')
add_p(doc, '• Loại bỏ dữ liệu thiếu → mẫu có thể thiên lệch (tuy sự khác biệt nhỏ).')
add_p(doc, '• Mẫu gender diverse nhỏ (n = 88).')
add_p(doc, '• Không đo biến thiên ngắn hạn hoặc tính ổn định của screen time theo thời gian.')

# ========== KẾT LUẬN ==========
add_heading(doc, '5. KẾT LUẬN', 2)
add_p(doc, 'Nghiên cứu quy mô lớn này xác định mối liên quan cắt ngang đáng kể nhưng chỉ mối liên quan dọc yếu hoặc không đáng kể giữa screen time với trầm cảm và lo âu. Giới tính và sử dụng MXH tiêu cực không có ảnh hưởng đáng kể. Kết quả cho thấy mối quan hệ PHỨC TẠP hơn giả định thông thường — có thể hai chiều, và cần nghiên cứu thực nghiệm để xác lập nhân quả. Đặc biệt cần RCT với các hình thức screen time cụ thể (doom scrolling, passive scrolling, tương tác bạn bè) để xác định tác động nhân quả lên SKTT.')

# ========== TÀI LIỆU THAM KHẢO ==========
add_heading(doc, 'TÀI LIỆU THAM KHẢO', 2)
refs = [
    'Beyens, I., et al. (2020). The effect of social media on well-being differs from adolescent to adolescent. Scientific Reports, 10(1), 10763.',
    'Boers, E., et al. (2019). Association of screen time and depression in adolescence. JAMA Pediatrics, 173(9), 853–859.',
    'Heffer, T., et al. (2019). The longitudinal association between social-media use and depressive symptoms. Clinical Psychological Science, 7(3), 462–470.',
    'Johnson, J.G., et al. (2002). The Patient Health Questionnaire for Adolescents. Journal of Adolescent Health, 30(3), 196–204.',
    'Li, S.H., et al. (2025). Cross-sectional and longitudinal associations of screen time with adolescent depression and anxiety. BJCP, 64, 873–887.',
    'Sanders, T., et al. (2024). Dose–response associations between screen use and mental health symptoms. JAMA Pediatrics, 178, 1210–1217.',
    'Smith, R.H., et al. (2013). Dispositional envy. Personality and Social Psychology Bulletin, 39(7), 911–921.',
    'Stiglic, N. & Viner, R.M. (2019). Effects of screentime on the health and well-being of children and adolescents. BMJ Open, 9(1), e023191.',
    'Tang, S., et al. (2021). Digital screen time and its effect on preschoolers\' behavior in China. Computers in Human Behavior, 121, 106817.',
    'Werner-Seidler, A., et al. (2023). The Future Proofing Study. BMJ Open, 13(4), e066249.',
]
for ref in refs:
    add_p(doc, ref, size=10)
add_p(doc, '(Xem danh mục đầy đủ trong bài gốc — 40+ tài liệu tham khảo)', size=10, italic=True)

# ========== BẢNG VIẾT TẮT ==========
add_abbreviation_table(doc, [
    ('PHQ-A', 'Patient Health Questionnaire for Adolescents — Bảng hỏi Sức khỏe Bệnh nhân dành cho VTN'),
    ('CAS-8', "Children's Anxiety Scale Short-Form — Thang Lo âu Trẻ em Rút gọn 8 mục"),
    ('MXH', 'Mạng xã hội (Social Media)'),
    ('SMU', 'Social Media Use — Sử dụng mạng xã hội'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('VTN', 'Vị thành niên (thanh thiếu niên)'),
    ('T1', 'Thời điểm 1 (baseline)'),
    ('T2', 'Thời điểm 2 (theo dõi 12 tháng)'),
    ('RCT', 'Randomized Controlled Trial — Thử nghiệm ngẫu nhiên có đối chứng'),
    ('KTC', 'Khoảng tin cậy (Confidence Interval)'),
    ('DSM-IV', 'Diagnostic and Statistical Manual of Mental Disorders, 4th Edition'),
    ('NSW', 'New South Wales, Úc'),
])

# ========== PHẢN BIỆN ==========
add_red_heading(doc, 'QUAN ĐIỂM PHẢN BIỆN / CRITICAL REVIEW')

add_red(doc, 'Điểm mạnh:', bold=True)
for s in [
    'BJCP Q1 — tạp chí tâm lý lâm sàng uy tín. Thiết kế DỌC — hiếm trong NC screen time.',
    'Mẫu lớn (n = 4.058), 134 trường, đại diện tốt cho VTN Úc.',
    'Thuộc Future Proofing Study — nghiên cứu thuần tập 5 năm, thiết kế chặt chẽ.',
    'Công cụ đo tin cậy cao: PHQ-A (α = 0,88), CAS-8 (α = 0,90).',
    'Phân tích đa mô hình (3 mô hình) — kiểm soát dần cho MXH tiêu cực và giới tính.',
]:
    add_red(doc, f'• {s}')

add_red(doc, 'Hạn chế chi tiết:', bold=True)
for s in [
    'Screen time chỉ đo bằng 1 câu tự báo cáo — KHÔNG phân biệt MXH, game, học tập, TV. Đây là hạn chế chung của đa số NC.',
    'Chỉ Úc — quốc gia phát triển, văn hóa phương Tây. Kết quả có thể KHÁC ở VN/châu Á (áp lực học tập khác, kiểm soát cha mẹ khác).',
    'Theo dõi chỉ 12 tháng — có thể quá ngắn để phát hiện tác động tích lũy dài hạn.',
    'Thang MXH tiêu cực dựa trên Facebook (2013) — lỗi thời cho TikTok, Instagram Reels 2024.',
    'Mẫu gender diverse quá nhỏ (n = 88 / 4.058 = 2,2%) — không đủ sức mạnh thống kê.',
    'Không đo: giấc ngủ, hoạt động thể chất, bắt nạt mạng — các biến giao thoa quan trọng.',
    'Phát hiện dọc YẾU có thể do: (a) thực sự yếu, hoặc (b) do đo screen time không chính xác bằng tự báo cáo.',
]:
    add_red(doc, f'• {s}')

add_red(doc, 'Khoảng trống nghiên cứu / Research Gap:', bold=True)
for s in [
    'CHƯA CÓ nghiên cứu dọc tương tự tại Việt Nam hoặc ASEAN — đây là GAP lớn.',
    'Cần phân biệt LOẠI screen time (MXH, game, học tập) — Hoàng Trung Học 2025 đã bước đầu (game điện tử Beta = 0,176).',
    'Cần RCT giảm screen time tại trường VN — kết hợp với JAMA 2024 (đã chứng minh RCT hiệu quả).',
    'Chưa rõ cơ chế: screen time → SKTT hay SKTT → screen time? Cần thiết kế thực nghiệm (ecological momentary assessment) để giải quyết.',
    'Liên hệ VN: Norway 2025 tìm thấy MXH giải thích xu hướng tăng, nhưng bài này tìm thấy tác động dọc YẾU — mâu thuẫn cần giải quyết. Có thể do đo khác nhau (thời gian vs nội dung).',
]:
    add_red(doc, f'• {s}')

# ========== SAVE ==========
outpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), '22_ScreenTime_2025_BJCP.docx')
doc.save(outpath)
print(f'Saved: {outpath}')
