# -*- coding: utf-8 -*-
"""Dịch toàn văn QT23 - Mojtabai & Olfson 2024 - JAACAP - Trends in Mental Disorders US"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()

# ========== 1. LINK ==========
add_p(doc, 'Link bài báo gốc: https://doi.org/10.1016/j.jaac.2024.08.008', size=10)

# ========== 2. TIÊU ĐỀ ==========
add_heading(doc, 'Xu hướng rối loạn tâm thần ở trẻ em và thanh thiếu niên nhận điều trị trong hệ thống sức khỏe tâm thần công lập Hoa Kỳ', 1)
h2 = doc.add_paragraph()
r = h2.add_run('Trends in Mental Disorders in Children and Adolescents Receiving Treatment in the State Mental Health System')
r.font.name = 'Times New Roman'; r.font.size = Pt(13); r.italic = True

# ========== 3. THÔNG TIN THƯ MỤC ==========
add_heading(doc, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(doc, [
    ('Tiêu đề gốc', 'Trends in Mental Disorders in Children and Adolescents Receiving Treatment in the State Mental Health System'),
    ('Tiêu đề dịch', 'Xu hướng rối loạn tâm thần ở trẻ em và thanh thiếu niên nhận điều trị trong hệ thống sức khỏe tâm thần công lập Hoa Kỳ'),
    ('Tác giả', 'Ramin Mojtabai, MD, PhD, MPH; Mark Olfson, MD, MPH'),
    ('Cơ quan', 'Tulane University School of Medicine, New Orleans; Columbia University, New York'),
    ('Tạp chí', 'Journal of the American Academy of Child & Adolescent Psychiatry — JAACAP (Q1, IF ≈ 11,0)'),
    ('Thông tin xuất bản', 'Vol. 64, No. 8, August 2025, 15 trang'),
    ('DOI', '10.1016/j.jaac.2024.08.008'),
    ('Loại nghiên cứu', 'Phân tích xu hướng — dữ liệu hành chính quốc gia (administrative data)'),
    ('Mẫu', '13.684.154 hồ sơ trẻ em/VTN (≤17 tuổi) trong hệ thống SKTT công lập Mỹ, 2013–2021'),
])

add_page_ref(doc, '1–15', 'JAACAP', 'Vol. 64(8), 2025')

# ========== TÓM TẮT ==========
add_heading(doc, 'TÓM TẮT', 2)
add_p(doc, 'Mục tiêu: Kiểm tra xu hướng gần đây trong chẩn đoán lâm sàng của trẻ em và thanh thiếu niên nhận điều trị tại các dịch vụ SKTT được tài trợ công trong hệ thống y tế Hoa Kỳ.')

p = doc.add_paragraph()
r = p.add_run('Phương pháp: Dữ liệu về trẻ em và VTN (≤17 tuổi) nhận điều trị từ dịch vụ SKTT tài trợ công, ghi nhận trong Mental Health Client-Level Data (MH-CLD) 2013–2021 (tổng 13.684.154 hồ sơ), được sử dụng để kiểm tra xu hướng thời gian trong tỷ lệ các rối loạn tâm thần khác nhau. Xu hướng được kiểm tra tổng thể và phân tầng theo tuổi, giới, chủng tộc/dân tộc, và loại hình dịch vụ, tập trung vào chương trình cộng đồng.')
r.font.name = 'Times New Roman'; r.font.size = Pt(12)
r.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

p = doc.add_paragraph()
r = p.add_run('Kết quả: Tỷ lệ rối loạn lo âu TĂNG từ 9,6% năm 2013 lên 19,2% năm 2021 (AOR = 2,17; KTC 95%: 1,85–2,55; p < 0,001). Rối loạn liên quan sang chấn/stress tăng từ 22,7% lên 27,4% (AOR = 1,31; p = 0,004). Rối loạn trầm cảm tăng từ 13,4% lên 17,0% (AOR = 1,20; p = 0,04). Trong cùng giai đoạn, rối loạn lưỡng cực GIẢM gần 8 lần từ 10,0% xuống 1,3% (AOR = 0,07; p < 0,001). Rối loạn hành vi và chống đối cũng giảm.')
r.font.name = 'Times New Roman'; r.font.size = Pt(12)
r.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

add_p(doc, 'Kết luận: Cấu trúc chẩn đoán tâm thần ở trẻ em trong hệ thống SKTT công thay đổi đáng kể trong thập kỷ qua. Tăng rối loạn lo âu và trầm cảm song song xu hướng dân số chung, nhấn mạnh nhu cầu ngày càng lớn trong việc xác định và điều trị hiệu quả các rối loạn này.')

# ========== TÓM TẮT ĐÁNH GIÁ NHANH ==========
add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
for b in [
    'JAACAP — tạp chí hàng đầu thế giới về tâm thần trẻ em (Q1, IF ≈ 11,0).',
    'Mẫu CỰC LỚN: 13,7 triệu hồ sơ lâm sàng — dữ liệu chẩn đoán thực tế (không phải sàng lọc).',
    'Lo âu TĂNG GẤP ĐÔI trong 8 năm (AOR = 2,17) — xu hướng mạnh nhất, đặc biệt ở VTN 15–17 tuổi (AOR = 2,93).',
    'Lưỡng cực GIẢM 8 lần — phản ánh thay đổi thực hành chẩn đoán (từ hành vi hóa sang nội hóa).',
    'Dữ liệu phân tầng theo giới, tuổi, chủng tộc — phân tích rất chi tiết.',
]:
    add_p(doc, f'• {b}')

add_p(doc, 'Hạn chế:', bold=True)
for b in [
    'Chỉ trẻ ĐANG nhận điều trị trong hệ thống công — thiên lệch chọn (không bao gồm trẻ chưa chẩn đoán, trẻ hệ thống tư nhân).',
    'Xu hướng tăng có thể phản ánh tăng phát hiện + tăng thực sự cùng lúc — không tách được.',
    'Chỉ Mỹ — hệ thống SKTT phát triển, khác biệt rất lớn với VN.',
    'Dữ liệu hành chính — chất lượng chẩn đoán phụ thuộc vào thực hành lâm sàng tại từng tiểu bang.',
]:
    add_p(doc, f'• {b}')

add_p(doc, 'Hướng cải thiện:', bold=True)
for b in [
    'Phân tích xu hướng tương tự từ dữ liệu Bệnh viện Nhi TW VN.',
    'So sánh hệ thống y tế Mỹ (82,6% tiếp cận) vs VN (8,4%).',
    'Đánh giá nguyên nhân xu hướng tăng (phát hiện vs thực sự).',
]:
    add_p(doc, f'• {b}')

# ========== GIỚI THIỆU ==========
add_page_ref(doc, '1–2', 'JAACAP', 'Vol. 64(8), 2025')
add_heading(doc, '1. GIỚI THIỆU', 2)

add_p(doc, 'Trong hai thập kỷ qua, ngày càng nhiều bằng chứng cho thấy tỷ lệ ngày càng tăng của nhiều rối loạn tâm thần ở trẻ em và thanh thiếu niên tại Hoa Kỳ. Theo Khảo sát Sức khỏe Quốc gia 2005–2018, trầm cảm, ý tưởng tự tử và các vấn đề nội hóa chiếm tỷ lệ ngày càng tăng trong chăm sóc SKTT ngoại trú VTN, trong khi các vấn đề ngoại hóa (kiểm soát giận dữ, gây gổ) giảm.')

add_p(doc, 'Phân tích dữ liệu MarketScan 2012–2018 (bảo hiểm thương mại) cho thấy: rối loạn lo âu tăng từ 3,7% lên 7,2%, rối loạn trầm cảm từ 1,8% lên 3,1%, ADHD từ 0,9% lên 1,5% ở trẻ em/VTN.')

add_p(doc, 'Nghiên cứu hiện tại kiểm tra xu hướng chẩn đoán lâm sàng của trẻ em/VTN nhận điều trị trong dịch vụ SKTT vận hành hoặc tài trợ công tại Hoa Kỳ từ 2013 đến 2021, tập trung vào chương trình cộng đồng. Các chương trình này do tiểu bang vận hành hoặc nhận tài trợ liên bang qua quy trình tài trợ khối (block grant).')

add_p(doc, 'Hệ thống SKTT công là thành phần thiết yếu của mạng lưới an toàn SKTT, phục vụ trẻ em/VTN có hoàn cảnh khó khăn (nghèo, không bảo hiểm, nhóm thiểu số, vùng nông thôn) — đặc biệt quan trọng vì nhóm này dễ bị tổn thương hơn.')

# ========== PHƯƠNG PHÁP ==========
add_page_ref(doc, '2–4', 'JAACAP', 'Vol. 64(8), 2025')
add_heading(doc, '2. PHƯƠNG PHÁP', 2)

add_p(doc, '2.1. Nguồn dữ liệu', bold=True)
add_p(doc, 'Dữ liệu Cấp Khách hàng Sức khỏe Tâm thần (MH-CLD — Mental Health Client-Level Data), thu thập bởi Cơ quan Dịch vụ Lạm dụng Chất và SKTT (SAMHSA — Substance Abuse and Mental Health Services Administration). MH-CLD ghi nhận đặc điểm nhân khẩu, chẩn đoán, và dịch vụ của người nhận điều trị từ các chương trình SKTT do tiểu bang tài trợ hoặc vận hành. Tất cả 50 tiểu bang và DC báo cáo dữ liệu.')

add_p(doc, '2.2. Mẫu và chẩn đoán', bold=True)
add_p(doc, 'Tổng cộng 13.684.154 hồ sơ trẻ em/VTN ≤17 tuổi (2013–2021). Tám nhóm rối loạn được phân tích: rối loạn lo âu, rối loạn trầm cảm, rối loạn lưỡng cực, rối loạn liên quan sang chấn/stress, ADHD, rối loạn hành vi (conduct disorder), rối loạn chống đối (ODD), và rối loạn phát triển lan tỏa (PDD). Tám nhóm này chiếm 98,1% tổng bệnh nhân.')

add_p(doc, 'Trẻ em/VTN được phân loại có rối loạn nếu chẩn đoán đó được ghi nhận ở vị trí 1, 2, hoặc 3 trong hồ sơ. Do nhiều trẻ có chẩn đoán ở nhiều nhóm, các nhóm KHÔNG loại trừ lẫn nhau.')

add_p(doc, '2.3. Cơ sở điều trị', bold=True)
add_p(doc, '(1) Chương trình cộng đồng (SMHA): trung tâm SKTT cộng đồng, phòng khám ngoại trú, chương trình bán nội trú, chương trình hỗ trợ cộng đồng.')
add_p(doc, '(2) Cơ sở nội trú: tổ chức không phải bệnh viện tâm thần, cung cấp chăm sóc nội trú kết hợp điều trị SKTT.')
add_p(doc, '(3) Bệnh viện tâm thần tiểu bang: bệnh viện nội trú cho bệnh nhân tâm thần.')
add_p(doc, '(4) Cơ sở nội trú khác: nhà cung cấp tư nhân hoặc y tế được cấp phép/hợp đồng qua SMHA.')

add_p(doc, '2.4. Phân tích', bold=True)
add_p(doc, 'Xu hướng thời gian kiểm tra bằng hồi quy logistic (odds ratio — OR và adjusted OR — AOR). Điều chỉnh cho tuổi, giới, chủng tộc/dân tộc. Sai số chuẩn điều chỉnh cho clustering theo tiểu bang. Kiểm tra tác động COVID-19 bằng linear splines (2013–2019 vs 2020–2021). Stata phiên bản 18.')

# ========== KẾT QUẢ ==========
add_page_ref(doc, '4–13', 'JAACAP', 'Vol. 64(8), 2025')
add_heading(doc, '3. KẾT QUẢ', 2)

add_p(doc, '3.1. Đặc điểm mẫu', bold=True)

# Bảng 1: Đặc điểm nhân khẩu
add_heading(doc, 'Bảng 1. Đặc điểm nhân khẩu trẻ em/VTN trong hệ thống SKTT công Mỹ, 2013–2021', 3)
add_table(doc,
    ['Đặc điểm', '2013 (%)', '2021 (%)', 'OR', 'p'],
    [
        ['Nữ', '41,1', '46,5', '1,19', '<0,001'],
        ['Nam', '58,8', '53,4', '0,84', '<0,001'],
        ['<12 tuổi', '49,1', '46,2', '0,92', '0,011'],
        ['12–14 tuổi', '24,5', '26,6', '1,11', '<0,001'],
        ['15–17 tuổi', '26,4', '27,2', '1,01', '0,618'],
        ['Hispanic', '10,2', '15,4', '1,55', '0,010'],
        ['Trắng không Hispanic', '47,8', '39,9', '0,77', '0,022'],
        ['Da đen không Hispanic', '19,5', '14,4', '0,71', '<0,001'],
    ],
    widths=[4.0, 2.0, 2.0, 2.0, 2.0]
)
add_p(doc, 'Ghi chú: N = 1,6–1,9 triệu/năm (tổng 13,7 triệu hồ sơ). OR = Odds Ratio cho xu hướng tuyến tính 2013–2021.', size=9, italic=True)

add_p(doc, '3.2. Xu hướng rối loạn tâm thần (tất cả cơ sở)', bold=True)

# Bảng 2: Xu hướng chính
add_heading(doc, 'Bảng 2. Xu hướng tỷ lệ rối loạn tâm thần — Tất cả cơ sở, 2013–2021', 3)
add_table(doc,
    ['Rối loạn', '2013 (%)', '2021 (%)', 'AOR', 'KTC 95%', 'p', 'Xu hướng'],
    [
        ['Rối loạn lo âu', '9,6', '19,2', '2,17', '1,85–2,55', '<0,001', '↑ TĂNG GẤP ĐÔI'],
        ['Rối loạn trầm cảm', '13,4', '17,0', '1,20', '1,03–1,41', '0,023', '↑ Tăng'],
        ['Sang chấn/stress', '22,7', '27,4', '1,31', '1,09–1,57', '0,004', '↑ Tăng'],
        ['ADHD', '29,8', '25,3', '0,90', '0,77–1,06', '0,208', '— Ổn định'],
        ['Rối loạn lưỡng cực', '10,0', '1,3', '0,07', '0,06–0,09', '<0,001', '↓ GIẢM 8 LẦN'],
        ['Rối loạn hành vi', '9,7', '4,4', '0,42', '0,33–0,55', '<0,001', '↓ Giảm mạnh'],
        ['Rối loạn chống đối', '11,1', '7,8', '0,79', '0,65–0,98', '0,029', '↓ Giảm'],
        ['Rối loạn PT lan tỏa', '3,8', '3,8', '0,98', '0,48–1,99', '0,949', '— Ổn định'],
    ],
    widths=[3.5, 1.5, 1.5, 1.5, 2.5, 1.5, 3.0]
)
add_p(doc, 'Ghi chú: AOR điều chỉnh cho tuổi, giới, chủng tộc/dân tộc. Clustering theo tiểu bang. Chỉ tính chương trình cộng đồng tương tự.', size=9, italic=True)

# Insert Figure 1
fig1_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Charts', 'JAACAP_Fig1_trends.png')
if os.path.exists(fig1_path):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(fig1_path, width=Cm(15))
    pc = doc.add_paragraph()
    pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = pc.add_run('Hình 1. Xu hướng tỷ lệ các rối loạn tâm thần ở trẻ em/VTN. (A) Rối loạn lo âu, trầm cảm, lưỡng cực, sang chấn/stress. (B) ADHD, ODD, rối loạn hành vi, PDD. Phân tích dựa trên MH-CLD, 2013–2021. (Figure 1. Prevalence Trends of Child and Adolescent Psychiatric Disorders.)')
    r.font.name = 'Times New Roman'; r.font.size = Pt(9); r.italic = True

add_p(doc, '3.3. Xu hướng theo TUỔI', bold=True)

# Bảng 3: Xu hướng theo tuổi (lo âu)
add_heading(doc, 'Bảng 3. Xu hướng rối loạn lo âu theo nhóm tuổi (chương trình cộng đồng)', 3)
add_table(doc,
    ['Nhóm tuổi', '2013 (%)', '2021 (%)', 'AOR', 'KTC 95%', 'p', 'Tương tác p'],
    [
        ['<12 tuổi', '9,3', '14,6', '1,66', '1,37–2,02', '<0,001', 'Reference'],
        ['12–14 tuổi', '9,7', '21,1', '2,38', '2,03–2,80', '<0,001', '<0,001'],
        ['15–17 tuổi', '10,1', '25,4', '2,93', '2,51–3,41', '<0,001', '<0,001'],
    ],
    widths=[3.0, 2.0, 2.0, 2.0, 2.5, 2.0, 2.0]
)
add_p(doc, 'Ghi chú: Lo âu tăng mạnh nhất ở 15–17 tuổi (AOR = 2,93 — gần GẤP BA). Tương tác tuổi × thời gian có ý nghĩa (p < 0,001).', size=9, italic=True)

add_p(doc, '3.4. Xu hướng theo GIỚI TÍNH', bold=True)
add_p(doc, 'Xu hướng tăng lo âu và trầm cảm rõ rệt hơn ở NỮ so với nam. Xu hướng giảm rối loạn hành vi và chống đối chủ yếu ở NAM hoặc không có ý nghĩa ở nữ.')

# Bảng 4: Xu hướng theo giới
add_heading(doc, 'Bảng 4. Xu hướng rối loạn lo âu theo giới tính', 3)
add_table(doc,
    ['Giới', '2013 (%)', '2021 (%)', 'AOR', 'Tương tác p'],
    [
        ['Nữ', '11,3', '22,3', '2,24', 'Reference'],
        ['Nam', '8,5', '16,5', '2,14', '0,207 (không khác biệt đáng kể)'],
    ],
    widths=[2.5, 2.0, 2.0, 2.0, 6.0]
)

add_p(doc, '3.5. Xu hướng theo CHỦNG TỘC/DÂN TỘC', bold=True)
add_p(doc, 'Lo âu tăng ở tất cả nhóm: Trắng không Hispanic (11,1% → 23,5%, AOR = 2,41), Da đen (5,1% → 10,8%, AOR = 2,28), Hispanic (11,0% → 17,8%, AOR = 1,64). Đáng chú ý: tỷ lệ lo âu ở trẻ Da đen thấp nhất nhưng tăng nhanh nhất (AOR = 2,28).')

add_p(doc, '3.6. Tác động COVID-19', bold=True)
add_p(doc, 'Phân tích linear splines cho thấy xu hướng tăng lo âu đã diễn ra TRƯỚC COVID (2013–2019). COVID không tạo ra "bước nhảy" đột ngột mà tiếp tục xu hướng có sẵn. Tuy nhiên, lưỡng cực giảm mạnh nhất sau 2016 — có thể liên quan DSM-5 (2013) với DMDD thay thế.')

# ========== THẢO LUẬN ==========
add_page_ref(doc, '8–14', 'JAACAP', 'Vol. 64(8), 2025')
add_heading(doc, '4. THẢO LUẬN', 2)

add_p(doc, 'Phát hiện nổi bật nhất: rối loạn lo âu TĂNG GẤP ĐÔI trong 8 năm (AOR = 2,17) — xu hướng mạnh nhất và nhất quán nhất qua tất cả phân tầng (tuổi, giới, chủng tộc, cơ sở). Lo âu đã vượt qua trầm cảm để trở thành rối loạn tăng nhanh nhất trong hệ thống SKTT trẻ em/VTN Mỹ.')

add_p(doc, 'Tăng lo âu có thể phản ánh nhiều yếu tố: tăng thực sự trong tỷ lệ dân số (song song với dữ liệu cộng đồng), tăng nhận thức và phát hiện (khuyến nghị sàng lọc USPSTF 2022), và có thể tăng stress xã hội (MXH, COVID). Không thể tách riêng các yếu tố này bằng dữ liệu hành chính.')

add_p(doc, 'Giảm lưỡng cực gần 8 lần:', bold=True)
add_p(doc, 'Xu hướng giảm nhất quán nhất qua tất cả phân tầng. Phản ánh thay đổi thực hành chẩn đoán: tỷ lệ chẩn đoán lưỡng cực ở trẻ em tăng 40 lần giữa 1994–2003, gây lo ngại về chẩn đoán quá mức. DSM-5 (2013) giới thiệu Rối loạn Rối loạn Điều hòa Tâm trạng (DMDD — Disruptive Mood Dysregulation Disorder) nhằm giảm chẩn đoán sai lưỡng cực ở trẻ em.')

add_p(doc, 'Tăng sang chấn/stress:', bold=True)
add_p(doc, 'Tăng rối loạn liên quan sang chấn/stress (AOR = 1,31) đáng ngạc nhiên vì nhiều trải nghiệm bất lợi thời thơ ấu (ACEs) giảm trong những năm gần đây. Có thể phản ánh tăng nhận thức và sàng lọc sang chấn, cùng với tác động COVID.')

add_p(doc, 'Giảm rối loạn hành vi và chống đối:', bold=True)
add_p(doc, 'Phản ánh chuyển dịch từ "ngoại hóa" sang "nội hóa" — cả trong thực tế và trong thực hành chẩn đoán. Trẻ em trước đây có thể được chẩn đoán rối loạn hành vi nay được chẩn đoán lo âu hoặc sang chấn.')

add_p(doc, 'Hạn chế nghiên cứu:', bold=True)
add_p(doc, '• Chỉ trẻ đang nhận điều trị trong hệ thống công — thiên lệch chọn. Không bao gồm trẻ chưa chẩn đoán, trẻ hệ thống bảo hiểm tư nhân, hoặc trẻ không tiếp cận dịch vụ.')
add_p(doc, '• Không thể phân biệt tăng thực sự vs tăng phát hiện/thay đổi thực hành chẩn đoán.')
add_p(doc, '• Dữ liệu hành chính — chất lượng chẩn đoán khác nhau giữa tiểu bang.')
add_p(doc, '• Trẻ có thể xuất hiện nhiều lần trong cùng năm hoặc qua nhiều năm — không phải mẫu cá nhân độc lập.')
add_p(doc, '• Giới hạn 3 chẩn đoán đầu tiên — có thể bỏ sót chẩn đoán phụ.')

# ========== KẾT LUẬN ==========
add_heading(doc, '5. KẾT LUẬN', 2)
add_p(doc, 'Cấu trúc chẩn đoán tâm thần ở trẻ em trong hệ thống SKTT công thay đổi đáng kể trong thập kỷ qua. Một số xu hướng có thể phản ánh thay đổi trong thực hành chẩn đoán, nhưng tăng rối loạn lo âu và trầm cảm song song với xu hướng dân số chung — nhấn mạnh nhu cầu ngày càng tăng trong việc xác định và điều trị hiệu quả các rối loạn này ở trẻ em/VTN. Kết quả nhấn mạnh tầm quan trọng của đào tạo lực lượng SKTT trẻ em công trong đánh giá và quản lý rối loạn lo âu, trầm cảm, và sang chấn ngày càng phổ biến.')

# ========== TÀI LIỆU THAM KHẢO ==========
add_heading(doc, 'TÀI LIỆU THAM KHẢO', 2)
refs = [
    'Mojtabai, R. & Olfson, M. (2024). Trends in Mental Disorders in Children and Adolescents Receiving Treatment in the State Mental Health System. JAACAP, 64(8).',
    'Mojtabai, R., et al. (2016). National trends in the prevalence and treatment of depression in adolescents and young adults. Pediatrics, 138(6), e2016-1878.',
    'Twenge, J.M., et al. (2019). Age, period, and cohort trends in mood disorder indicators. J Abnorm Psychol, 128(3), 185–199.',
    'Bommersbach, T.J., et al. (2023). National trends in mental health-related emergency department visits among youth, 2011–2020. JAMA, 329(17), 1469–1477.',
    'Bitsko, R.H., et al. (2022). Mental health surveillance among children — United States, 2013–2019. MMWR Suppl, 71(2), 1–42.',
]
for ref in refs:
    add_p(doc, ref, size=10)
add_p(doc, '(Xem danh mục đầy đủ trong bài gốc — 41 tài liệu tham khảo)', size=10, italic=True)

# ========== BẢNG VIẾT TẮT ==========
add_abbreviation_table(doc, [
    ('MH-CLD', 'Mental Health Client-Level Data — Dữ liệu Cấp Khách hàng Sức khỏe Tâm thần'),
    ('SAMHSA', 'Substance Abuse and Mental Health Services Administration — Cơ quan Dịch vụ Lạm dụng Chất và SKTT'),
    ('SMHA', 'State Mental Health Agency — Cơ quan SKTT Tiểu bang'),
    ('AOR', 'Adjusted Odds Ratio — Tỷ số chênh điều chỉnh'),
    ('ADHD', 'Attention-Deficit/Hyperactivity Disorder — Rối loạn Tăng động Giảm chú ý'),
    ('ODD', 'Oppositional Defiant Disorder — Rối loạn Chống đối'),
    ('DMDD', 'Disruptive Mood Dysregulation Disorder — Rối loạn Rối loạn Điều hòa Tâm trạng'),
    ('PDD', 'Pervasive Developmental Disorders — Rối loạn Phát triển Lan tỏa'),
    ('DSM-5', 'Diagnostic and Statistical Manual of Mental Disorders, 5th Edition'),
    ('USPSTF', 'US Preventive Services Task Force — Lực lượng Dịch vụ Phòng ngừa Hoa Kỳ'),
    ('ACEs', 'Adverse Childhood Experiences — Trải nghiệm Bất lợi Thời thơ ấu'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('VTN', 'Vị thành niên (thanh thiếu niên)'),
])

# ========== PHẢN BIỆN ==========
add_red_heading(doc, 'QUAN ĐIỂM PHẢN BIỆN / CRITICAL REVIEW')

add_red(doc, 'Điểm mạnh:', bold=True)
for s in [
    'JAACAP Q1 IF ≈ 11 — tạp chí hàng đầu thế giới về tâm thần trẻ em.',
    'Mẫu cực lớn: 13,7 TRIỆU hồ sơ — lớn nhất trong tất cả NC của Đề tài.',
    'Dữ liệu CHẨN ĐOÁN LÂM SÀNG (không phải sàng lọc) — phản ánh thực tế dịch vụ.',
    'Xu hướng 8 năm liên tục (2013–2021) — đủ dài để xác định xu hướng ổn định.',
    'Phân tầng theo tuổi, giới, chủng tộc — phân tích rất chi tiết, cho phép so sánh tiểu nhóm.',
]:
    add_red(doc, f'• {s}')

add_red(doc, 'Hạn chế chi tiết:', bold=True)
for s in [
    'THIÊN LỆCH CHỌN LỚN: chỉ trẻ đang nhận điều trị trong hệ thống công. Trẻ chưa chẩn đoán, hệ thống tư nhân, không tiếp cận → bị loại.',
    'Không tách được: tăng thực sự vs tăng phát hiện vs thay đổi thực hành chẩn đoán. Ba yếu tố xảy ra đồng thời.',
    'Chỉ Mỹ — hệ thống SKTT phát triển nhất thế giới. Tại VN nơi chỉ 8,4% tiếp cận dịch vụ, xu hướng có thể hoàn toàn khác.',
    'Dữ liệu hành chính: chất lượng chẩn đoán phụ thuộc thực hành lâm sàng tại từng tiểu bang — không chuẩn hóa.',
    'Trẻ xuất hiện nhiều lần (cùng trẻ qua nhiều năm hoặc nhiều lần/năm) — vi phạm giả định mẫu độc lập.',
    'Không có thông tin về SES, screen time, MXH, stress — chỉ mô tả xu hướng, KHÔNG giải thích nguyên nhân.',
    'Giới hạn 3 chẩn đoán đầu tiên — có thể bỏ sót rối loạn đi kèm (comorbidity).',
]:
    add_red(doc, f'• {s}')

add_red(doc, 'Khoảng trống nghiên cứu / Research Gap:', bold=True)
for s in [
    'VN CHƯA CÓ phân tích xu hướng chẩn đoán tâm thần từ dữ liệu bệnh viện. Bệnh viện Nhi TW có dữ liệu nhưng chưa được phân tích hệ thống.',
    'So sánh khoảng cách tiếp cận: Mỹ 82,6% vs VN 8,4% (V-NAMHS 2022) — tỷ lệ thực ở VN có thể cao hơn nhiều so với báo cáo.',
    'Cần NC đánh giá: xu hướng tăng ở VN có phải do tăng nhận thức hay tăng thực sự?',
    'Bài này + Norway 2025 + GBD 2025 + Korea 2024 → bằng chứng mạnh rằng lo âu tăng TOÀN CẦU — nhưng VN chưa có dữ liệu xu hướng dài hạn.',
    'Giảm lưỡng cực ở Mỹ do DSM-5/DMDD — VN có mô hình tương tự không? Thực hành chẩn đoán ở VN có bị ảnh hưởng bởi DSM-5?',
]:
    add_red(doc, f'• {s}')

# ========== SAVE ==========
outpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), '23_JAACAP_US_Trends_2024.docx')
doc.save(outpath)
print(f'Saved: {outpath}')
