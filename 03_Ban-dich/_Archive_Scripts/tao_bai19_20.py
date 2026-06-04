# -*- coding: utf-8 -*-
"""Bài 19 (Thảo Vi Huế) + Bài 20 (Vĩnh Lộc TPHCM) — nhận xét đầy đủ với số liệu"""
import sys, io, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

# ============================================================
# BÀI 19: HỒ THỊ TRÚC QUỲNH — Lo âu và trầm cảm ở TTN Huế
# ============================================================
d1 = create_doc()
add_heading(d1, 'Lo âu và trầm cảm ở thanh thiếu niên tại Thừa Thiên Huế và mối quan hệ của nó với sự lạc quan', 1)
add_heading(d1, 'Anxiety and Depression in Adolescents in Thua Thien Hue and Their Relationship with Optimism', 2)

add_heading(d1, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d1, [
    ('Tiêu đề', 'Lo âu và trầm cảm ở thanh thiếu niên tại Thừa Thiên Huế và mối quan hệ của nó với sự lạc quan'),
    ('Tác giả', 'Hồ Thị Trúc Quỳnh — Khoa Tâm lý và Giáo dục, ĐH Sư phạm, ĐH Huế'),
    ('Tạp chí', 'Tạp chí Tâm lý học, Số 3 (312), tháng 3/2025, tr. 55\u201365'),
    ('Cỡ mẫu', '685 HS THPT (233 nam, 452 nữ), tuổi 15\u201319 (TB = 16,09)'),
    ('Thời điểm', 'Tháng 10/2021 (làn sóng 4 COVID-19)'),
    ('Công cụ', 'DASS-21 (lo âu + trầm cảm) + LOT-R (sự lạc quan)'),
    ('Phân tích', 'PROCESS macro v4.2 (mô hình 4) — phân tích trung gian'),
])

add_heading(d1, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d1, 'Điểm nổi bật:', bold=True)
add_p(d1, '\u2022 Lo âu 65,8%, trầm cảm 45,1% — mức cao, COVID-19 làn sóng 4.')
add_p(d1, '\u2022 Phân tích trung gian: sự lạc quan \u2192 giảm lo âu \u2192 giảm trầm cảm (vai trò trung gian của lo âu).')
add_p(d1, '\u2022 Sự lạc quan có tác động trực tiếp (\u03b2 = \u20130,21) VÀ gián tiếp qua lo âu (\u03b2 = \u20130,24) lên trầm cảm.')
add_p(d1, '\u2022 Tạp chí Tâm lý học — uy tín trong ngành tại Việt Nam.')
add_p(d1, '\u2022 PROCESS macro — phương pháp nâng cao, ít NC VN sử dụng.')

add_p(d1, 'Hạn chế:', bold=True)
add_p(d1, '\u2022 Lấy mẫu thuận tiện, khảo sát trực tuyến — thiên lệch tự chọn.')
add_p(d1, '\u2022 Nữ chiếm 67,4% mẫu (452/685) — thiên lệch giới.')
add_p(d1, '\u2022 Cắt ngang — không suy luận nhân quả (lạc quan \u2192 giảm lo âu hay ngược lại?).')
add_p(d1, '\u2022 Chỉ DASS-21 sàng lọc, không chẩn đoán. Tạp chí không lập chỉ mục quốc tế.')

add_p(d1, 'Nói cách khác, nghiên cứu này không chỉ đo mức độ lo âu/trầm cảm mà còn khám phá CƠ CHẾ — sự lạc quan giúp giảm trầm cảm thông qua việc giảm lo âu trước (trung gian một phần).', bold=True)

add_page_ref(d1, '55\u201365', 'Tạp chí Tâm lý học', 'Số 3 (312), 3/2025')

add_heading(d1, 'KẾT QUẢ CHÍNH', 2)
add_p(d1, 'Tỷ lệ lo âu và trầm cảm (N = 685, DASS-21):', bold=True)
add_table(d1,
    ['Mức độ', 'Lo âu n (%)', 'Trầm cảm n (%)'],
    [['Bình thường', '234 (34,2%)', '376 (54,9%)'],
     ['Nhẹ', '80 (11,7%)', '107 (15,6%)'],
     ['Vừa phải', '213 (31,1%)', '129 (18,8%)'],
     ['Nghiêm trọng', '69 (10,1%)', '45 (6,6%)'],
     ['Rất nghiêm trọng', '89 (13,0%)', '28 (4,1%)'],
     ['Tổng có triệu chứng', '451 (65,8%)', '309 (45,1%)']],
    widths=[3.5, 3.5, 3.5])
d1.add_paragraph()

add_p(d1, 'Tương quan và phân tích trung gian:', bold=True)
add_table(d1,
    ['Mối quan hệ', '\u03b2', 'SE', 'KTC 95%'],
    [['Lạc quan \u2192 Lo âu (trực tiếp)', '\u20130,15***', '0,08', '[\u20130,49; \u20130,17]'],
     ['Lạc quan \u2192 Trầm cảm (trực tiếp)', '\u20130,21***', '0,08', '[\u20130,69; \u20130,39]'],
     ['Lo âu \u2192 Trầm cảm', '0,59***', '0,03', '[0,65; 0,79]'],
     ['Lạc quan \u2192 Lo âu \u2192 Trầm cảm (gián tiếp)', '\u20130,24', '0,06', '[\u20130,36; \u20130,12]'],
     ['Tổng tác động', '\u20130,78***', '0,10', '[\u20130,97; \u20130,60]']],
    widths=[5.5, 1.5, 1.0, 3.0])
add_p(d1, '*** p < 0,001. Lo âu có vai trò TRUNG GIAN MỘT PHẦN trong mối quan hệ lạc quan \u2192 trầm cảm.', italic=True, size=10)
d1.add_paragraph()

add_p(d1, 'Tương quan Pearson:', bold=True)
add_table(d1,
    ['Biến', 'Lo âu', 'Lạc quan', 'Trầm cảm'],
    [['Lo âu (M=10,61; SD=7,10)', '1', '', ''],
     ['Lạc quan (M=14,87; SD=3,27)', '\u20130,154**', '1', ''],
     ['Trầm cảm (M=9,51; SD=8,57)', '0,625**', '\u20130,299**', '1']],
    widths=[5.0, 2.0, 2.0, 2.0])
d1.add_paragraph()

bh = d1.add_heading('Nhận định tổng hợp (Kiến trúc B)', 3)
for r in bh.runs: r.font.name = 'Times New Roman'; r.font.color.rgb = RGBColor(0, 0x70, 0xC0)
add_p(d1, 'Dữ liệu của nghiên cứu này, cho thấy gần 2/3 thanh thiếu niên Huế trải qua lo âu (65,8%) và gần 1/2 trầm cảm (45,1%) trong bối cảnh COVID-19, gợi ý rằng sự lạc quan không chỉ liên quan trực tiếp đến giảm trầm cảm mà còn tác động gián tiếp thông qua việc giảm lo âu — cung cấp cơ sở cho can thiệp tăng cường tư duy lạc quan ở thanh thiếu niên.', bold=True)

add_heading(d1, 'QUAN ĐIỂM PHẢN BIỆN / CRITICAL REVIEW', 2)
add_red(d1, 'Điểm mạnh:', bold=True)
add_red(d1, '1. Phân tích trung gian (mediation) bằng PROCESS macro — phương pháp nâng cao, ít NC VN sử dụng.')
add_red(d1, '2. Khám phá CƠ CHẾ — không chỉ đo tỷ lệ mà giải thích con đường lạc quan \u2192 lo âu \u2192 trầm cảm.')
add_red(d1, '3. Cỡ mẫu khá (685 HS), 3 khối lớp (10-12), tại Huế — miền Trung ít NC.')
add_red(d1, '4. DASS-21 tiếng Việt đã xác thực (Cronbach alpha tốt). LOT-R alpha = 0,82.')

add_red(d1, 'Hạn chế:', bold=True)
add_red(d1, '1. Lấy mẫu thuận tiện, khảo sát trực tuyến (10/2021 — COVID) — thiên lệch tự chọn.')
add_red(d1, '2. Nữ chiếm 67,4% — thiên lệch giới, không phân tích giới tính riêng cho lo âu/trầm cảm.')
add_red(d1, '3. Cắt ngang — mô hình trung gian GIẢ ĐỊNH chiều nhân quả nhưng không thể xác nhận.')
add_red(d1, '4. DASS-21 sàng lọc, ngưỡng lo âu ≥8 bao gồm mức nhẹ. So V-NAMHS (DISC-5: 2,3%), chênh lệch 29 lần.')
add_red(d1, '5. Không đánh giá yếu tố nguy cơ cụ thể (gia đình, học tập, mạng xã hội).')

add_red(d1, 'Đối chiếu giới tính (KT3):', bold=True)
add_red(d1, 'Nghiên cứu KHÔNG phân tích giới tính riêng cho lo âu/trầm cảm — đây là khoảng trống quan trọng khi mẫu có 67,4% nữ. Y văn xác nhận nữ > nam (McLean, 2011). Phát hiện tỷ lệ cao (65,8% lo âu) có thể phần nào do mẫu thiên nữ.')

add_red(d1, 'Gap:', bold=True)
add_red(d1, '1. Nghiên cứu dọc xác nhận chiều nhân quả: lạc quan \u2192 giảm lo âu/trầm cảm hay ngược lại.')
add_red(d1, '2. Can thiệp tăng cường tư duy lạc quan (positive psychology) tại trường THPT Huế.')
add_red(d1, '3. So sánh vùng miền: Huế vs Hà Nội (Hoa 2024: 40,6%) vs TPHCM (Vĩnh Lộc 2024: 25,1%).')
add_red(d1, '4. Phân tích giới tính chi tiết. Bổ sung yếu tố gia đình, học tập, mạng xã hội.')

add_abbreviation_table(d1, [
    ('DASS-21', 'Depression Anxiety Stress Scale 21 — Thang đo Trầm cảm Lo âu Căng thẳng'),
    ('LOT-R', 'Life Orientation Test-Revised — Trắc nghiệm Định hướng Cuộc sống (đo sự lạc quan)'),
    ('PROCESS', 'Phần mềm phân tích trung gian/điều tiết (Hayes, 2013)'),
    ('\u03b2', 'Beta — Hệ số hồi quy chuẩn hóa'),
])

d1.save('19_TranThaoVi_2025_TLH.docx')
sys.stderr.write('19_TranThaoVi OK\n')

# ============================================================
# BÀI 20: TRẦN HỒ VĨNH LỘC — TPHCM
# ============================================================
d2 = create_doc()
add_heading(d2, 'Trầm cảm, lo âu, căng thẳng và các yếu tố liên quan ở học sinh THPT tại Thành phố Hồ Chí Minh', 1)
add_heading(d2, 'Depression, Anxiety, Stress and Associated Factors among High School Students in Ho Chi Minh City', 2)

add_heading(d2, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d2, [
    ('Tiêu đề', 'Trầm cảm, lo âu, căng thẳng và các yếu tố liên quan ở HS THPT tại TPHCM'),
    ('Tác giả', 'Trần Hồ Vĩnh Lộc, Huỳnh Ngọc Vân Anh, Tô Gia Kiên'),
    ('Cơ quan', 'Khoa Y tế Công cộng, ĐH Y Dược TPHCM'),
    ('Tạp chí', 'Tạp chí Y học TPHCM, Tập 27, Số 5, 2024, tr. 100\u2013110'),
    ('DOI', '10.32895/hcjm.m.2024.05.12'),
    ('Cỡ mẫu', '976 HS THPT (53,4% nữ), 2 trường tại TPHCM'),
    ('Thời điểm', 'Tháng 2\u20134/2024'),
    ('Công cụ', 'DASS-Y (DASS dành cho thanh thiếu niên) + ESSA (áp lực học tập)'),
])

add_heading(d2, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d2, 'Điểm nổi bật:', bold=True)
add_p(d2, '\u2022 DASS-Y — phiên bản DÀNH RIÊNG cho thanh thiếu niên (không phải DASS-21 người lớn).')
add_p(d2, '\u2022 n = 976 — cỡ mẫu lớn nhất trong các NC TPHCM về chủ đề này.')
add_p(d2, '\u2022 Trầm cảm 31,7%, lo âu 25,1%, căng thẳng 23,8% — thấp hơn đáng kể so với các NC dùng DASS-21.')
add_p(d2, '\u2022 Phân tích yếu tố liên quan: giới tính, áp lực học tập, hôn nhân cha mẹ, kinh tế gia đình, học vấn mẹ.')
add_p(d2, '\u2022 ESSA — thang đo áp lực học tập chuẩn hóa, ít NC VN sử dụng.')
add_p(d2, '\u2022 ĐH Y Dược TPHCM — cơ sở uy tín, TC Y học TPHCM có DOI.')

add_p(d2, 'Hạn chế:', bold=True)
add_p(d2, '\u2022 Chỉ 2 trường TPHCM — không đại diện toàn thành phố.')
add_p(d2, '\u2022 Chọn lớp thuận tiện (4 lớp/khối) — không hoàn toàn ngẫu nhiên.')
add_p(d2, '\u2022 Pilot test chỉ 34 HS, 88,2% nữ — thiên lệch giới trong pilot.')

add_p(d2, 'Nói cách khác, đây là một trong số ít nghiên cứu tại Việt Nam sử dụng DASS-Y (phiên bản thanh thiếu niên, ngưỡng cắt khác DASS-21 người lớn) kết hợp đo lường áp lực học tập bằng thang ESSA — cho kết quả tỷ lệ thấp hơn đáng kể so với các NC dùng DASS-21.', bold=True)

add_page_ref(d2, '100\u2013110', 'TC Y học TPHCM', 'Tập 27, Số 5, 2024')

add_heading(d2, 'KẾT QUẢ CHÍNH', 2)
add_p(d2, 'Tỷ lệ trầm cảm, lo âu, căng thẳng (N = 976, DASS-Y):', bold=True)
add_table(d2,
    ['Tình trạng', 'Tỷ lệ', 'Ghi chú'],
    [['Trầm cảm (≥7 điểm)', '31,7%', ''],
     ['Lo âu (≥6 điểm)', '25,1%', 'Thấp hơn DASS-21 (40\u201386%)'],
     ['Căng thẳng (≥12 điểm)', '23,8%', ''],
     ['Có ≥1 vấn đề', '42,4%', ''],
     ['Có cả 3 vấn đề', '13,2%', '']],
    widths=[4.0, 2.5, 5.0])
d2.add_paragraph()

add_p(d2, 'Yếu tố liên quan (hồi quy đa biến):', bold=True)
add_table(d2,
    ['Yếu tố', 'Liên quan với', 'Hướng'],
    [['Giới tính nữ', 'Trầm cảm + Lo âu + Căng thẳng', 'Nữ > Nam (p < 0,05)'],
     ['Áp lực học tập nặng (ESSA ≥59)', 'Cả 3', 'Tăng nguy cơ'],
     ['Cha mẹ ly hôn/ly thân', 'Trầm cảm', 'Tăng nguy cơ'],
     ['Kinh tế gia đình khó khăn', 'Trầm cảm + Căng thẳng', 'Tăng nguy cơ'],
     ['Trình độ học vấn mẹ thấp', 'Lo âu', 'Tăng nguy cơ']],
    widths=[4.5, 4.0, 3.0])
d2.add_paragraph()

bh2 = d2.add_heading('Đối chiếu giới tính (KT3)', 3)
for r in bh2.runs: r.font.name = 'Times New Roman'; r.font.color.rgb = RGBColor(0, 0x70, 0xC0)
add_p(d2, 'Y văn quốc tế xác nhận nữ > nam về lo âu (McLean, 2011). Nghiên cứu tại TPHCM phù hợp xu hướng — nữ có nguy cơ cao hơn cả 3 chỉ số (p < 0,05). Đặc biệt, tác giả khuyến nghị "các chương trình dành cho nữ sinh" — gợi ý can thiệp nhạy cảm giới.')

bh3 = d2.add_heading('Nhận định tổng hợp (Kiến trúc B)', 3)
for r in bh3.runs: r.font.name = 'Times New Roman'; r.font.color.rgb = RGBColor(0, 0x70, 0xC0)
add_p(d2, 'Dữ liệu của nghiên cứu này, cho thấy tỷ lệ lo âu 25,1% ở 976 HS THPT TPHCM — thấp hơn đáng kể so với các NC dùng DASS-21 (Hoa 2024: 40,6%; Bảo Quyên 2025: 86,2%), gợi ý rằng DASS-Y với ngưỡng cắt riêng cho thanh thiếu niên có thể cho kết quả bảo thủ hơn và phù hợp hơn cho sàng lọc dân số trẻ. Áp lực học tập (ESSA) và cấu trúc gia đình là hai nhóm yếu tố mạnh nhất — phù hợp với Wen 2020 (áp lực OR=11,6) và Hoàng Trung Học 2025 (gia đình Beta=0,272).', bold=True)

add_heading(d2, 'QUAN ĐIỂM PHẢN BIỆN / CRITICAL REVIEW', 2)
add_red(d2, 'Điểm mạnh:', bold=True)
add_red(d2, '1. DASS-Y — phiên bản DÀNH RIÊNG cho VTN, ngưỡng cắt phù hợp hơn DASS-21 người lớn.')
add_red(d2, '2. n = 976, lấy mẫu ngẫu nhiên nhiều bậc — phương pháp tốt hơn đa số NC VN.')
add_red(d2, '3. ESSA đo áp lực học tập chuẩn hóa (Cronbach alpha = 0,81) — ít NC VN dùng.')
add_red(d2, '4. Phân tích hồi quy đa biến — xác định yếu tố độc lập.')
add_red(d2, '5. ĐH Y Dược TPHCM — cơ sở đào tạo hàng đầu, có DOI, pilot test.')

add_red(d2, 'Hạn chế:', bold=True)
add_red(d2, '1. Chỉ 2 trường — không đại diện TPHCM (239.501 HS THPT).')
add_red(d2, '2. Chọn lớp thuận tiện 4/khối — mâu thuẫn với "ngẫu nhiên nhiều bậc" (chỉ bậc trường ngẫu nhiên).')
add_red(d2, '3. DASS-Y chưa phổ biến quốc tế — khó so sánh với NC dùng DASS-21/GAD-7.')
add_red(d2, '4. Không đo giấc ngủ, mạng xã hội — 2 yếu tố mạnh nhất theo y văn (Chen 2023: giấc ngủ OR=6,99).')
add_red(d2, '5. Tạp chí Y học TPHCM — không lập chỉ mục PubMed/Scopus.')

add_red(d2, 'Gap:', bold=True)
add_red(d2, '1. So sánh DASS-Y vs DASS-21 vs GAD-7 trên cùng mẫu VTN — xác định ngưỡng phù hợp.')
add_red(d2, '2. Thêm biến giấc ngủ, mạng xã hội vào mô hình.')
add_red(d2, '3. Mở rộng nhiều trường TPHCM (công + tư + chuyên).')
add_red(d2, '4. Can thiệp giảm áp lực học tập + hỗ trợ gia đình ly hôn tại trường.')

add_abbreviation_table(d2, [
    ('DASS-Y', 'Depression Anxiety Stress Scales for Youth — Thang đo TCC dành cho thanh thiếu niên'),
    ('ESSA', 'Educational Stress Scale for Adolescents — Thang đo Áp lực Học tập'),
    ('TPHCM', 'Thành phố Hồ Chí Minh'),
    ('THPT', 'Trung học phổ thông'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('TCC', 'Trầm cảm, căng thẳng (Depression, Stress)'),
])

d2.save('20_TranHoVinhLoc_2024_YHTPHCM.docx')
sys.stderr.write('20_TranHoVinhLoc OK\n')
