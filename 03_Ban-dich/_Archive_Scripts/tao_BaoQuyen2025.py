# -*- coding: utf-8 -*-
"""Bảo Quyên 2025 — song ngữ, DASS-21, 501 HS Hà Nội, lo âu 86,2%"""
import sys, io, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()
add_link_and_qr(doc, 'https://doi.org/10.52163/yhc.v66iCD10.2617', 'QR_BaoQuyen2025.png')

add_heading(doc, 'Thực trạng sức khỏe tâm thần của học sinh THPT tại Hà Nội: Nghiên cứu cắt ngang về trầm cảm, lo âu và căng thẳng', 1)
add_heading(doc, 'Mental Health Status of High School Students in Hanoi: A Cross-Sectional Study on Depression, Anxiety, and Stress', 2)

add_heading(doc, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(doc, [
    ('Tiêu đề', 'Thực trạng sức khỏe tâm thần của học sinh THPT tại Hà Nội'),
    ('Tác giả', 'Nguyễn Ngọc Bảo Quyên, Lưu Minh Đức, Vũ Minh Thy, Nguyễn Yến Minh, Nguyễn Châu Anh, Hoàng Lan Vân, Giáp Thị Thanh Tình'),
    ('Cơ quan', 'Đại học VinUni; ĐH Y Dược, ĐHQG Hà Nội'),
    ('Tạp chí', 'Tạp chí Y học Cộng đồng, Tập 66, Chuyên đề 10, 2025, tr. 79\u201386'),
    ('DOI', '10.52163/yhc.v66iCD10.2617'),
    ('Cỡ mẫu', '501 HS THPT Hà Nội (78% nữ), 8\u201312/2023'),
    ('Công cụ', 'DASS-21, khảo sát trực tuyến, SPSS 27.0'),
])

add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(doc, 'Điểm nổi bật:', bold=True)
add_p(doc, '\u2022 Lo âu 86,2%, trầm cảm 78,8%, căng thẳng 76,6% — tỷ lệ CAO NHẤT trong tất cả nghiên cứu VN.')
add_p(doc, '\u2022 Nữ bị ảnh hưởng nặng hơn nam ở cả 3 chỉ số (p < 0,05).')
add_p(doc, '\u2022 Nghiên cứu tại Hà Nội, VinUni — cơ sở đào tạo uy tín.')

add_p(doc, 'Hạn chế:', bold=True)
add_p(doc, '\u2022 Tỷ lệ CỰC CAO (86,2% lo âu) — DASS-21 với ngưỡng thấp (≥4 lo âu) bao gồm cả "nhẹ", đặt câu hỏi về tính phân biệt lâm sàng.')
add_p(doc, '\u2022 Nữ chiếm 78% mẫu — thiên lệch giới tính nghiêm trọng.')
add_p(doc, '\u2022 Khảo sát trực tuyến — thiên lệch tự chọn. Cỡ mẫu 501 — nhỏ cho Hà Nội.')
add_p(doc, '\u2022 So với V-NAMHS 2022 (DISC-5: 2,3%), chênh lệch 37 lần — phương pháp quyết định kết quả.')

add_heading(doc, 'TÓM TẮT', 2)
add_page_ref(doc, '79\u201380', 'TC Y học Cộng đồng', 'Tập 66, CD10, 2025')
add_red(doc, 'Phương pháp: Cắt ngang, 501 HS THPT Hà Nội (8\u201312/2023). DASS-21 trực tuyến. SPSS 27.0: ANOVA, t-test, chi-square, hồi quy.')
add_red(doc, 'Kết quả: Trầm cảm 78,8% (nặng+rất nặng 33,2%), lo âu 86,2% (nặng+rất nặng 54,5%), căng thẳng 76,6% (nặng+rất nặng 38,3%). Giới tính liên quan có ý nghĩa (p < 0,05). Khối lớp và loại trường không ý nghĩa.')

add_heading(doc, 'KẾT QUẢ CHI TIẾT', 2)
add_page_ref(doc, '82\u201385', 'TC Y học Cộng đồng', 'Tập 66, CD10, 2025')

add_p(doc, 'Tỷ lệ trầm cảm, lo âu, căng thẳng (N = 501):', bold=True)
add_table(doc,
    ['Mức độ', 'Trầm cảm n (%)', 'Lo âu n (%)', 'Căng thẳng n (%)'],
    [['Bình thường', '106 (21,2%)', '69 (13,8%)', '117 (23,4%)'],
     ['Nhẹ', '78 (15,6%)', '33 (6,6%)', '72 (14,4%)'],
     ['Vừa', '151 (30,1%)', '126 (25,1%)', '120 (24,0%)'],
     ['Nặng', '61 (12,2%)', '87 (17,4%)', '126 (25,1%)'],
     ['Rất nặng', '105 (21,0%)', '186 (37,1%)', '66 (13,2%)'],
     ['Tổng có triệu chứng', '395 (78,8%)', '432 (86,2%)', '384 (76,6%)'],
     ['Điểm TB \u00b1 SD', '17,71 \u00b1 10,14', '16,58 \u00b1 8,45', '22,14 \u00b1 9,04']],
    widths=[3.0, 3.0, 3.0, 3.0])
doc.add_paragraph()

add_p(doc, 'Nhân khẩu học: 78% nữ, 21,6% nam. Công lập 68%, tư thục 19%, chuyên 13%. 90,4% không dùng thuốc, 92,8% không dùng chất kích thích.', italic=True)

add_heading(doc, 'THẢO LUẬN', 2)
add_p(doc, 'Tỷ lệ lo âu 86,2% là CAO NHẤT trong tất cả NC tại Việt Nam. So sánh: Hoa 2024 (GAD-7): 40,6%; Ngo Anh Vinh 2024 (DASS-21): 54,4%; V-NAMHS 2022 (DISC-5): 2,3%. Sự chênh lệch chủ yếu do ngưỡng cắt DASS-21 rất thấp (≥4 cho lo âu) — bao gồm cả mức "nhẹ" mà có thể không có ý nghĩa lâm sàng.')
add_p(doc, 'Nữ chiếm 78% mẫu — thiên lệch giới tính có thể phóng đại tỷ lệ tổng (nữ luôn có tỷ lệ cao hơn nam). Khảo sát trực tuyến cũng có thiên lệch tự chọn.')

add_heading(doc, 'TÀI LIỆU THAM KHẢO', 2)
add_p(doc, '[Giữ nguyên TLTK từ bài gốc — xem PDF trang 86]', italic=True)

add_abbreviation_table(doc, [
    ('DASS-21', 'Depression Anxiety Stress Scale 21 — Thang đo Trầm cảm Lo âu Căng thẳng'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('THPT', 'Trung học phổ thông'),
    ('YHCĐ', 'Tạp chí Y học Cộng đồng'),
])

add_red_heading(doc, 'PHẢN BIỆN CHI TIẾT')
add_red(doc, 'Điểm mạnh:', bold=True)
add_red(doc, '1. Đề tài thời sự — SKTT HS THPT Hà Nội, dữ liệu 2023.')
add_red(doc, '2. VinUni — cơ sở đào tạo uy tín, có IRB từ Vinmec.')
add_red(doc, '3. Đánh giá cả 3 chỉ số cùng lúc (trầm cảm, lo âu, căng thẳng).')

add_red(doc, 'Hạn chế:', bold=True)
add_red(doc, '1. Tỷ lệ lo âu 86,2% CỰC CAO — DASS-21 ngưỡng ≥4 bao gồm cả "nhẹ". Nếu chỉ tính nặng+rất nặng: 54,5% — vẫn rất cao.')
add_red(doc, '2. Nữ chiếm 78% mẫu — thiên lệch giới tính nghiêm trọng, không đại diện dân số HS.')
add_red(doc, '3. Cỡ mẫu 501 — nhỏ cho thành phố Hà Nội (so với Hoa 2024: 3.910).')
add_red(doc, '4. Khảo sát trực tuyến — thiên lệch tự chọn, HS có vấn đề SKTT có thể tham gia nhiều hơn.')
add_red(doc, '5. Tạp chí YHCĐ — không lập chỉ mục quốc tế (PubMed/Scopus).')
add_red(doc, '6. Chỉ dùng thống kê đơn biến (ANOVA, t-test, chi-square) + hồi quy — chưa rõ loại hồi quy.')

add_red(doc, 'Gap:', bold=True)
add_red(doc, '1. Cần tái lặp với cỡ mẫu lớn hơn, tỷ lệ giới cân bằng.')
add_red(doc, '2. So sánh DASS-21 với GAD-7 hoặc DISC-5 trên cùng mẫu.')
add_red(doc, '3. Nghiên cứu can thiệp tại trường THPT Hà Nội.')

fname = '16_BaoQuyen_2025_YHCD.docx'
doc.save(fname)
sys.stderr.write(f'{fname} saved OK\n')
