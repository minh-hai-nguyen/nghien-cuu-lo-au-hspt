# -*- coding: utf-8 -*-
"""Nhận xét + phản biện cho 4 bài VN còn lại"""
import sys, io, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

# ============================================================
# BÀI 1: NGUYỄN DANH LÂM 2024 — Yên Định, Thanh Hóa
# ============================================================
d1 = create_doc()
add_heading(d1, 'Thực trạng nguy cơ stress, lo âu, trầm cảm của học sinh THPT huyện Yên Định, Thanh Hóa', 1)
add_heading(d1, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d1, [
    ('Tiêu đề', 'Thực trạng nguy cơ stress, lo âu, trầm cảm của HS THPT huyện Yên Định, Thanh Hóa'),
    ('Tác giả', 'Nguyễn Danh Lâm, Lê Minh Giang, Nguyễn Thị Phương Mai, Nguyễn Thị Diệu Thúy, Nguyễn Thị Thanh Mai'),
    ('Cơ quan', 'BV huyện Yên Định; ĐH Y Hà Nội; ĐH Y Dược ĐH Thái Nguyên'),
    ('Tạp chí', 'Tạp chí Y học Việt Nam, Tập 516, Số 1, tháng 7/2022, tr. 67\u201370'),
    ('Cỡ mẫu', '482 HS THPT (2 trường), 9/2021, DASS-21'),
])

add_heading(d1, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d1, 'Điểm nổi bật:', bold=True)
add_p(d1, '\u2022 Vùng bán đô thị Thanh Hóa — ít NC tại khu vực này.')
add_p(d1, '\u2022 Stress 41,7%, lo âu 49,0%, trầm cảm 43,6% — mức trung bình-cao.')
add_p(d1, '\u2022 Dữ liệu tự hại: 10% đã tự làm đau, 1,4% đã cố tự tử — đáng báo động.')
add_p(d1, '\u2022 25% đã từng nghĩ đến tự tử — cao hơn nhiều NC khác.')

add_p(d1, 'Hạn chế:', bold=True)
add_p(d1, '\u2022 Chỉ 2 trường (chọn mẫu có chủ đích), không ngẫu nhiên.')
add_p(d1, '\u2022 Chỉ thống kê mô tả, KHÔNG có phân tích yếu tố liên quan (hồi quy).')
add_p(d1, '\u2022 Tạp chí Y học VN — không PubMed/Scopus. Bài chỉ 4 trang — thiếu chi tiết.')
add_p(d1, '\u2022 Không phân tích giới tính, khối lớp chi tiết.')

add_heading(d1, 'KẾT QUẢ CHÍNH', 2)
add_page_ref(d1, '67\u201370', 'TC Y học Việt Nam', 'Tập 516, 7/2022')
add_table(d1,
    ['Tình trạng', 'Tỷ lệ', 'Nhẹ', 'Vừa', 'Nặng', 'Rất nặng'],
    [['Stress', '41,7%', '16,0%', '14,1%', '6,8%', '4,8%'],
     ['Lo âu', '49,0%', '7,7%', '24,5%', '8,1%', '4,6%'],
     ['Trầm cảm', '43,6%', '12,2%', '27,2%', '2,9%', '1,2%']],
    widths=[2.5, 1.5, 1.5, 1.5, 1.5, 2.0])
d1.add_paragraph()
add_table(d1,
    ['Tự hại / Tự tử', 'Chưa bao giờ', 'Đã nghĩ đến', 'Đã thực hiện'],
    [['Tự làm đau', '58,4%', '31,6%', '10,0%'],
     ['Tự tử', '75,0%', '23,6%', '1,4%']],
    widths=[3.5, 3.0, 3.0, 3.0])

add_red_heading(d1, 'PHẢN BIỆN CHI TIẾT')
add_red(d1, 'Điểm mạnh:', bold=True)
add_red(d1, '1. Vùng bán đô thị Thanh Hóa — lấp khoảng trống giữa đô thị và nông thôn.')
add_red(d1, '2. Dữ liệu tự hại/tự tử quan trọng — 10% đã tự làm đau, 1,4% cố tự tử.')
add_red(d1, '3. DASS-21 tiếng Việt đã được xác thực (Thach et al., 2013).')

add_red(d1, 'Hạn chế:', bold=True)
add_red(d1, '1. Chọn mẫu có chủ đích (2 trường) — không ngẫu nhiên, thiên lệch.')
add_red(d1, '2. Chỉ thống kê mô tả — KHÔNG phân tích yếu tố liên quan.')
add_red(d1, '3. Bài chỉ 4 trang — thiếu phần thảo luận sâu, không so sánh giới tính.')
add_red(d1, '4. Tạp chí không lập chỉ mục quốc tế. Năm xuất bản 2022 nhưng tiêu đề ghi 2024.')

add_red(d1, 'Gap:', bold=True)
add_red(d1, '1. Phân tích yếu tố nguy cơ (hồi quy). 2. So sánh giới tính, khối lớp. 3. Can thiệp giảm tự hại.')

d1.save('17_NguyenDanhLam_2022_YHVN.docx')
sys.stderr.write('17_NguyenDanhLam OK\n')

# ============================================================
# BÀI 2: AN GIANG 2024 — Long Bình
# ============================================================
d2 = create_doc()
add_heading(d2, 'Kết quả sàng lọc lo âu, trầm cảm, stress bằng DASS-21 ở học sinh phổ thông Long Bình, tỉnh An Giang năm 2024', 1)
add_heading(d2, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d2, [
    ('Tiêu đề', 'Kết quả sàng lọc lo âu, trầm cảm, stress bằng DASS-21 ở HS phổ thông Long Bình, An Giang 2024'),
    ('Tác giả', 'Lê Minh T., Nguyễn Đăng K., Ngô Anh V.'),
    ('Tạp chí', 'Tạp chí Y học Việt Nam, Tập 549, Số 1, tháng 4/2025, tr. 32\u201335'),
    ('DOI', '10.51298/vmj.v549i1.13506'),
    ('Cỡ mẫu', '366 HS THPT Long Bình, An Giang, 6/2024, DASS-21'),
])

add_heading(d2, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d2, 'Điểm nổi bật:', bold=True)
add_p(d2, '\u2022 Vùng ĐBSCL — ít NC SKTT tại khu vực này.')
add_p(d2, '\u2022 Lo âu 61,2%, trầm cảm 47,3%, stress 38,0%.')
add_p(d2, '\u2022 27,3% có đồng thời cả 3 triệu chứng.')
add_p(d2, '\u2022 Tỷ lệ nam/nữ cân bằng (50,5/49,5%).')

add_p(d2, 'Hạn chế:', bold=True)
add_p(d2, '\u2022 Chỉ 1 trường (Long Bình) — không đại diện.')
add_p(d2, '\u2022 Chỉ thống kê mô tả, KHÔNG phân tích yếu tố liên quan.')
add_p(d2, '\u2022 Bài ngắn (4 trang nội dung), thiếu thảo luận sâu.')

add_heading(d2, 'KẾT QUẢ CHÍNH', 2)
add_page_ref(d2, '32\u201335', 'TC Y học Việt Nam', 'Tập 549, 4/2025')
add_table(d2,
    ['Tình trạng', 'Tỷ lệ', 'Nhẹ', 'Vừa', 'Nặng', 'Rất nặng'],
    [['Lo âu', '61,2%', '9,3%', '24,0%', '12,6%', '15,3%'],
     ['Trầm cảm', '47,3%', '15,8%', '18,0%', '8,5%', '4,9%'],
     ['Stress', '38,0%', '12,8%', '16,4%', '11,5%', '4,9%']],
    widths=[2.5, 1.5, 1.5, 1.5, 1.5, 2.0])
d2.add_paragraph()
add_p(d2, 'Đồng mắc: 28,1% không có dấu hiệu, 24,6% có 1, 20,0% có 2, 27,3% có cả 3 triệu chứng.', bold=True)

add_red_heading(d2, 'PHẢN BIỆN CHI TIẾT')
add_red(d2, 'Điểm mạnh:', bold=True)
add_red(d2, '1. Vùng ĐBSCL — lấp khoảng trống địa lý. 2. Tỷ lệ giới cân bằng. 3. Phân tích đồng mắc.')
add_red(d2, 'Hạn chế:', bold=True)
add_red(d2, '1. 1 trường duy nhất. 2. Không phân tích yếu tố liên quan. 3. Bài ngắn. 4. Không lập chỉ mục QT.')
add_red(d2, 'Gap:', bold=True)
add_red(d2, '1. Mở rộng nhiều trường An Giang. 2. Phân tích hồi quy. 3. So sánh ĐBSCL với đô thị.')

d2.save('18_AnGiang_2025_YHVN.docx')
sys.stderr.write('18_AnGiang OK\n')

# ============================================================
# BÀI 3: TRẦN THẢO VI — Huế
# ============================================================
d3 = create_doc()
add_heading(d3, 'Lo âu và trầm cảm ở thanh thiếu niên tại Thừa Thiên Huế', 1)
add_heading(d3, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d3, [
    ('Tiêu đề', 'Lo âu và trầm cảm ở thanh thiếu niên tại Thừa Thiên Huế'),
    ('Tạp chí', 'Tạp chí Tâm lý học, Số 3 (312), 3/2025, tr. 55+'),
    ('Số trang', '11 trang'),
    ('Ghi chú', 'Bài tiếng Việt — cần đọc PDF đầy đủ để nhận xét chi tiết'),
])

add_heading(d3, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d3, '\u2022 Tạp chí Tâm lý học — tạp chí chuyên ngành uy tín tại Việt Nam.')
add_p(d3, '\u2022 Nghiên cứu tại Thừa Thiên Huế — vùng miền Trung, ít NC.')
add_p(d3, '\u2022 11 trang — bài dài, có thể chi tiết hơn các bài VN khác.')
add_p(d3, '\u2022 Cần đọc chi tiết PDF để nhận xét số liệu cụ thể.', bold=True)

add_red_heading(d3, 'GHI CHÚ')
add_red(d3, 'Bài viết bằng tiếng Việt. Cần đọc toàn bộ PDF để trích xuất số liệu và viết phản biện chi tiết. Hiện tại chỉ xác nhận: 11 trang, Tạp chí Tâm lý học, Số 3/2025, chủ đề lo âu và trầm cảm VTN tại Huế.')

d3.save('19_TranThaoVi_2025_TLH.docx')
sys.stderr.write('19_TranThaoVi OK\n')

# ============================================================
# BÀI 4: TRẦN HỒ VĨNH LỘC — TPHCM
# ============================================================
d4 = create_doc()
add_heading(d4, 'Trầm cảm, lo âu, căng thẳng và các yếu tố liên quan của học sinh THPT tại TP.HCM', 1)
add_heading(d4, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(d4, [
    ('Tiêu đề', 'Trầm cảm, lo âu, căng thẳng và các yếu tố liên quan ở HS THPT tại TPHCM'),
    ('Tạp chí', 'Tạp chí Y học TP.HCM, Vol. 27, No. 5, 2024'),
    ('Số trang', '11 trang'),
    ('Ghi chú', 'Bài tiếng Việt — cần đọc PDF đầy đủ'),
])

add_heading(d4, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(d4, '\u2022 Tạp chí Y học TPHCM — tạp chí y khoa uy tín miền Nam.')
add_p(d4, '\u2022 Nghiên cứu tại TPHCM — thành phố lớn nhất VN.')
add_p(d4, '\u2022 11 trang — bài dài, có phân tích yếu tố liên quan.')
add_p(d4, '\u2022 Cần đọc chi tiết PDF để nhận xét.', bold=True)

add_red_heading(d4, 'GHI CHÚ')
add_red(d4, 'Bài viết bằng tiếng Việt. Có 2 file có thể liên quan:\n- VN_TPHCM_DASS_THPT_2024.pdf (11 trang, TC Y học TPHCM)\n- VN_Multicenter_2631HS_TPHCM_2025.pdf (11 trang, Springer)\nCần xác định đúng file và đọc chi tiết.')

d4.save('20_TranHoVinhLoc_2024_YHTPHCM.docx')
sys.stderr.write('20_TranHoVinhLoc OK\n')
