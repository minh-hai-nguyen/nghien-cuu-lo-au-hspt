# -*- coding: utf-8 -*-
"""Tóm tắt CTH v5 cho Saikia 2023 & Wen 2020 — 2 bài nam>nữ / giới tính đặc biệt"""
import sys, io, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

# ============================================================
# BÀI 1: SAIKIA et al. 2023
# ============================================================
doc = create_doc()

add_heading(doc, 'TÓM TẮT BÀI: Saikia et al. 2023', 1)
add_p(doc, 'Phát hiện đặc biệt: NAM > NỮ về lo âu — trái ngược y văn quốc tế', bold=True)

# ĐỊNH DANH
add_red_heading(doc, 'Tên công trình, tác giả, năm, mẫu khảo sát, đối tượng, khách thể, địa bàn')
add_p(doc, 'Công trình « Các bệnh lý sức khỏe tâm thần và các yếu tố liên quan ở thanh thiếu niên tại Kamrup (Metro), Assam: Một nghiên cứu tại trường học » (Mental Health Morbidities and their Correlates among the Adolescents in Kamrup (Metro), Assam: A School-Based Study), do Anku M. Saikia, Hemen Das và Vinoth Rajendran (2023), Khoa Y học Cộng đồng, Trường Y khoa Gauhati, Guwahati, Assam, Ấn Độ, khảo sát 360 thanh thiếu niên 14–17 tuổi tại 10 trường THCS công lập quận Kamrup (Metro), vùng Đông Bắc Ấn Độ. Indian Journal of Community Medicine, 48(5), 733–738. PMID: 37970140.')

# PHƯƠNG PHÁP
add_red_heading(doc, 'Phương pháp nghiên cứu')
add_p(doc, 'Công trình sử dụng Thang đo Trầm cảm Lo âu Căng thẳng 21 mục (DASS-21 — Depression Anxiety Stress Scale) phiên bản tiếng Assam, kết hợp thang Kuppuswamy sửa đổi đánh giá kinh tế xã hội. Nói cách khác, đây là nghiên cứu sàng lọc cắt ngang tại trường học.', bold=True)
add_p(doc, 'DASS-21 là phiên bản rút gọn 21 mục của thang gốc 42 mục (Lovibond & Lovibond, 1995). Phiên bản tiếng Assam được dịch ngược và xác nhận bởi nhóm chuyên gia tâm thần học. Lấy mẫu ngẫu nhiên đơn giản 10/120 trường, 36 HS/trường (12/lớp 8,9,10), nam nữ bằng nhau. Thu thập 4/2019–6/2020. Phân tích chi-square, SPSS.')

# KẾT QUẢ
add_red_heading(doc, 'Kết quả nghiên cứu định lượng')
add_table(doc,
    ['Tình trạng', 'Tỷ lệ', 'n/N'],
    [['Trầm cảm', '22,2%', '80/360'],
     ['Lo âu', '24,4%', '88/360'],
     ['Căng thẳng', '6,9%', '25/360']],
    widths=[5.0, 3.0, 3.0])
doc.add_paragraph()

add_table(doc,
    ['Yếu tố', 'P (lo âu)', 'Ghi chú'],
    [['Giới tính: NAM 30,0% > NỮ 18,9%', '0,049', '⚠️ TRÁI Y VĂN'],
     ['Cha mẹ đơn thân', '<0,001', '63,2% nếu đơn thân'],
     ['Cha mẹ sử dụng rượu', '<0,001', '38,3–40,4%'],
     ['Chơi trò chơi điện tử', '0,003', '27,9–30,8%'],
     ['Lưu ban', '<0,001', '51,7–58,6%'],
     ['Tầng lớp KT-XH thấp', '0,028', 'Chỉ ý nghĩa với lo âu']],
    widths=[5.5, 2.5, 4.0])
doc.add_paragraph()

# ĐỐI CHIẾU GIỚI TÍNH
h = doc.add_heading('Khác biệt giới tính — PHÁT HIỆN ĐẶC BIỆT', 3)
for r in h.runs: r.font.name = 'Times New Roman'; r.font.color.rgb = RGBColor(0, 0x70, 0xC0)
add_p(doc, 'Nam giới có tỷ lệ lo âu CAO HƠN nữ giới (30,0% so với 18,9%, P = 0,049) — TRÁI NGƯỢC hoàn toàn với y văn quốc tế, nơi nữ luôn cao hơn nam (McLean et al., 2011; Salk et al., 2017). Trong 11 bài nghiên cứu của Đề tài, chỉ có 2 bài cho thấy nam > nữ (Saikia 2023 và Xu 2021). Phát hiện này có thể được giải thích bởi đặc thù văn hóa bộ lạc vùng Đông Bắc Ấn Độ, nơi kỳ vọng về vai trò nam giới khác biệt so với phần còn lại của Ấn Độ. Tuy nhiên, nghiên cứu tại cùng khu vực (Kumar & Akoijam 2017 tại Manipur; Nag et al. 2019 tại Tripura) đều cho thấy nữ > nam — khiến phát hiện này càng bất thường.')

# NHẬN XÉT
add_red_heading(doc, 'Nhận xét, phát hiện qua kết quả nghiên cứu')
add_p(doc, '• Cha mẹ đơn thân: tỷ lệ lo âu 63,2% — gấp gần 3 lần nhóm sống với cả hai bố mẹ (22,3%). Đây là yếu tố nguy cơ gia đình mạnh, tương tự Pham et al. (2024) tại Việt Nam.')
add_p(doc, '• Chơi game: liên quan có ý nghĩa với cả lo âu (P=0,003) và trầm cảm (P<0,001), tương tự Chen et al. (2023) tại Trung Quốc (OR=5,00 cho rối loạn chơi game).')
add_p(doc, '• Đông Bắc Ấn Độ: khu vực hầu như chưa được nghiên cứu về SKTT thanh thiếu niên — nghiên cứu tiên phong.')

# KẾT LUẬN
add_red_heading(doc, 'Kết luận')
add_p(doc, 'Tỷ lệ đáng kể trầm cảm (22,2%), lo âu (24,4%) và căng thẳng (6,9%) ở thanh thiếu niên Đông Bắc Ấn Độ. Phát hiện nam > nữ về lo âu (P=0,049) gợi ý yếu tố văn hóa bộ lạc đặc thù có thể đảo ngược xu hướng giới tính. Các yếu tố gia đình (đơn thân, rượu) và lối sống (game) cần được giải quyết trong can thiệp.', bold=True)

# PHẢN BIỆN
add_red_heading(doc, 'Phản biện')
add_red(doc, 'Chỉ dùng chi-square đơn biến — thiếu hồi quy đa biến. DASS-21 tiếng Assam thiếu Cronbach alpha. Thu thập 4/2019–6/2020 trùng COVID-19 không kiểm soát. P=0,049 rất sát ngưỡng — có thể thay đổi với cỡ mẫu lớn hơn. Tuy nhiên, lấy mẫu ngẫu nhiên 10/120 trường là điểm mạnh.')

add_red_heading(doc, 'Hướng nghiên cứu tiếp theo')
add_red(doc, 'Mở rộng cỡ mẫu với hồi quy đa biến. Khảo sát sâu văn hóa bộ lạc giải thích nam > nữ. Xác thực DASS-21 tiếng Assam. Can thiệp giảm game và hỗ trợ gia đình đơn thân.')

p = doc.add_paragraph()
r = p.add_run('Đánh giá: ⭐⭐⭐ Trung bình. Tiên phong cho Đông Bắc Ấn Độ, lấy mẫu ngẫu nhiên tốt, nhưng thiếu đa biến và P sát ngưỡng.')
r.bold = True; r.font.name = 'Times New Roman'; r.font.size = Pt(12)

doc.save('TomTat_Saikia_2023.docx')
sys.stderr.write('TomTat_Saikia_2023.docx OK\n')

# ============================================================
# BÀI 2: WEN et al. 2020
# ============================================================
doc2 = create_doc()

add_heading(doc2, 'TÓM TẮT BÀI: Wen et al. 2020', 1)
add_p(doc2, 'Phát hiện đặc biệt: Nữ lo âu nặng gấp gần 4 lần nam — áp lực học tập OR=11,6', bold=True)

# ĐỊNH DANH
add_red_heading(doc2, 'Tên công trình, tác giả, năm, mẫu khảo sát, đối tượng, khách thể, địa bàn')
add_p(doc2, 'Công trình « Phân tích hồ sơ tiềm ẩn về lo âu ở học sinh trung học cơ sở tại các vùng nông thôn kém phát triển của Trung Quốc » (A Latent Profile Analysis of Anxiety among Junior High School Students in Less Developed Rural Regions of China), do Xiaotong Wen, Yixiang Lin, Yuchen Liu và cộng sự (2020), Đại học Nam Xương, tỉnh Giang Tây, khảo sát 900 học sinh lớp 9–12 (tuổi TB 14,14) tại 6 huyện nông thôn tỉnh Giang Tây, miền Đông Trung Quốc. Int J Environ Res Public Health, 17(11):4079. PMID: 32521646.')

# PHƯƠNG PHÁP
add_red_heading(doc2, 'Phương pháp nghiên cứu')
add_p(doc2, 'Công trình sử dụng Trắc nghiệm Sức khỏe Tâm thần (MHT — Mental Health Test, 100 mục, 8 thang nội dung, Cronbach α = 0,878) — thang chẩn đoán lo âu tiêu chuẩn hóa cho THCS Trung Quốc. Nói cách khác, nghiên cứu sử dụng phân tích hồ sơ tiềm ẩn (LPA — Latent Profile Analysis) kết hợp hồi quy logistic đa biến — phương pháp tập trung vào cá nhân (person-centered).', bold=True)
add_p(doc2, 'Lấy mẫu cụm ngẫu nhiên phân tầng nhiều giai đoạn: Ủy ban Y tế tỉnh chọn huyện → chọn xã → chọn trường → chọn lớp ngẫu nhiên. Phân tích bằng Mplus 7.4 (LPA) và SPSS 24.0 (hồi quy logistic).')

# KẾT QUẢ
add_red_heading(doc2, 'Kết quả nghiên cứu định lượng')
add_p(doc2, 'LPA xác định 3 hồ sơ lo âu (N = 900):', bold=True)
add_table(doc2,
    ['Hồ sơ', 'n', 'Tỷ lệ', 'Đặc điểm'],
    [['Lo âu nhẹ', '173', '19,2%', 'Điểm thấp nhất trên 8 yếu tố'],
     ['Lo âu trung bình', '504', '56,0%', 'Điểm trung bình'],
     ['Lo âu nặng', '223', '24,8%', 'Điểm cao nhất trên 8 yếu tố']],
    widths=[3.5, 1.5, 2.0, 5.0])
doc2.add_paragraph()

add_p(doc2, 'Yếu tố liên quan (hồi quy logistic, nhóm tham chiếu: lo âu nhẹ):', bold=True)
add_table(doc2,
    ['Yếu tố', 'Lo âu TB (OR)', 'Lo âu nặng (OR)', 'Ý nghĩa'],
    [['Áp lực học tập rất cao', '6,523*', '11,579*', 'Yếu tố mạnh nhất'],
     ['Áp lực học tập cao', '6,122*', '5,894*', '—'],
     ['Giới tính nam (vs nữ)', '0,649*', '0,262*', 'Nam = bảo vệ → NỮ > NAM'],
     ['SKTT trường đầy đủ', '—', '0,562*', 'Yếu tố bảo vệ'],
     ['Thành tích xuất sắc (vs không đạt)', '0,377*', '—', 'Yếu tố bảo vệ']],
    widths=[4.0, 2.5, 2.5, 3.5])
add_p(doc2, '* P < 0,05', italic=True, size=10)
doc2.add_paragraph()

# ĐỐI CHIẾU GIỚI TÍNH
h2 = doc2.add_heading('Khác biệt giới tính — NỮ > NAM (phù hợp y văn)', 3)
for r in h2.runs: r.font.name = 'Times New Roman'; r.font.color.rgb = RGBColor(0, 0x70, 0xC0)
add_p(doc2, 'OR nam = 0,262 cho lo âu nặng → nữ có nguy cơ lo âu nặng gấp gần 4 lần nam. Nữ chiếm 63,23% nhóm lo âu nặng so với nam 36,77% (χ² = 31,337, P < 0,001). Phát hiện này PHÙ HỢP với y văn quốc tế — nữ lo âu hơn nam.')
add_p(doc2, '⚠️ LƯU Ý: Bài này ban đầu bị nhầm là "nam > nữ". Sau khi so khớp với PDF gốc, xác nhận NỮ > NAM. OR = 0,262 cho nam nghĩa là nam có nguy cơ CHỈ BẰNG 26,2% so với nữ.', bold=True)

# NHẬN XÉT
add_red_heading(doc2, 'Nhận xét, phát hiện qua kết quả nghiên cứu')
add_p(doc2, '• Áp lực học tập — yếu tố nguy cơ mạnh nhất: OR = 11,579 cho lo âu nặng khi áp lực "rất cao". Đây là OR lớn nhất trong tất cả 11 bài NC, phản ánh đặc thù hệ thống giáo dục Trung Quốc tập trung vào thi cử.')
add_p(doc2, '• Hỗ trợ SKTT tại trường — yếu tố bảo vệ: OR = 0,562, gợi ý thiết lập phòng tư vấn tâm lý chuyên nghiệp tại trường giảm lo âu nặng.')
add_p(doc2, '• Phương pháp LPA độc đáo: xác định 3 hồ sơ khách quan bằng tiêu chuẩn thống kê (BIC, entropy = 0,816, BLRT P < 0,001) — ít NC nào sử dụng phương pháp này.')

# KẾT LUẬN
add_red_heading(doc2, 'Kết luận')
add_p(doc2, '24,8% học sinh THCS nông thôn Trung Quốc có lo âu nặng. Áp lực học tập là yếu tố nguy cơ mạnh nhất (OR = 11,6). Nữ lo âu nặng gấp gần 4 lần nam (OR nam = 0,262). Hỗ trợ SKTT tại trường giảm nguy cơ (OR = 0,562). Giảm áp lực học tập và thiết lập phòng tư vấn tâm lý chuyên nghiệp là hai khuyến nghị chính.', bold=True)

# PHẢN BIỆN
add_red_heading(doc2, 'Phản biện')
add_red(doc2, 'LPA độc đáo nhưng cỡ mẫu 900 khiêm tốn, chỉ nông thôn 1 tỉnh. MHT là thang đặc thù Trung Quốc — khó so sánh quốc tế (không dùng GAD-7/DASS-21). Tạp chí IJERPH bị loại khỏi SCIE năm 2023. Thiết kế cắt ngang — không suy luận nhân quả.')

add_red_heading(doc2, 'Hướng nghiên cứu tiếp theo')
add_red(doc2, 'Nghiên cứu dọc THCS → THPT. So sánh nông thôn–thành thị cùng tỉnh. Can thiệp giảm áp lực học tập. Sử dụng GAD-7 song song MHT để so sánh quốc tế.')

p2 = doc2.add_paragraph()
r2 = p2.add_run('Đánh giá: ⭐⭐⭐ Trung bình-Khá. LPA độc đáo, áp lực học tập OR=11,6 là dữ liệu mạnh, nhưng cỡ mẫu nhỏ và tạp chí bị loại SCIE.')
r2.bold = True; r2.font.name = 'Times New Roman'; r2.font.size = Pt(12)

doc2.save('TomTat_Wen_2020.docx')
sys.stderr.write('TomTat_Wen_2020.docx OK\n')
