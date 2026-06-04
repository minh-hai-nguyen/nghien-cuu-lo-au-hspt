# -*- coding: utf-8 -*-
"""Bài 09: Puyat 2025 - Filipino Youth Depression"""
import sys, io, os
try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except: pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()
add_link_and_qr(doc, 'https://doi.org/10.1017/gmh.2025.39', 'QR_Hoa2024.png')

add_heading(doc, 'Gia tăng tỷ lệ trầm cảm và mở rộng bất bình đẳng nhân khẩu xã hội trong triệu chứng trầm cảm ở thanh niên Philippines: Phát hiện từ hai khảo sát cắt ngang quy mô quốc gia', 1)

add_info_table(doc, [
    ('Tiêu đề gốc', 'Rising prevalence of depression and widening sociodemographic disparities in depressive symptoms among Filipino youth: findings from two large nationwide cross-sectional surveys'),
    ('Tác giả', 'Joseph H. Puyat, Divine L. Salvador, Anna C. Tuazon, Sanny D. Afable'),
    ('Cơ quan', 'University of British Columbia (Canada), University of the Philippines Diliman, University of St Andrews (UK)'),
    ('Tạp chí', 'Global Mental Health, 2025'),
    ('DOI', '10.1017/gmh.2025.39'),
    ('Mẫu', '2013: n = 19.178; 2021: n = 10.949 — Filipinos 15–24 tuổi'),
    ('Công cụ', 'CES-D-11 (Center for Epidemiologic Studies Depression Scale, 11 mục)'),
    ('Thiết kế', 'Phân tích dữ liệu từ 2 khảo sát cắt ngang quốc gia (YAFS4 2013 + YAFS5 2021)'),
])

doc.add_paragraph()
add_page_ref(doc, '1', 'Global Mental Health', '2025')

add_heading(doc, 'TÓM TẮT', 2)
add_p(doc, 'Bối cảnh: Trầm cảm ở thanh niên là mục tiêu quan trọng cho can thiệp sớm do mối liên hệ mạnh với trầm cảm ở người trưởng thành và suy giảm chức năng dài hạn. Tại Philippines — quốc gia thu nhập trung bình thấp (LMIC — Low- and Middle-Income Country) — dữ liệu dịch tễ học hạn chế cản trở việc lập kế hoạch dịch vụ sức khỏe tâm thần (SKTT).')
add_p(doc, 'Phương pháp: Phân tích dữ liệu đại diện quốc gia từ Khảo sát Thanh niên và Vị thành niên Philippines (YAFS — Young Adult Fertility and Sexuality Survey) năm 2013 (n = 19.178) và 2021 (n = 10.949). Sử dụng CES-D-11 (Center for Epidemiologic Studies Depression Scale, phiên bản 11 mục) để ước tính tỷ lệ triệu chứng trầm cảm trung bình đến nặng (MSDS — Moderate to Severe Depressive Symptoms).')
add_p(doc, 'Kết quả: Tỷ lệ MSDS tăng hơn gấp đôi từ 9,6% (2013) lên 20,9% (2021). Tăng mạnh nhất ở nữ (10,8% → 24,3%), nhóm không thuộc giới chuẩn tắc (9,7% → 32,3%), nhóm nghèo nhất (12,3% → 25,1%), và nhóm không đi học (8,5% → 26,5%). Bất bình đẳng nhân khẩu xã hội mở rộng đáng kể.')

add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
add_p(doc, 'Điểm nổi bật:', bold=True)
add_p(doc, '\u2022 Cỡ mẫu rất lớn (n = 30.127 tổng cộng), đại diện quốc gia Philippines — hiếm có ở Đông Nam Á (GBD 2021 ASEAN Collaborators, 2025).', size=11)
add_p(doc, '\u2022 So sánh 2 thời điểm (2013 vs 2021) cho thấy xu hướng tăng gấp đôi — bằng chứng mạnh về gia tăng trầm cảm (Racine và cộng sự, 2021).', size=11)
add_p(doc, '\u2022 Phân tích chi tiết theo SOGIE (Sexual Orientation, Gender Identity and Expression) — rất tiến bộ, ít nghiên cứu ĐNA nào làm (WHO, 2022).', size=11)
add_p(doc, '\u2022 Sử dụng khung SDH (Social Determinants of Health) — nhấn mạnh yếu tố cấu trúc, không chỉ cá nhân.', size=11)

add_p(doc, 'Hạn chế:', bold=True)
add_p(doc, '\u2022 CES-D-11 sàng lọc, không chẩn đoán — tỷ lệ thực có thể khác nhiều (COVID-19 Mental Disorders Collaborators, 2021).', size=11)
add_p(doc, '\u2022 Hai khảo sát khác nhau (YAFS4 vs YAFS5) — không phải theo dõi cùng nhóm (Sedgwick, 2014).', size=11)
add_p(doc, '\u2022 Chỉ đo trầm cảm, không đo lo âu — thiếu bức tranh SKTT toàn diện.', size=11)

add_p(doc, 'Hướng cải thiện:', bold=True)
add_p(doc, '\u2022 Bổ sung đo lo âu (GAD-7) song song với CES-D trong khảo sát tiếp theo (Spitzer và cộng sự, 2006).', size=11)
add_p(doc, '\u2022 Nghiên cứu dọc theo dõi cùng nhóm thanh niên Philippines qua nhiều năm (Caruana và cộng sự, 2015).', size=11)
add_p(doc, '\u2022 So sánh với dữ liệu Việt Nam, Indonesia, Thái Lan để đánh giá xu hướng khu vực (GBD 2021 ASEAN Collaborators, 2025).', size=11)

doc.add_paragraph()

add_page_ref(doc, '2', 'Global Mental Health', '2025')
add_heading(doc, 'GIỚI THIỆU', 2)
add_p(doc, 'Trầm cảm ở thanh niên là vấn đề SKTT phổ biến tại Philippines. Nghiên cứu trước cho thấy 8,9% thanh niên Philippines có MSDS năm 2013 (Puyat và cộng sự, 2021). Đại dịch COVID-19 làm trầm trọng thêm các bất bình đẳng hiện có về SKTT.')
add_p(doc, 'Khung Yếu tố Quyết định Xã hội về Sức khỏe (SDH — Social Determinants of Health) giải thích rằng sức khỏe tâm thần bị định hình bởi điều kiện xã hội, kinh tế và môi trường — không chỉ là vấn đề cá nhân. Can thiệp hiệu quả cần vượt ra ngoài cấp độ cá nhân để giải quyết các điều kiện xã hội nền tảng.')

add_page_ref(doc, '3-4', 'Global Mental Health', '2025')
add_heading(doc, 'PHƯƠNG PHÁP', 2)
add_p(doc, 'Dữ liệu từ YAFS4 (2013, n = 19.178) và YAFS5 (2021, n = 10.949) — khảo sát quốc gia đại diện thanh niên Philippines 15-24 tuổi. CES-D-11 được dịch sang 6 ngôn ngữ Philippines chính (Tagalog, Cebuano, Ilonggo, Waray, Ilocano, Bicol). Ngưỡng MSDS: CES-D-11 > 11 (> 1 độ lệch chuẩn so với trung bình). Phân tích hồi quy log-binomial có trọng số khảo sát.')

add_page_ref(doc, '4-5', 'Global Mental Health', '2025')
add_heading(doc, 'KẾT QUẢ', 2)

add_p(doc, 'Tỷ lệ MSDS theo nhóm (2013 → 2021):', bold=True)
add_table(doc,
    ['Nhóm', '2013', '2021', 'Thay đổi'],
    [
        ['Tổng', '9,6%', '20,9%', '\u2191 +11,3%'],
        ['Nam', '8,4%', '17,0%', '\u2191 +8,6%'],
        ['Nữ', '10,8%', '24,3%', '\u2191 +13,5%'],
        ['Không thuộc giới chuẩn tắc', '9,7%', '32,3%', '\u2191 +22,6%'],
        ['Nghèo nhất (Q1)', '12,3%', '25,1%', '\u2191 +12,8%'],
        ['Giàu nhất (Q5)', '7,6%', '16,9%', '\u2191 +9,3%'],
        ['Không đi học/tiểu học', '8,5%', '26,5%', '\u2191 +18,0%'],
        ['Đại học+', '7,9%', '15,3%', '\u2191 +7,4%'],
        ['Nông thôn', '9,8%', '21,4%', '\u2191 +11,6%'],
        ['Đô thị', '9,3%', '20,2%', '\u2191 +10,9%'],
    ], widths=[4.0, 2.0, 2.0, 2.5])

doc.add_paragraph()
add_p(doc, 'Phát hiện then chốt: Tỷ lệ tăng ở TẤT CẢ nhóm, nhưng tăng NHANH NHẤT ở nhóm yếu thế — mở rộng khoảng cách bất bình đẳng. Nữ có tỷ lệ cao hơn nam 41% (aPR = 1,41). Nhóm không thuộc giới chuẩn tắc cao hơn 60% (aPR = 1,60).')

add_page_ref(doc, '6-7', 'Global Mental Health', '2025')
add_heading(doc, 'THẢO LUẬN', 2)
add_p(doc, 'Sự gia tăng tỷ lệ trầm cảm từ 9,6% lên 20,9% trong 8 năm phản ánh xu hướng toàn cầu — meta-analysis của Racine và cộng sự (2021) ghi nhận tỷ lệ trầm cảm lâm sàng tăng gấp đôi trong COVID-19 (25,2%). Philippines đặc biệt dễ bị tổn thương do nghèo đói, bất bình đẳng, thiên tai, và hệ thống SKTT yếu.')
add_p(doc, 'Sự mở rộng bất bình đẳng theo SOGIE đáng lo ngại — nhóm không thuộc giới chuẩn tắc tăng từ 9,7% lên 32,3% (gấp 3,3 lần). Đây là bằng chứng mạnh cho chính sách SKTT cần tích hợp bình đẳng giới và đa dạng giới.')

add_heading(doc, 'KẾT LUẬN', 2)
add_p(doc, 'Tỷ lệ trầm cảm ở thanh niên Philippines tăng hơn gấp đôi trong 8 năm, với bất bình đẳng mở rộng đáng kể. Cần chiến lược SKTT quốc gia tích hợp khung SDH, ưu tiên nhóm yếu thế: nữ, LGBTQ+, người nghèo, người ít học.')

add_abbreviation_table(doc, [
    ('MSDS', 'Moderate to Severe Depressive Symptoms — Triệu chứng trầm cảm trung bình đến nặng'),
    ('CES-D-11', 'Center for Epidemiologic Studies Depression Scale, 11 mục'),
    ('YAFS', 'Young Adult Fertility and Sexuality Survey — Khảo sát Thanh niên Philippines'),
    ('SDH', 'Social Determinants of Health — Yếu tố quyết định xã hội về sức khỏe'),
    ('SOGIE', 'Sexual Orientation, Gender Identity and Expression'),
    ('LMIC', 'Low- and Middle-Income Country — Nước thu nhập thấp và trung bình'),
    ('aPR', 'Adjusted Prevalence Ratio — Tỷ số tỷ lệ hiện mắc điều chỉnh'),
])

doc.add_paragraph()
add_red_heading(doc, 'PHẢN BIỆN CHI TIẾT')
add_red(doc, 'Điểm mạnh:', bold=True)
add_red(doc, '\u2022 Cỡ mẫu rất lớn, đại diện quốc gia (n = 30.127 tổng) — hiếm có ở Đông Nam Á (GBD 2021 ASEAN Collaborators, 2025).')
add_red(doc, '\u2022 So sánh 2 thời điểm cho thấy xu hướng tăng rõ ràng — không chỉ một điểm cắt (Caruana và cộng sự, 2015).')
add_red(doc, '\u2022 Phân tích SOGIE tiến bộ — ít NC nào ở ĐNA bao gồm nhóm LGBTQ+ (WHO, 2022).')
add_red(doc, '\u2022 CES-D-11 được dịch sang 6 ngôn ngữ Philippines — tính phù hợp văn hóa cao (Bernal và cộng sự, 2009).')

add_red(doc, 'Hạn chế:', bold=True)
add_red(doc, '\u2022 CES-D-11 là sàng lọc, không phải chẩn đoán lâm sàng — 20,9% là triệu chứng, không phải rối loạn (COVID-19 Mental Disorders Collaborators, 2021; Ferrari và cộng sự, 2013).')
add_red(doc, '\u2022 Hai khảo sát khác nhau — không phải nghiên cứu dọc cùng nhóm, nên "tăng gấp đôi" có thể phản ánh khác biệt mẫu (Sedgwick, 2014).')
add_red(doc, '\u2022 Chỉ đo trầm cảm — thiếu lo âu, stress, tự tử, lạm dụng chất (Polanczyk và cộng sự, 2015).')
add_red(doc, '\u2022 Ngưỡng MSDS (> 1 SD) không phải ngưỡng lâm sàng chuẩn hóa — khó so sánh quốc tế (Radloff, 1977).')
add_red(doc, '\u2022 Thiếu phân tích yếu tố nguy cơ cụ thể: mạng xã hội, áp lực học tập, ACEs (Twenge và cộng sự, 2018; Felitti và cộng sự, 1998).')

add_red(doc, 'Khoảng trống nghiên cứu (Gap):', bold=True)
add_red(doc, '\u2022 Cần bổ sung thang đo lo âu (GAD-7) trong YAFS tiếp theo để có bức tranh SKTT toàn diện (Spitzer và cộng sự, 2006).')
add_red(doc, '\u2022 Nghiên cứu dọc theo dõi cùng nhóm thanh niên Philippines 15-24 tuổi qua nhiều năm (Caruana và cộng sự, 2015).')
add_red(doc, '\u2022 So sánh liên quốc gia Đông Nam Á (Philippines vs Việt Nam vs Indonesia) với cùng công cụ đo (GBD 2021 ASEAN Collaborators, 2025).')
add_red(doc, '\u2022 Đánh giá tác động COVID-19 cụ thể — YAFS5 (2021) trùng đại dịch, cần tách ảnh hưởng dịch vs xu hướng dài hạn (Racine và cộng sự, 2021).')
add_red(doc, '\u2022 Nghiên cứu can thiệp SKTT tại trường học Philippines — hiện chỉ có 33% VTN báo cáo rối nhiễu tâm lý nhưng chỉ 2% dùng dịch vụ (Mallari và cộng sự, 2025).')

doc.save('09_Puyat_2025_Filipino.docx')
