# -*- coding: utf-8 -*-
"""Dịch QT35 — Jefferies & Ungar 2020 — Social Anxiety 7 Countries — PLOS ONE"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import *

doc = create_doc()
add_p(doc, 'Link: https://doi.org/10.1371/journal.pone.0239133', size=10)

add_heading(doc, 'Lo âu xã hội ở thanh niên: Nghiên cứu tỷ lệ tại 7 quốc gia', 1)
h = doc.add_paragraph()
r = h.add_run('Social Anxiety in Young People: A Prevalence Study in Seven Countries')
r.font.name = 'Times New Roman'; r.font.size = Pt(13); r.italic = True

add_heading(doc, 'THÔNG TIN THƯ MỤC', 2)
add_info_table(doc, [
    ('Tiêu đề gốc', 'Social Anxiety in Young People: A Prevalence Study in Seven Countries'),
    ('Tiêu đề dịch', 'Lo âu xã hội ở thanh niên: Nghiên cứu tỷ lệ tại 7 quốc gia'),
    ('Tác giả', 'Philip Jefferies, Michael Ungar'),
    ('Cơ quan', 'Resilience Research Centre, Faculty of Health, Dalhousie University, Halifax, Nova Scotia, Canada'),
    ('Tạp chí', 'PLOS ONE (Q1, IF ≈ 3,7)'),
    ('Xuất bản', '2020, Vol. 15(9), e0239133, 18 trang'),
    ('DOI', '10.1371/journal.pone.0239133'),
    ('Loại NC', 'Nghiên cứu cắt ngang đa quốc gia'),
    ('Mẫu', '6.825 thanh niên 16–29 tuổi (TB = 22,84; SD = 3,97) từ 7 quốc gia: Brazil, Trung Quốc, Indonesia, Nga, Thái Lan, Mỹ, VIỆT NAM'),
])

add_page_ref(doc, '1–18', 'PLOS ONE', 'Vol. 15(9), 2020')

# ========== TÓM TẮT ==========
add_heading(doc, 'TÓM TẮT', 2)
add_p(doc, 'Lo âu xã hội là hiện tượng tăng nhanh, được cho là ảnh hưởng không cân xứng đến thanh niên. Trong nghiên cứu này, chúng tôi khám phá tỷ lệ lo âu xã hội trên toàn thế giới sử dụng khảo sát tự báo cáo trên 6.825 cá nhân (nam = 3.342, nữ = 3.428, khác = 55), 16–29 tuổi (TB = 22,84; SD = 3,97), từ 7 quốc gia được chọn vì sự đa dạng văn hóa và kinh tế: Brazil, Trung Quốc, Indonesia, Nga, Thái Lan, Mỹ, và Việt Nam.')

p = doc.add_paragraph()
r = p.add_run('Người tham gia hoàn thành Thang Lo âu Tương tác Xã hội (SIAS — Social Interaction Anxiety Scale). Tỷ lệ lo âu xã hội toàn cầu cao hơn đáng kể so với báo cáo trước: hơn 1/3 (36%) đạt ngưỡng Rối loạn Lo âu Xã hội (SAD). Tỷ lệ và mức độ triệu chứng KHÔNG khác biệt theo giới tính nhưng khác biệt theo tuổi, quốc gia, tình trạng việc làm, trình độ học vấn, và khu vực sống. 1/6 (18%) tự nhận không có lo âu xã hội nhưng vẫn đạt hoặc vượt ngưỡng SAD.')
r.font.name = 'Times New Roman'; r.font.size = Pt(12)
r.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

# ========== TÓM TẮT ĐÁNH GIÁ NHANH ==========
add_heading(doc, 'TÓM TẮT ĐÁNH GIÁ NHANH', 2)
for b in [
    'PLOS ONE Q1 — 7 quốc gia đa dạng VH/KT, BAO GỒM VIỆT NAM.',
    'Mẫu lớn: 6.825 thanh niên 16–29 tuổi.',
    'Tỷ lệ SAD toàn cầu = 36% — cao hơn nhiều so với báo cáo trước (12% US lifetime).',
    'VIỆT NAM: SAD = 30,7% — cứ 3 thanh niên VN có 1 đạt ngưỡng lo âu xã hội.',
    'Mỹ cao nhất (57,6%), Indonesia thấp nhất (22,9%).',
    'KHÔNG khác biệt giới tính — trái với đa số NC lo âu khác (thường nữ > nam).',
    '18% "false negatives" — có SAD nhưng không nhận ra.',
]:
    add_p(doc, f'• {b}')

add_p(doc, 'Hạn chế:', bold=True)
for b in [
    'Mẫu market research (Unilever/Clear) — không phải khảo sát SKTT chuyên biệt.',
    'SIAS đo lo âu tương tác xã hội — không bao gồm lo âu biểu diễn (performance anxiety).',
    'Cắt ngang — không theo dõi dọc.',
    'Ngưỡng SAD (SIAS ≥ 29) — tỷ lệ có thể cao do đo sàng lọc, không phải chẩn đoán lâm sàng.',
]:
    add_p(doc, f'• {b}')

add_p(doc, 'Hướng cải thiện:', bold=True)
for b in [
    'NC lo âu xã hội tại VN với chẩn đoán lâm sàng (DAWBA, DISC) thay vì sàng lọc.',
    'So sánh chi tiết VN vs các nước — yếu tố văn hóa nào ảnh hưởng.',
    'Đánh giá "false negatives" tại VN — thanh niên VN có nhận ra lo âu XH không?',
]:
    add_p(doc, f'• {b}')

# ========== GIỚI THIỆU ==========
add_page_ref(doc, '1–3', 'PLOS ONE', 'Vol. 15(9), 2020')
add_heading(doc, '1. GIỚI THIỆU', 2)

add_p(doc, 'Lo âu xã hội (social anxiety) là nỗi sợ bị đánh giá tiêu cực bởi người khác trong các tình huống xã hội hoặc biểu diễn. Biểu hiện: sợ nói trước đám đông, gặp gỡ người lạ, duy trì cuộc trò chuyện, tiếp xúc với người có thẩm quyền. Người có lo âu xã hội ít biểu cảm khuôn mặt hơn, tránh ánh mắt, khó khăn khởi xướng và duy trì hội thoại.')

add_p(doc, 'Tác động rộng: nạn nhân bắt nạt, bỏ học sớm, ít bạn bè hơn, ít kết hôn, ly hôn nhiều hơn, ít sinh con, vắng mặt nơi làm việc nhiều hơn, hiệu suất kém. Tỷ lệ SAD suốt đời lên đến 12% ở Mỹ, 0,8% ở châu Âu, 0,2% ở Trung Quốc — nhưng các con số này có thể thấp vì đo khác nhau.')

add_p(doc, 'Lo âu xã hội khác lo âu tổng quát (GAD) — tập trung vào sợ đánh giá xã hội, không phải lo lắng chung. Tăng đặc biệt ở thanh niên do áp lực xã hội, MXH, và giai đoạn phát triển bản sắc. Nghiên cứu nhằm đánh giá tỷ lệ tại 7 quốc gia đa dạng, bao gồm Việt Nam.')

# ========== PHƯƠNG PHÁP ==========
add_page_ref(doc, '3–5', 'PLOS ONE', 'Vol. 15(9), 2020')
add_heading(doc, '2. PHƯƠNG PHÁP', 2)

add_p(doc, '2.1. Thiết kế', bold=True)
add_p(doc, 'Phân tích thứ cấp dữ liệu thu thập bởi Edelman Intelligence cho chiến dịch nghiên cứu thị trường (Clear/Unilever), tháng 11/2019. Khảo sát trực tuyến 20 phút. Tuyển chọn ngẫu nhiên qua Dynata, OMI, GMO Research — các công ty giữ panel đại diện quốc gia. Phê duyệt đạo đức: Đại học Dalhousie.')

add_p(doc, '2.2. Người tham gia', bold=True)
add_p(doc, '6.825 người (nam 3.342, nữ 3.428, khác 55), 16–29 tuổi (TB = 22,84; SD = 3,97), từ 7 quốc gia. Mẫu quota: giới, vùng, tuổi đại diện dân số mỗi nước. Bồi thường bằng hệ thống điểm (voucher, từ thiện).')

add_p(doc, '2.3. Thang đo', bold=True)
add_p(doc, 'Thang Lo âu Tương tác Xã hội 17 mục (SIAS-17 — Social Interaction Anxiety Scale; Mattick & Clarke, 1998). Gốc: 20 mục, nhưng phân tích chỉ dùng 17 mục thuận chiều (loại 3 mục nghịch đảo). Đo nỗi sợ và lo âu liên quan tương tác xã hội. Ngưỡng SAD: SIAS-17 ≥ 29 (tương đương SIAS-20 ≥ 34; Heimberg et al., 2014). Phân biệt tốt giữa lâm sàng và không lâm sàng, giữa SAD và GAD. Đã kiểm tra đa văn hóa.')

add_p(doc, '2.4. Phân tích', bold=True)
add_p(doc, 'ANOVA, chi-square, hồi quy đa biến. So sánh theo giới, tuổi (16–17, 18–24, 25–29), quốc gia, tình trạng việc làm, khu vực sống, trình độ học vấn. Cũng đánh giá tự nhận thức (self-perception) vs ngưỡng SIAS.')

# ========== KẾT QUẢ ==========
add_page_ref(doc, '5–10', 'PLOS ONE', 'Vol. 15(9), 2020')
add_heading(doc, '3. KẾT QUẢ', 2)

add_heading(doc, 'Bảng 1. Tỷ lệ lo âu xã hội theo quốc gia', 3)
add_table(doc,
    ['Quốc gia', 'Điểm TB (SD)', 'SAD ≥ 29 (%)', 'So sánh'],
    [
        ['Brazil', '26,18 (15,23)', '42,4%', ''],
        ['Trung Quốc', '22,30 (13,52)', '32,1%', ''],
        ['Indonesia', '18,94 (13,21)', '22,9%', 'Thấp nhất'],
        ['Nga', '20,78 (12,79)', '27,0%', ''],
        ['Thái Lan', '25,57 (13,92)', '41,4%', ''],
        ['Hoa Kỳ', '30,35 (15,44)', '57,6%', 'Cao nhất'],
        ['VIỆT NAM', '22,68 (11,77)', '30,7%', 'SD thấp nhất'],
        ['TỔNG', '23,82 (14,18)', '36,2%', 'F = 74,85, p < 0,001'],
    ],
    widths=[3.0, 3.5, 2.5, 4.0])
add_p(doc, 'VN có SD thấp nhất (11,77) — ít biến thiên nhất, gợi ý mẫu VN đồng nhất hơn.', size=9, italic=True)

add_heading(doc, 'Bảng 2. Tỷ lệ SAD theo giới tính, tuổi, việc làm, khu vực', 3)
add_table(doc,
    ['Phân nhóm', 'Điểm TB (SD)', 'SAD (%)', 'p'],
    [
        ['Nam', '23,53 (14,12)', '35,6%', 'n.s.'],
        ['Nữ', '24,00 (14,18)', '36,5%', 'KHÔNG khác biệt'],
        ['16–17 tuổi', '21,92 (14,24)', '30,8%', '<0,001'],
        ['18–24 tuổi', '25,33 (13,98)', '40,3%', 'Cao nhất'],
        ['25–29 tuổi', '22,44 (14,22)', '32,8%', ''],
        ['Có việc làm', '23,28 (14,32)', '35,3%', '<0,001'],
        ['Đang học', '23,96 (13,50)', '36,5%', ''],
        ['Thất nghiệp', '26,27 (14,54)', '41,7%', 'Cao nhất'],
        ['Trung tâm đô thị', '22,70 (14,67)', '33,0%', '<0,001'],
        ['Ngoại ô', '25,64 (14,08)', '42,4%', 'Cao nhất'],
        ['Nông thôn', '25,37 (13,91)', '41,9%', ''],
        ['Chưa xong cấp 3', '27,94 (15,07)', '52,0%', '<0,001'],
        ['Xong cấp 3', '23,40 (14,15)', '34,8%', ''],
    ],
    widths=[3.5, 3.0, 2.0, 3.0])

add_p(doc, '3.1. Giới tính', bold=True)
add_p(doc, 'KHÔNG có khác biệt giới tính đáng kể (F(1,6768) = 1,88, n.s.). Đây là phát hiện TRÁI VỚI đa số NC lo âu khác (thường nữ > nam). Có thể do lo âu xã hội ảnh hưởng nam và nữ TƯƠNG ĐƯƠNG, khác với lo âu tổng quát nơi nữ bị ảnh hưởng nhiều hơn.')

add_p(doc, '3.2. Tuổi', bold=True)
add_p(doc, 'Nhóm 18–24 tuổi có lo âu xã hội CAO NHẤT (M = 25,33; SAD = 40,3%), cao hơn đáng kể so với 16–17 (p < 0,001) và 25–29 (p < 0,001). Gợi ý: giai đoạn đại học/đầu sự nghiệp là đỉnh lo âu xã hội.')

add_p(doc, '3.3. Quốc gia', bold=True)
add_p(doc, 'Khác biệt ĐÁNG KỂ giữa các quốc gia (F = 74,85, p < 0,001, ηp² = 0,039). Mỹ cao nhất (57,6% SAD), Indonesia thấp nhất (22,9%). Tương tác quốc gia × tuổi có ý nghĩa: 16–17 tuổi Indonesia thấp nhất (M = 15,70), 25–29 Mỹ cao nhất (M = 30,47). VIỆT NAM: 30,7% — trung bình, nhưng vẫn cao hơn nhiều so với tỷ lệ lâm sàng đã báo cáo trước đó.')

add_p(doc, '3.4. Tự nhận thức (Self-perception)', bold=True)
add_p(doc, '34% tự nhận là có lo âu xã hội, tương đương 36% vượt ngưỡng SIAS. Tuy nhiên:')
add_p(doc, '• 48% — KHÔNG lo âu XH + KHÔNG vượt ngưỡng (đúng)')
add_p(doc, '• 18% — CÓ lo âu XH + CÓ vượt ngưỡng (đúng)')
add_p(doc, '• 16% — tự nhận CÓ nhưng KHÔNG vượt ngưỡng (dương tính giả)')
add_p(doc, '• 18% — tự nhận KHÔNG nhưng CÓ vượt ngưỡng (ÂM TÍNH GIẢ — đáng lo)')
add_p(doc, 'Nghĩa là: 1/6 thanh niên có SAD nhưng KHÔNG BIẾT mình có — "lo âu xã hội ẩn". Tại VN, tỷ lệ này có thể cao hơn do nhận thức SKTT còn thấp.')

add_heading(doc, 'Bảng 3. Mối quan tâm theo bối cảnh — so sánh VN với các nước', 3)
add_table(doc,
    ['Mối quan tâm', 'VN (xếp hạng)', 'Trung Quốc', 'Indonesia', 'Mỹ', 'Tổng'],
    [
        ['Nói với người có thẩm quyền', '3', '1', '1', '5', '1'],
        ['Nói về bản thân/cảm xúc', '5', '8', '5', '2', '4'],
        ['Tiếp xúc người lạ', '—', '—', '—', '—', '—'],
        ['Hòa nhập đồng nghiệp', '16', '17', '15', '16', '16'],
        ['Gặp tình cờ người quen', '—', '—', '—', '—', 'Thấp nhất'],
    ],
    widths=[4.0, 2.5, 2.5, 2.5, 2.5, 2.5])
add_p(doc, 'Ghi chú: Xếp hạng từ 1 (lo lắng nhất) đến 17. VN: nói với người có thẩm quyền xếp hạng 3 — phù hợp văn hóa tôn trọng thứ bậc.', size=9, italic=True)

# ========== THẢO LUẬN ==========
add_page_ref(doc, '10–15', 'PLOS ONE', 'Vol. 15(9), 2020')
add_heading(doc, '4. THẢO LUẬN', 2)

add_p(doc, 'Tỷ lệ SAD = 36% — cao hơn ĐÁNG KỂ so với báo cáo trước (12% US lifetime). Có thể do: (1) nghiên cứu tập trung vào thanh niên 16–29 (nhóm tuổi nhạy cảm nhất), (2) thay đổi xã hội (MXH, giảm tương tác trực tiếp), (3) tăng nhận thức về SKTT.')

add_p(doc, 'Không khác biệt giới tính — trái với kỳ vọng. Có thể do: (1) lo âu xã hội ảnh hưởng tương đương nam nữ (khác GAD), (2) thang SIAS đo tương tác (không phải biểu diễn — nơi có thể có chênh lệch giới).')

add_p(doc, 'Khác biệt quốc gia phản ánh yếu tố văn hóa. Mỹ cao nhất — có thể do văn hóa cá nhân chủ nghĩa nhấn mạnh thành tích xã hội. Indonesia thấp nhất — có thể do văn hóa tập thể, ít áp lực thể hiện cá nhân. VN trung bình (30,7%) — thấp hơn Brazil, Thái Lan, Mỹ nhưng cao hơn Indonesia, Nga.')

add_p(doc, '"False negatives" (18% — có SAD nhưng không biết) là phát hiện quan trọng cho chính sách. Cần tăng nhận thức về lo âu xã hội — nhiều người trải nghiệm nhưng coi là "nhút nhát" bình thường, không tìm kiếm giúp đỡ.')

# ========== KẾT LUẬN ==========
add_heading(doc, '5. KẾT LUẬN', 2)
add_p(doc, 'Dữ liệu 6.825 thanh niên từ 7 quốc gia cho thấy tỷ lệ lo âu xã hội rất cao (36% đạt ngưỡng SAD), không khác biệt giới tính, cao nhất ở 18–24 tuổi và người thất nghiệp. VIỆT NAM: 30,7% — cứ 3 thanh niên VN có 1 người có khả năng mắc SAD. 18% là "lo âu ẩn" — có SAD nhưng không nhận ra. Gợi ý rằng cần sàng lọc lo âu xã hội rộng rãi ở thanh niên VN, không chỉ lo âu tổng quát.')

# ========== TLTK ==========
add_heading(doc, 'TÀI LIỆU THAM KHẢO', 2)
for ref in [
    'Jefferies, P. & Ungar, M. (2020). Social anxiety in young people: A prevalence study in seven countries. PLOS ONE, 15(9), e0239133.',
    'Mattick, R.P. & Clarke, J.C. (1998). Development and validation of measures of social phobia scrutiny fear and social interaction anxiety. Behaviour Research and Therapy, 36, 455–470.',
    'Heimberg, R.G. et al. (2014). Social anxiety disorder in DSM-5. Depression and Anxiety, 31, 472–479.',
    'Kessler, R.C. et al. (2005). Lifetime prevalence and age-of-onset distributions of DSM-IV disorders. Archives of General Psychiatry, 62, 593–602.',
]:
    add_p(doc, ref, size=10)
add_p(doc, '(Xem đầy đủ trong bài gốc — 60+ TLTK)', size=10, italic=True)

# ========== VIẾT TẮT ==========
add_abbreviation_table(doc, [
    ('SIAS', 'Social Interaction Anxiety Scale — Thang Lo âu Tương tác Xã hội'),
    ('SAD', 'Social Anxiety Disorder — Rối loạn Lo âu Xã hội'),
    ('GAD', 'Generalized Anxiety Disorder — Rối loạn Lo âu Tổng quát'),
    ('DASS-21', 'Depression Anxiety Stress Scale 21'),
    ('DSM', 'Diagnostic and Statistical Manual of Mental Disorders'),
    ('ISCED', 'International Standard Classification of Education'),
    ('VN', 'Việt Nam'),
    ('SKTT', 'Sức khỏe tâm thần'),
    ('VTN', 'Vị thành niên'),
    ('MXH', 'Mạng xã hội'),
])

# ========== PHẢN BIỆN ==========
add_red_heading(doc, 'QUAN ĐIỂM PHẢN BIỆN / CRITICAL REVIEW')

add_red(doc, 'Điểm mạnh:', bold=True)
for s in [
    'PLOS ONE Q1. 7 quốc gia đa dạng VH/KT — BAO GỒM VIỆT NAM (dữ liệu rất hiếm).',
    'Mẫu lớn: 6.825 thanh niên, mẫu quota đại diện quốc gia.',
    'SIAS — thang đo chuyên biệt cho lo âu xã hội, đã kiểm chứng đa văn hóa.',
    'Phát hiện "false negatives" (18%) — đóng góp quan trọng cho chính sách sàng lọc.',
    'So sánh đa quốc gia trên cùng thang đo — hiếm trong lĩnh vực.',
]:
    add_red(doc, f'• {s}')

add_red(doc, 'Hạn chế chi tiết:', bold=True)
for s in [
    'Mẫu market research (Unilever/Clear) — không phải khảo sát SKTT chuyên biệt. Động cơ tham gia: điểm thưởng, không phải SKTT.',
    'SIAS đo lo âu TƯƠNG TÁC xã hội — không bao gồm lo âu BIỂU DIỄN (performance: phát biểu, thi cử). Có thể bỏ sót loại lo âu XH phổ biến ở VTN.',
    'Sàng lọc (SIAS ≥ 29) — KHÔNG phải chẩn đoán lâm sàng. Tỷ lệ 36% có thể cao hơn thực tế (so với chẩn đoán DISC/DAWBA).',
    'Cắt ngang — không theo dõi dọc, không xác lập nhân quả.',
    'Tuổi 16–29 — rộng hơn phạm vi VTN thông thường (10–19). Bao gồm cả thanh niên trưởng thành.',
    'Khảo sát trực tuyến — thiên lệch chọn (loại bỏ thanh niên không có internet, đặc biệt ở VN nông thôn).',
    'Ngưỡng SIAS-17 ≥ 29 — chuyển đổi từ SIAS-20 ≥ 34, có thể không chính xác hoàn toàn.',
]:
    add_red(doc, f'• {s}')

add_red(doc, 'Khoảng trống nghiên cứu / Research Gap:', bold=True)
for s in [
    'VN có dữ liệu trong bài này (SAD = 30,7%) nhưng CHƯA CÓ NC lo âu xã hội chuyên biệt ở VTN VN.',
    'Lo âu xã hội KHÁC lo âu tổng quát — nhưng đa số NC VN (GAD-7, DASS-21) chỉ đo lo âu tổng quát. Cần sàng lọc SAD riêng.',
    'Tỷ lệ "false negatives" ở VN — thanh niên VN coi lo âu XH là "nhút nhát" → không tìm giúp đỡ. Cần NC đánh giá nhận thức SKTT.',
    'So sánh VN vs Indonesia, Thái Lan — cùng Đông Nam Á nhưng tỷ lệ rất khác (22,9% vs 30,7% vs 41,4%). Yếu tố văn hóa nào giải thích?',
    'Cần NC lo âu xã hội ở VTN VN THCS/THPT (10–17) — bài này chỉ có 16–29.',
    'MXH có thể đặc biệt ảnh hưởng lo âu XH (so sánh, phản hồi, cyberbullying) — nhưng bài này không đo MXH.',
]:
    add_red(doc, f'• {s}')

# ========== SAVE ==========
outpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), '35_SocialAnxiety_7Countries_2020.docx')
doc.save(outpath)
print(f'Saved: {outpath}')
