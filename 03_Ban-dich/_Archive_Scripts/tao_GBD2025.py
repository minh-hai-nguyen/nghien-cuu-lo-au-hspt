# -*- coding: utf-8 -*-
"""
Tao ban dich DOCX cho bai bao GBD 2021 ASEAN (Lancet Public Health 2025)
Su dung tao_dich_template.py
"""
import sys, os, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import (
    create_doc, add_link_and_qr, add_heading, add_page_ref, add_p,
    add_red, add_red_heading, add_table, add_info_table, add_image,
    add_abbreviation_table
)

JOURNAL = 'Lancet Public Health'
VOL = 'Vol 10, June 2025'
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   '03_GBD_ASEAN_2025_Lancet.docx')

doc = create_doc()

# ============================================================
# 1. LINK + QR
# ============================================================
add_link_and_qr(doc, 'https://doi.org/10.1016/S2468-2667(25)00098-2',
                'QR_GBD2025.png')

# ============================================================
# 2. TITLE
# ============================================================
add_heading(doc,
    'BÀI BÁO: Dịch tễ học và gánh nặng 10 rối loạn tâm thần '
    'tại ASEAN, 1990\u20132021', level=1)

# ============================================================
# 3. INFO TABLE
# ============================================================
add_info_table(doc, [
    ('Tiêu đề gốc',
     'Epidemiology and burden of ten mental disorders in the '
     'Association of Southeast Asian Nations (ASEAN) countries, '
     '1990\u20132021: findings from the Global Burden of Disease Study 2021'),
    ('Tác giả',
     'GBD 2021 ASEAN Mental Disorders Collaborators '
     '(Anna Szucs, Stephanie C C van der Lubbe, Jorge Arias de la Torre, '
     'Jose M Valderas, Simon I Hay, v.v.)'),
    ('Tạp chí', 'The Lancet Public Health'),
    ('Năm / Tập / Trang', '2025; Vol 10: e480\u2013e491 (12 trang)'),
    ('DOI', '10.1016/S2468-2667(25)00098-2'),
    ('Loại bài', 'Nghiên cứu gốc (Original Research) \u2013 Truy cập mở CC BY 4.0'),
    ('Tài trợ', 'Quỹ Gates (Gates Foundation)'),
    ('Thiết kế NC',
     'Phân tích dữ liệu GBD 2021: ước tính tỷ lệ hiện mắc, YLDs, YLLs, '
     'DALYs cho 10 rối loạn tâm thần tại 10 quốc gia ASEAN, 1990\u20132021'),
    ('Công cụ mô hình',
     'DisMod-MR 2.1 (Bayesian meta-regression); CODEm cho tử vong'),
    ('Số liệu chính',
     '80,4 triệu ca (95% UI 73,8\u201387,2); tăng 70,0%; '
     'tỷ lệ chuẩn hóa 11,9% (10,9\u201312,9); '
     'DALYs 11,2 triệu (8,54\u201314,3)'),
])

doc.add_paragraph()

# ============================================================
# PAGE e480 \u2013 TÓM TẮT
# ============================================================
add_page_ref(doc, 'e480', JOURNAL, VOL)

add_heading(doc, 'TÓM TẮT (Summary)', 2)

add_p(doc,
    'Bối cảnh (Background): Hiệp hội các Quốc gia Đông Nam Á '
    '(Association of Southeast Asian Nations \u2013 ASEAN), một mạng lưới '
    'địa chính trị và kinh tế gồm mười quốc gia thành viên, công nhận '
    'các rối loạn tâm thần (mental disorders) là một ưu tiên y tế; tuy nhiên, '
    'dữ liệu dịch tễ học còn hạn chế cản trở việc phát triển các chiến lược '
    'hiệu quả để giảm tỷ lệ hiện mắc và gánh nặng bệnh tật. Chúng tôi nhằm '
    'mục đích khảo sát tỷ lệ hiện mắc (prevalence), bệnh tật (morbidity) và '
    'gánh nặng bệnh tật (disease burden) liên quan đến mười rối loạn tâm thần '
    'từ năm 1990 đến 2021 tại ASEAN.', bold=False)

add_p(doc,
    'Phương pháp (Methods): Là một phần của Nghiên cứu Gánh nặng Bệnh tật, '
    'Thương tích và Yếu tố Nguy cơ Toàn cầu (Global Burden of Diseases, '
    'Injuries, and Risk Factors Study \u2013 GBD 2021), chúng tôi phân tích '
    'các ước tính cho rối loạn trầm cảm (depressive disorders), rối loạn lo âu '
    '(anxiety disorders), rối loạn lưỡng cực (bipolar disorders), tâm thần phân liệt '
    '(schizophrenia), rối loạn phổ tự kỷ (autism spectrum disorders), rối loạn hành vi '
    '(conduct disorder), rối loạn tăng động giảm chú ý (attention-deficit hyperactivity '
    'disorder \u2013 ADHD), rối loạn ăn uống (eating disorders), khuyết tật trí tuệ '
    'phát triển tự phát (idiopathic developmental intellectual disability), và các '
    'rối loạn tâm thần khác (other mental disorders) tại mười quốc gia thành viên '
    'ASEAN (Brunei, Campuchia, Indonesia, Lào, Malaysia, Myanmar, Philippines, '
    'Singapore, Thái Lan và Việt Nam). Định nghĩa ca bệnh dựa trên tiêu chuẩn của '
    'Sổ tay Chẩn đoán và Thống kê các Rối loạn Tâm thần (Diagnostic and Statistical '
    'Manual of Mental Disorders \u2013 DSM) hoặc tiêu chuẩn Phân loại Bệnh Quốc tế '
    '(International Classification of Diseases \u2013 ICD). Các ước tính tỷ lệ hiện mắc '
    'theo tuổi, giới tính, năm và địa điểm được tính bằng DisMod-MR 2.1, một công cụ '
    'mô hình hóa hồi quy siêu phân tích Bayes. Gánh nặng bệnh tật được lượng hóa bằng '
    'ước tính số năm sống với khuyết tật (years lived with disability \u2013 YLDs), '
    'số năm mất đi do tử vong sớm (years of life lost \u2013 YLLs) và số năm sống '
    'điều chỉnh theo khuyết tật (disability-adjusted life-years \u2013 DALYs). '
    'Các ước tính được trình bày với khoảng bất định 95% '
    '(95% uncertainty intervals \u2013 UIs).', bold=False)

add_p(doc,
    'Kết quả (Findings): Vào năm 2021, 80,4 triệu (95% UI 73,8\u201387,2) '
    'trường hợp rối loạn tâm thần được ghi nhận trên toàn ASEAN, tăng 70,0% '
    '(63,5\u201377,2) so với năm 1990. Tỷ lệ hiện mắc chuẩn hóa theo tuổi '
    '(age-standardised prevalence) của rối loạn tâm thần là 11,9% (10,9\u201312,9) '
    'vào năm 2021, dao động từ 10,1% (9,1\u201311,3) ở Việt Nam đến 13,2% '
    '(11,6\u201315,3) ở Malaysia, trong đó rối loạn lo âu và rối loạn trầm cảm '
    'là phổ biến nhất. Tỷ lệ hiện mắc chuẩn hóa theo tuổi của rối loạn tâm thần '
    'tăng 6,5% (3,7\u20139,8) giữa năm 1990 và 2021. Rối loạn tâm thần chiếm '
    '11,2 triệu (8,54\u201314,3) DALYs vào năm 2021, tăng 87,4% (81,1\u201394,0) '
    'so với năm 1990. Nhóm tuổi 10\u201314 có gánh nặng bệnh tật cao nhất do rối '
    'loạn tâm thần, chiếm 16,3% (12,7\u201320,5) tổng số DALYs trong nhóm tuổi này. '
    'Mức tăng tương đối lớn nhất về số ca rối loạn tâm thần giữa 1990 và 2021 được '
    'ghi nhận ở người lớn tuổi (182,8% [174,9\u2013192,1] ở nhóm từ 70 tuổi trở lên), '
    'mặc dù thay đổi tương đối nhỏ về tỷ lệ hiện mắc ở các nhóm tuổi này.', bold=False)

add_p(doc,
    'Giải thích (Interpretation): Sự gia tăng tỷ lệ hiện mắc và gánh nặng rối loạn '
    'tâm thần tìm thấy trong nghiên cứu này có thể phản ánh một phần những cải thiện '
    'gần đây trong phát hiện bệnh. Tuy nhiên, rối loạn tâm thần hiện xếp trong mười '
    'nguyên nhân hàng đầu gây gánh nặng bệnh tật tại tất cả các quốc gia ASEAN ngoại '
    'trừ Myanmar, nhấn mạnh nhu cầu cấp thiết về một cách tiếp cận liên ngành toàn diện '
    'để giải quyết khoảng trống trong phòng ngừa và điều trị trên toàn bộ dân số.', bold=False)

add_p(doc, 'Tài trợ (Funding): Quỹ Gates (Gates Foundation).', bold=True)
add_p(doc,
    'Bản quyền: © 2025 Tác giả. Xuất bản bởi Elsevier Ltd. '
    'Đây là bài báo Truy cập Mở (Open Access) theo giấy phép CC BY 4.0.',
    size=10, italic=True)

# ============================================================
# PAGE e481 \u2013 GIỚI THIỆU
# ============================================================
add_page_ref(doc, 'e481', JOURNAL, VOL)
add_heading(doc, 'GIỚI THIỆU (Introduction)', 2)

add_p(doc,
    'Các rối loạn tâm thần (mental disorders) ảnh hưởng đến sức khỏe cá nhân và '
    'phúc lợi xã hội. Vào năm 2021, rối loạn tâm thần đóng góp 17,2% tổng số năm '
    'sống với khuyết tật (YLDs) trên toàn thế giới. "Thúc đẩy sức khỏe tâm thần và '
    'hạnh phúc" là một phần của Mục tiêu Phát triển Bền vững số 3 của Liên Hợp Quốc '
    '(UN Sustainable Development Goal 3). Các nỗ lực quốc tế gần đây, như Sáng kiến '
    'Đặc biệt của Tổ chức Y tế Thế giới (WHO) về Sức khỏe Tâm thần (WHO Special '
    'Initiative for Mental Health) bắt đầu từ năm 2019, đã nỗ lực tăng cường khả năng '
    'tiếp cận chăm sóc sức khỏe tâm thần, đặc biệt tại các quốc gia có thu nhập thấp '
    'và trung bình, nơi cơ sở hạ tầng sức khỏe tâm thần thiết yếu còn thiếu. Tuy nhiên, '
    'ngay cả tại các quốc gia có thu nhập cao với hệ thống chăm sóc sức khỏe tâm thần '
    'được triển khai tốt, những khoảng trống đáng kể trong chẩn đoán và điều trị vẫn '
    'còn tồn tại.')

add_p(doc,
    'Là nền kinh tế lớn thứ năm thế giới, và với dân số gấp 1,5 lần dân số EU, '
    'ASEAN là một mạng lưới kinh tế và địa chính trị bao gồm mười quốc gia có nền '
    'tảng lịch sử, chính trị, văn hóa và tôn giáo đa dạng: Brunei, Campuchia, '
    'Indonesia, Lào, Malaysia, Myanmar, Philippines, Singapore, Thái Lan và Việt Nam. '
    'ASEAN là nơi sinh sống của 672 triệu người, gần 9% dân số thế giới. Khu vực này '
    'bao gồm các quốc gia ở các mức kinh tế xã hội khác nhau: các quốc gia có thu nhập '
    'trung bình thấp (lower-middle-income nations), cụ thể là Campuchia, Myanmar và Lào, '
    'xếp hạng thứ 148, 144 và 139 tương ứng trong tổng số 192 quốc gia trên Chỉ số '
    'Phát triển Con người của Liên Hợp Quốc (UN Human Development Index \u2013 HDI); '
    'các quốc gia có thu nhập trung bình, cụ thể là Philippines (thứ 113), Indonesia '
    '(thứ 112) và Việt Nam (thứ 107); và các quốc gia có thu nhập trung bình cao đến '
    'cao (upper-middle to high-income countries), cụ thể là Thái Lan (thứ 66), Malaysia '
    '(thứ 63), Brunei (thứ 55) và Singapore (thứ 9). Mặc dù có những khác biệt này, '
    'việc thúc đẩy sức khỏe tâm thần là một ưu tiên y tế quan trọng của khu vực theo '
    'Chương trình Nghị sự Phát triển Y tế ASEAN Hậu 2015 (ASEAN Post-2015 Health '
    'Development Agenda) (ưu tiên y tế số 5 trong mục tiêu 5 năm cho giai đoạn 2021\u201325).')

add_p(doc,
    'Các chiến lược quốc gia về phát triển chăm sóc sức khỏe tâm thần khác nhau đáng '
    'kể trong ASEAN. Các quốc gia có thu nhập trung bình thấp có xu hướng có ít cơ sở '
    'hạ tầng sức khỏe tâm thần, và nó đã bị gián đoạn bởi các cuộc xung đột vũ trang '
    'và thảm họa thiên nhiên gần đây. Trong khi các quốc gia có thu nhập trung bình đã '
    'bắt đầu đầu tư vào chăm sóc sức khỏe tâm thần trong hai thập kỷ qua, tình trạng '
    'chẩn đoán thiếu (underdiagnosis) và điều trị thiếu (undertreatment) vẫn phổ biến. '
    'Các quốc gia có thu nhập trung bình cao và cao sở hữu cơ sở hạ tầng sức khỏe tâm '
    'thần, nhưng vẫn phải đối mặt với các rào cản khác đối với chăm sóc sức khỏe tâm '
    'thần, như mức độ hiểu biết và nhận thức cộng đồng chưa tối ưu, kỳ thị văn hóa xã '
    'hội (sociocultural stigma), và phạm vi bảo hiểm và khả năng tiếp cận dịch vụ không '
    'đầy đủ. Những chênh lệch này trở nên đặc biệt rõ ràng trong đại dịch COVID-19.')

add_p(doc,
    'Hiểu được tình trạng hiện tại và xu hướng của rối loạn tâm thần tại ASEAN là rất '
    'quan trọng để xây dựng các can thiệp và chính sách y tế phù hợp. Tuy nhiên, cần '
    'có một phân tích toàn diện về tỷ lệ hiện mắc và gánh nặng bệnh tật của rối loạn '
    'tâm thần trong khu vực. Để giải quyết khoảng trống này, chúng tôi đã tận dụng '
    'Nghiên cứu Gánh nặng Bệnh tật, Thương tích và Yếu tố Nguy cơ Toàn cầu '
    '(GBD) 2021 để khảo sát tỷ lệ hiện mắc, tử vong và bệnh tật '
    'của mười rối loạn tâm thần theo tuổi và giới tính từ 1990 đến 2021 tại mười '
    'quốc gia ASEAN nhằm cung cấp những hiểu biết dịch tễ học về bức tranh rối loạn '
    'tâm thần trong khu vực. Bài báo này được thực hiện như một phần của Mạng lưới '
    'Cộng tác viên GBD (GBD Collaborator Network) và phù hợp với Quy trình GBD '
    '(GBD Protocol).')

# ============================================================
# PAGE e482 \u2013 BỐI CẢNH NGHIÊN CỨU
# ============================================================
add_page_ref(doc, 'e482', JOURNAL, VOL)
add_heading(doc, 'BỐI CẢNH NGHIÊN CỨU (Research in context)', 2)

add_p(doc, 'Bằng chứng trước nghiên cứu này (Evidence before this study)', bold=True)
add_p(doc,
    'Chúng tôi đã tìm kiếm PubMed và Google Scholar các nghiên cứu được xuất bản '
    'từ khi có cơ sở dữ liệu đến ngày 21 tháng 10, 2024, về rối loạn tâm thần tại '
    'ASEAN hoặc bất kỳ quốc gia thành viên nào sử dụng các cụm từ tìm kiếm liên quan. '
    'Kết quả tìm kiếm cho thấy ít nhưng đang tăng các nghiên cứu về tỷ lệ hiện mắc '
    'và gánh nặng rối loạn tâm thần trong khu vực ASEAN. Các nghiên cứu cụ thể theo '
    'quốc gia trong ASEAN khác nhau về thiết kế nghiên cứu và công cụ đánh giá được '
    'sử dụng, dẫn đến các phát hiện không đồng nhất. Trên tất cả các nghiên cứu, '
    'rối loạn lo âu (anxiety disorders) có tỷ lệ hiện mắc cao nhất, dao động từ '
    'khoảng 5% đến 20%, tiếp theo là rối loạn trầm cảm (depressive disorders), với '
    'ước tính tỷ lệ hiện mắc từ 2% đến 15%. Chưa có nghiên cứu nào phân tích gánh '
    'nặng bệnh tật liên quan đến rối loạn tâm thần tại ASEAN.')

add_p(doc, 'Giá trị gia tăng của nghiên cứu này (Added value of this study)', bold=True)
add_p(doc,
    'Nghiên cứu này cung cấp một cái nhìn tổng quan có hệ thống về tỷ lệ hiện mắc '
    'và gánh nặng của rối loạn tâm thần theo tuổi và giới tính trên tất cả các quốc '
    'gia thành viên ASEAN. Kết quả cho thấy các biến đổi theo địa lý, thời gian, '
    'giới tính và tuổi trong dịch tễ học của rối loạn tâm thần. Các phát hiện đáng lo '
    'ngại cho thấy gánh nặng cao của rối loạn tâm thần ở nhóm dân số trẻ và mức tăng '
    'tương đối lớn về số ca bệnh ở nhóm dân số lớn tuổi. Bằng việc nhấn mạnh những '
    'vấn đề này, công trình cung cấp bằng chứng định lượng ủng hộ cải cách chính sách.')

add_p(doc,
    'Ý nghĩa của tất cả các bằng chứng hiện có (Implications of all the '
    'available evidence)', bold=True)
add_p(doc,
    'Với kế hoạch 5 năm tiếp theo của Chương trình Nghị sự Phát triển Y tế ASEAN '
    'Hậu 2015 cho giai đoạn 2026\u201330 đang đến gần, các phát hiện của chúng tôi '
    'kêu gọi đầu tư nhiều hơn để giải quyết thách thức sức khỏe cộng đồng ngày càng '
    'tăng của rối loạn tâm thần. Vì rối loạn tâm thần ở cấp độ dân số chủ yếu được '
    'thúc đẩy bởi các yếu tố xã hội, môi trường và cấu trúc, ASEAN phải tận dụng '
    'ảnh hưởng địa chính trị của mình để thúc đẩy các thay đổi chính sách đa ngành '
    'và liên chính phủ nhằm vượt qua các rào cản hiện tại đối với chăm sóc sức khỏe '
    'tâm thần.')

# ============================================================
# PAGE e482\u2013e483 \u2013 PHƯƠNG PHÁP
# ============================================================
add_page_ref(doc, 'e482\u2013e483', JOURNAL, VOL)
add_heading(doc, 'PHƯƠNG PHÁP (Methods)', 2)

add_p(doc, 'Tổng quan (Overview)', bold=True)
add_p(doc,
    'GBD 2021 đánh giá có hệ thống 371 bệnh và thương tích, cùng với 88 yếu tố '
    'nguy cơ, trên 204 quốc gia và vùng lãnh thổ. Nghiên cứu ước tính các chỉ số '
    'khác nhau, như tỷ lệ mới mắc (incidence), tỷ lệ hiện mắc (prevalence), tử vong '
    'theo nguyên nhân cụ thể, số năm mất đi do tử vong sớm (YLLs; đại diện cho gánh '
    'nặng sức khỏe tử vong), số năm sống với khuyết tật (YLDs; đại diện cho gánh nặng '
    'sức khỏe không tử vong), và số năm sống điều chỉnh theo khuyết tật (DALYs; tổng '
    'của YLLs và YLDs), tổng hợp dữ liệu từ nhiều nguồn bao gồm các khảo sát quốc gia, '
    'sổ đăng ký bệnh tật và tài liệu đã xuất bản. Bài báo tuân thủ các khuyến nghị '
    'GATHER.')

add_p(doc, 'Định nghĩa ca bệnh (Case definitions)', bold=True)
add_p(doc,
    'Các rối loạn tâm thần bao gồm trong GBD 2021 là rối loạn trầm cảm (rối loạn '
    'trầm cảm chủ yếu [major depressive disorder] và loạt trầm cảm [dysthymia]), '
    'rối loạn lo âu, rối loạn lưỡng cực, tâm thần phân liệt, rối loạn phổ tự kỷ, '
    'rối loạn hành vi (conduct disorder), ADHD, rối loạn ăn uống (chán ăn tâm thần '
    '[anorexia nervosa] và chứng ăn vô độ [bulimia nervosa]), khuyết tật trí tuệ phát '
    'triển tự phát, và một nhóm các rối loạn tâm thần khác. Để đảm bảo tính so sánh, '
    'các ca bệnh được định nghĩa sử dụng tiêu chuẩn từ DSM và ICD.')

# ============================================================
# PAGE e483\u2013e484 \u2013 PHƯƠNG PHÁP (tiếp)
# ============================================================
add_page_ref(doc, 'e483\u2013e484', JOURNAL, VOL)

add_p(doc, 'Nguồn dữ liệu (Data sources)', bold=True)
add_p(doc,
    'Dữ liệu được thu thập từ các khảo sát đa quốc gia và quốc gia, hệ thống '
    'đăng ký dân sự (cho nguyên nhân tử vong), và tài liệu đã xuất bản. Các nguồn '
    'được yêu cầu phải đại diện cho dân số chung. Dữ liệu liên quan được xác định '
    'cho tất cả các quốc gia ASEAN ngoại trừ Brunei và Campuchia.')

add_p(doc, 'Tỷ lệ hiện mắc (Prevalence)', bold=True)
add_p(doc,
    'Ước tính tỷ lệ hiện mắc bao gồm quy trình hai bước. Đầu tiên, các ước tính '
    'dịch tễ học từ các nghiên cứu khác nhau được đánh giá về sai số tiềm ẩn phát '
    'sinh từ sự khác biệt trong công cụ chẩn đoán, thời gian hồi tưởng và loại '
    'người phỏng vấn. Thứ hai, dữ liệu đã điều chỉnh sai số được phân tích bằng '
    'DisMod-MR 2.1. Một điều chỉnh cụ thể được thực hiện để tính đến tác động '
    'của COVID-19.')

add_p(doc, 'Tỷ lệ mức độ nặng (Severity proportions)', bold=True)
add_p(doc,
    'Tỷ lệ mức độ nặng được tính toán để phản ánh các mức độ khuyết tật khác nhau '
    'liên quan đến một rối loạn cụ thể, như nhẹ, trung bình và nặng. Tỷ lệ mức độ '
    'nặng được áp dụng cho tổng số ca hiện mắc, và ước tính tỷ lệ hiện mắc cho '
    'mỗi mức độ nặng được tính bằng DisMod-MR 2.1.')

add_p(doc, 'Tử vong và số năm mất đi do tử vong sớm (Mortality and YLLs)', bold=True)
add_p(doc,
    'Tử vong được ước tính cho rối loạn ăn uống, cụ thể là chán ăn tâm thần '
    '(anorexia nervosa), sử dụng mô hình Tập hợp Nguyên nhân Tử vong chuẩn '
    'của GBD (CODEm). YLLs được tính bằng cách nhân số tử vong ước tính do '
    'nguyên nhân cụ thể với tuổi thọ kỳ vọng chuẩn tại tuổi tương ứng.')

add_p(doc, 'Số năm sống với khuyết tật (Years lived with disability)', bold=True)
add_p(doc,
    'YLDs được tính bằng cách nhân tỷ lệ hiện mắc theo mức độ nặng cụ thể với '
    'các trọng số khuyết tật (disability weights) tương ứng. Trọng số khuyết tật '
    'đại diện cho mức mất mát sức khỏe tương đối do một tình trạng sức khỏe cụ thể, '
    'biểu thị trên thang đo từ 0 (không mất mát) đến 1 (tương đương tử vong). '
    'DALYs được tính bằng cách cộng YLLs và YLDs theo địa điểm, năm, tuổi, giới '
    'tính và nguyên nhân. Đối với tất cả các ước tính, khoảng bất định 95% (95% UIs) '
    'được suy ra sử dụng phân vị thứ 2,5 và 97,5 từ 500 lần rút mẫu từ phân phối '
    'hậu nghiệm được tạo bởi lấy mẫu Monte Carlo chuỗi Markov.')

# ============================================================
# PAGE e484\u2013e485 \u2013 KẾT QUẢ
# ============================================================
add_page_ref(doc, 'e484\u2013e485', JOURNAL, VOL)
add_heading(doc, 'KẾT QUẢ (Results)', 2)

add_p(doc,
    'Vào năm 2021, có ước tính 80,4 triệu (95% UI 73,8\u201387,2) trường hợp rối '
    'loạn tâm thần tại ASEAN (bảng). Số trường hợp cao nhất được ghi nhận tại Indonesia '
    '(32,9 triệu [30,3\u201335,8] trường hợp), Philippines (14,0 triệu [12,9\u201315,1]) '
    'và Việt Nam (10,1 triệu [9,09\u201311,3]), tương ứng với quy mô dân số các quốc gia. '
    'So với năm 1990, tổng số trường hợp rối loạn tâm thần tại ASEAN tăng 70,0% '
    '(63,5\u201377,2).')

# ============================================================
# MAIN DATA TABLE
# ============================================================
add_heading(doc,
    'Bảng: Tổng số và số ca hiện mắc phân theo giới tính và tỷ lệ hiện mắc '
    'chuẩn hóa theo tuổi của rối loạn tâm thần năm 1990 và 2021', level=3)

headers = [
    'Quốc gia',
    '1990\nCa (triệu)',
    '1990\nTỷ lệ (%)',
    '2021\nCa (triệu)',
    '2021\nTỷ lệ (%)',
    'Thay đổi\nsố ca (%)',
    'Thay đổi\ntỷ lệ (%)'
]

rows = [
    ['ASEAN', '47,3\n(42,9\u201352,3)', '11,2%\n(10,2\u201312,3)',
     '80,4\n(73,8\u201387,2)', '11,9%\n(10,9\u201312,9)',
     '70,0%\n(63,5\u201377,2)', '6,5%\n(3,7\u20139,8)'],
    ['  Nam', '22,6\n(20,4\u201325,0)', '10,7%\n(9,8\u201311,7)',
     '37,6\n(34,5\u201340,8)', '11,3%\n(10,4\u201312,3)',
     '66,2%\n(60,1\u201372,7)', '5,5%\n(2,7\u20138,3)'],
    ['  Nữ', '24,6\n(22,3\u201327,3)', '11,6%\n(10,6\u201312,7)',
     '42,7\n(39,1\u201346,8)', '12,5%\n(11,4\u201313,7)',
     '73,4%\n(66,5\u201381,8)', '7,5%\n(4,2\u201311,2)'],
    ['Brunei', '0,027\n(0,024\u20130,029)', '11,0%\n(10,1\u201311,9)',
     '0,051\n(0,045\u20130,058)', '11,7%\n(10,5\u201313,1)',
     '93,3%\n(79,7\u2013110,1)', '6,9%\n(\u22120,2\u201315,5)'],
    ['Campuchia', '1,12\n(0,993\u20131,27)', '12,4%\n(11,2\u201313,8)',
     '2,12\n(1,86\u20132,44)', '12,9%\n(11,3\u201314,8)',
     '89,0%\n(70,7\u2013113,1)', '3,8%\n(\u22125,8\u201316,6)'],
    ['Indonesia', '19,0\n(17,3\u201320,8)', '10,8%\n(9,9\u201311,8)',
     '32,9\n(30,3\u201335,8)', '11,8%\n(10,9\u201312,8)',
     '73,8%\n(67,2\u201382,1)', '9,0%\n(5,7\u201313,0)'],
    ['Lào', '0,465\n(0,414\u20130,525)', '12,5%\n(11,3\u201313,9)',
     '0,901\n(0,782\u20131,04)', '12,8%\n(11,1\u201314,8)',
     '93,8%\n(73,2\u2013116,8)', '2,6%\n(\u22127,9\u201314,9)'],
    ['Malaysia', '1,85\n(1,66\u20132,06)', '11,7%\n(10,6\u201312,9)',
     '4,13\n(3,63\u20134,78)', '13,2%\n(11,6\u201315,3)',
     '122,8%\n(100,0\u2013148,2)', '12,6%\n(1,7\u201325,1)'],
    ['Myanmar', '4,44\n(3,94\u20135,01)', '11,5%\n(10,3\u201312,8)',
     '6,64\n(5,72\u20137,63)', '11,9%\n(10,3\u201313,6)',
     '49,5%\n(35,0\u201367,2)', '4,0%\n(\u22126,5\u201315,3)'],
    ['Philippines', '6,77\n(6,15\u20137,41)', '12,0%\n(11,0\u201313,0)',
     '14,0\n(12,9\u201315,1)', '12,9%\n(11,9\u201313,9)',
     '106,3%\n(100,6\u2013113,1)', '7,5%\n(5,2\u201310,3)'],
    ['Singapore', '0,372\n(0,343\u20130,405)', '12,3%\n(11,4\u201313,4)',
     '0,653\n(0,586\u20130,730)', '12,3%\n(11,0\u201313,8)',
     '75,6%\n(62,3\u201389,4)', '\u22120,6%\n(\u22127,9\u20136,2)'],
    ['Thái Lan', '7,02\n(6,36\u20137,78)', '12,4%\n(11,3\u201313,7)',
     '8,88\n(7,98\u20139,92)', '12,8%\n(11,5\u201314,4)',
     '26,4%\n(16,0\u201339,8)', '3,3%\n(\u22124,6\u201313,1)'],
    ['Việt Nam', '6,26\n(5,61\u20136,94)', '9,9%\n(9,1\u201311,0)',
     '10,1\n(9,09\u201311,3)', '10,1%\n(9,1\u201311,3)',
     '61,1%\n(48,3\u201376,0)', '1,9%\n(\u22125,0\u201310,3)'],
]

add_table(doc, headers, rows, widths=[2.5, 2.3, 2.3, 2.3, 2.3, 2.3, 2.3])

add_p(doc,
    'Giá trị trong ngoặc đơn là khoảng bất định 95% (95% uncertainty intervals). '
    'Dữ liệu đếm được làm tròn đến ba chữ số có nghĩa.',
    size=9, italic=True)

doc.add_paragraph()

# ============================================================
# PAGE e485\u2013e486 \u2013 KẾT QUẢ (tiếp): Tỷ lệ hiện mắc
# ============================================================
add_page_ref(doc, 'e485\u2013e486', JOURNAL, VOL)

add_p(doc,
    'Về tỷ lệ hiện mắc chuẩn hóa theo tuổi (age-standardised prevalence), ước tính '
    '11,9% (95% UI 10,9\u201312,9) dân số ASEAN mắc rối loạn tâm thần vào năm 2021. '
    'Tỷ lệ hiện mắc dao động từ 10,1% (9,1\u201311,3) ở Việt Nam đến 13,2% '
    '(11,6\u201315,3) ở Malaysia (Hình 1A). Ở hầu hết các quốc gia ASEAN, tỷ lệ hiện '
    'mắc rối loạn tâm thần cao hơn ở nữ giới so với nam giới; tuy nhiên, điều ngược '
    'lại được quan sát thấy ở Brunei, Singapore và Thái Lan.')

add_p(doc,
    'Giữa 1990 và 2021, tỷ lệ hiện mắc chuẩn hóa theo tuổi trên toàn ASEAN tăng '
    '6,5% (95% UI 3,7\u20139,8). Trong khi hầu hết các quốc gia có tỷ lệ hiện mắc '
    'tương đối ổn định, mức tăng được ghi nhận ở Malaysia (12,6% [1,7\u201325,1]), '
    'Indonesia (9,0% [5,7\u201313,0]) và Philippines (7,5% [5,2\u201310,3]). Trong các '
    'nhóm tuổi, mức tăng lớn nhất về tỷ lệ hiện mắc rối loạn tâm thần được quan sát '
    'thấy ở nhóm 15\u201319 tuổi, nơi tỷ lệ hiện mắc tăng 10,8% (6,9\u201315,1). '
    'Sự gia tăng tỷ lệ hiện mắc đặc biệt rõ ràng giữa 2019 và 2021, trùng với đại '
    'dịch COVID-19. Mặc dù phần trăm thay đổi tỷ lệ hiện mắc ở người trưởng thành '
    'từ 70 tuổi trở lên là nhỏ (2,6%), mức tăng tương đối về số trường hợp là lớn, '
    'ở mức 182,8% (174,9\u2013192,1).')

add_p(doc,
    'Trong các tình trạng bệnh lý, rối loạn lo âu có số ca cao nhất '
    '(29,1 triệu [24,5\u201334,5] ca), tiếp theo là rối loạn trầm cảm '
    '(21,3 triệu [18,8\u201324,1]) và các rối loạn tâm thần khác '
    '(10,3 triệu [8,01\u201312,9]). Ba tình trạng này cũng có tỷ lệ hiện mắc '
    'chuẩn hóa theo tuổi cao nhất: 4,3% cho rối loạn lo âu, 3,1% cho rối loạn '
    'trầm cảm, và 1,5% cho các rối loạn tâm thần khác.')

# ============================================================
# Khác biệt giới tính
# ============================================================
add_p(doc,
    'Có sự khác biệt giữa nam và nữ trong tỷ lệ hiện mắc các rối loạn tâm thần. '
    'Tỷ lệ hiện mắc chuẩn hóa theo tuổi của rối loạn lo âu cao gấp 1,7 lần ở nữ '
    'giới (5,4%) so với nam giới (3,1%). Đối với rối loạn trầm cảm, tỷ lệ hiện mắc '
    'ở nữ giới (3,5%) cao gấp khoảng 1,3 lần so với nam giới (2,7%). Ngược lại, '
    'tỷ lệ hiện mắc cao hơn ở nam giới so với nữ giới đối với ADHD (cao gấp 2,4 lần), '
    'rối loạn phổ tự kỷ (cao gấp 2,0 lần), rối loạn hành vi (cao gấp 1,6 lần) và các '
    'rối loạn tâm thần khác (cao gấp 1,5 lần). Rối loạn lưỡng cực và tâm thần phân '
    'liệt có tỷ lệ hiện mắc tương tự ở nam và nữ.')

# ============================================================
# Nhóm tuổi
# ============================================================
add_p(doc,
    'Trong các nhóm tuổi, các rối loạn tâm thần phổ biến nhất ở trẻ em dưới 5 tuổi '
    'là khuyết tật trí tuệ phát triển tự phát (0,9%) và rối loạn phổ tự kỷ (0,9%). '
    'Ở trẻ em 5\u20139 tuổi, ADHD trở thành rối loạn phổ biến nhất (1,9%). Ở cá nhân '
    'từ 10 tuổi trở lên, rối loạn lo âu là rối loạn phổ biến nhất, đạt đỉnh 5,4% ở '
    'nhóm 20\u201324 tuổi. Tỷ lệ hiện mắc rối loạn trầm cảm cũng tăng từ tuổi vị '
    'thành niên trở đi, đạt đỉnh 5,1% ở nhóm 55\u201359 tuổi.')

# ============================================================
# PAGE e486\u2013e487 \u2013 DALYs
# ============================================================
add_page_ref(doc, 'e486\u2013e487', JOURNAL, VOL)

add_p(doc,
    'Chán ăn tâm thần (anorexia nervosa) là tình trạng duy nhất mà cả tử vong và '
    'YLLs có thể được quy là nguyên nhân tử vong cơ bản. Vào năm 2021, tổng số tử '
    'vong do chán ăn tâm thần tại ASEAN được ước tính là 3,4 (95% UI 1,5\u20135,4).')

add_p(doc,
    'Vào năm 2021, rối loạn tâm thần liên quan đến 11,2 triệu (95% UI 8,54\u201314,3) '
    'DALYs, tương ứng với tỷ lệ DALYs chuẩn hóa theo tuổi là 1583,7 (1203,3\u20132016,3) '
    'trên 100.000 dân tại ASEAN. Tổng số DALYs do rối loạn tâm thần tại ASEAN chiếm '
    '7,2% tổng DALYs toàn cầu liên quan đến rối loạn tâm thần. Giữa 1990 và 2021, '
    'số DALYs do rối loạn tâm thần tại ASEAN tăng 87,4% (81,1\u201394,0), và tỷ lệ '
    'của chúng trong tổng gánh nặng bệnh tật tăng 62,8% (49,1\u201377,8), từ 3,0% '
    'lên 4,9%.')

add_p(doc,
    'So sánh giữa các quốc gia ASEAN, tổng số DALYs liên quan đến rối loạn tâm thần '
    'năm 2021 cao nhất tại Indonesia (4,58 triệu), Philippines (1,96 triệu) và Việt Nam '
    '(1,44 triệu). Khi xem xét tỷ lệ DALYs chuẩn hóa theo tuổi (Hình 1B), Malaysia '
    'có tỷ lệ cao nhất với ước tính 1866,7 DALYs trên 100.000 cá nhân, tiếp theo là '
    'Philippines (1738,8) và Campuchia (1717,3). Giữa 1990 và 2021, mức tăng lớn nhất '
    'về gánh nặng rối loạn tâm thần được quan sát tại Malaysia, nơi tổng số DALYs '
    'tăng 133,6%.')

# ============================================================
# IMAGES \u2013 Hình 1
# ============================================================
add_page_ref(doc, 'e487', JOURNAL, VOL)

add_image(doc, 'GBD2025_figure1_p7.png', width_cm=14,
          caption='Hình 1: (A) Tỷ lệ hiện mắc chuẩn hóa theo tuổi và (B) tỷ lệ DALYs chuẩn hóa theo tuổi do rối loạn tâm thần tại ASEAN, 2021 (Figure 1: Age-standardised prevalence of and DALY rate attributable to mental disorders in the ASEAN, 2021)')
add_image(doc, 'GBD2025_figure2_p8.png', width_cm=14,
          caption='Hình 2: Xếp hạng nguyên nhân gánh nặng bệnh tật (DALYs) tại từng quốc gia ASEAN, 1990 và 2021 (Figure 2: Ranking of causes of DALYs in ASEAN countries, 1990 and 2021)')

# ============================================================
# PAGE e487\u2013e488 \u2013 Đóng góp theo tình trạng bệnh
# ============================================================
add_page_ref(doc, 'e487\u2013e488', JOURNAL, VOL)

add_p(doc,
    'Các tình trạng đóng góp nhiều DALYs nhất trên toàn ASEAN là rối loạn lo âu '
    '(anxiety disorders), chiếm 31,0% tổng DALYs do rối loạn tâm thần trong khu vực, '
    'rối loạn trầm cảm (depressive disorders) (29,6%) và tâm thần phân liệt '
    '(schizophrenia) (12,4%; Hình 2). Ở nữ giới, rối loạn lo âu chiếm tỷ lệ lớn '
    'nhất DALYs trong các rối loạn tâm thần, trong khi rối loạn trầm cảm là nguyên '
    'nhân đóng góp hàng đầu ở nam giới (Hình 3).')

add_p(doc,
    'Trong trẻ em và thanh thiếu niên trên toàn ASEAN, gánh nặng rối loạn tâm thần '
    'cao nhất ở nhóm 10\u201314 tuổi (16,3% [12,7\u201320,5] tổng DALYs), dao động '
    'từ 13,8% tổng gánh nặng bệnh tật ở Myanmar đến trên 20% ở Malaysia (20,7%), '
    'Brunei (25,5%) và Singapore (28,2%).')

add_p(doc,
    'Nhìn chung, rối loạn tâm thần nằm trong mười nguyên nhân hàng đầu gây DALYs '
    'năm 2021 tại tất cả các quốc gia ASEAN, ngoại trừ Myanmar, nơi chúng xếp hạng '
    'thứ 11 (Hình 4). Rối loạn tâm thần nằm trong năm nguyên nhân chính gây gánh '
    'nặng bệnh tật tại Singapore (đóng góp 9,2% tổng DALYs), Brunei (7,6%) và '
    'Malaysia (6,7%) năm 2021. Những thay đổi lớn nhất về gánh nặng tương đối của '
    'rối loạn tâm thần được ghi nhận tại Lào (thứ hạng tăng từ thứ 15 lên thứ 7) '
    'và Campuchia (từ thứ 14 lên thứ 7; Hình 4).')

# ============================================================
# PAGE e488\u2013e490 \u2013 BÀN LUẬN
# ============================================================
add_page_ref(doc, 'e488\u2013e490', JOURNAL, VOL)
add_heading(doc, 'BÀN LUẬN (Discussion)', 2)

add_p(doc,
    'Theo hiểu biết của chúng tôi, đây là nghiên cứu đầu tiên khảo sát tỷ lệ hiện '
    'mắc và gánh nặng của rối loạn tâm thần tại ASEAN. Kết quả chỉ ra rằng vào năm '
    '2021, có ước tính 80,4 triệu trường hợp rối loạn tâm thần, phản ánh mức tăng '
    'khoảng 70% kể từ năm 1990. Tỷ lệ hiện mắc chuẩn hóa theo tuổi tăng vừa phải '
    '6,5% trong cùng giai đoạn, với mức tăng đáng chú ý trong đại dịch COVID-19 '
    'giữa 2019 và 2021. Trong các nhóm tuổi, trẻ em và thanh thiếu niên có gánh '
    'nặng bệnh tật cao nhất liên quan đến rối loạn tâm thần, trong khi số ca tăng '
    'nhiều nhất ở người lớn tuổi. Sự biến động vừa phải được quan sát giữa các quốc '
    'gia, với tỷ lệ hiện mắc chuẩn hóa theo tuổi chung dao động từ 10,1% ở Việt Nam '
    'đến 13,2% ở Malaysia.')

add_p(doc,
    'Sự gia tăng của rối loạn tâm thần được thúc đẩy bởi sự tương tác phức tạp '
    'của các yếu tố cá nhân, xã hội, môi trường và cấu trúc. Một số động lực lớn '
    'nhất của rối loạn tâm thần được nêu trong Ủy ban Lancet 2018 về sức khỏe tâm '
    'thần toàn cầu và phát triển bền vững tồn tại trong ASEAN, bao gồm nghèo đói, '
    'mất an ninh lương thực và tiếp cận không đầy đủ đến chăm sóc y tế và cơ hội '
    'giáo dục. Sự phát triển kinh tế xã hội nhanh chóng trong thập kỷ qua đã làm '
    'trầm trọng thêm bất bình đẳng ở một số quốc gia, làm sâu thêm khoảng cách '
    'nông thôn \u2013 thành thị. Đồng thời, bạo loạn chính trị và bạo lực tại những '
    'nơi như Myanmar và Philippines đã gây ra chấn thương tâm lý lan rộng trong dân '
    'số. Sự gia tăng toàn cầu về rối loạn tâm thần do đại dịch COVID-19 cũng ảnh '
    'hưởng đến các quốc gia ASEAN. Các sự kiện khí hậu khắc nghiệt đã làm tăng thêm '
    'tính dễ tổn thương của dân số với rối loạn tâm thần.')

add_p(doc,
    'Là một mạng lưới hợp tác địa chính trị với các mối liên hệ lịch sử và văn hóa '
    'sâu sắc, ASEAN nên triển khai các nỗ lực tập thể để thúc đẩy sự ổn định khu vực, '
    'tiến bộ xã hội và khả năng phục hồi trước các mối đe dọa môi trường và khí hậu, '
    'qua đó bảo vệ sức khỏe tâm thần trong khu vực.')

add_p(doc,
    'Trong thập kỷ qua, nhiều quốc gia ASEAN đã đẩy mạnh nỗ lực cải thiện chăm sóc '
    'sức khỏe tâm thần. Philippines đã thông qua Luật Sức khỏe Tâm thần (Mental Health '
    'Act) vào năm 2018. Indonesia mở rộng bảo hiểm y tế quốc gia toàn diện cho rối '
    'loạn tâm thần vào năm 2014 và khởi động sáng kiến quốc gia vào năm 2010 để xóa '
    'bỏ pasung \u2013 tập tục truyền thống cách ly hoặc hạn chế thể chất các cá nhân '
    'mắc bệnh tâm thần. Tại Malaysia, luật mới được triển khai vào năm 2010 đưa điều '
    'trị sức khỏe tâm thần bắt buộc lên tiêu chuẩn nhân quyền quốc tế. Mặc dù những '
    'thay đổi chính sách này đang bắt đầu có tác động tích cực, kỳ thị xã hội (social '
    'stigma) và rào cản văn hóa vẫn phổ biến. Ngay cả tại Singapore, quốc gia có thu '
    'nhập cao nhất ASEAN, khoảng 78,6% người trưởng thành mắc rối loạn tâm thần không '
    'nhận được chăm sóc sức khỏe tâm thần phù hợp theo một nghiên cứu từ năm 2016.')

add_p(doc,
    'Tình trạng thiếu hụt chuyên gia sức khỏe tâm thần vẫn là một thách thức phổ biến '
    'trên toàn ASEAN. Tỷ lệ bác sĩ tâm thần trên dân số (psychiatrist-to-population '
    'ratios) trong ASEAN dao động từ 0,2 trên 100.000 dân ở Myanmar và Philippines đến '
    '3,7 trên 100.000 ở Brunei và 4,3 trên 100.000 ở Singapore, vẫn thấp hơn nhiều '
    'so với mức trung bình của châu Âu là 9,7 trên 100.000.')

add_p(doc,
    'Phù hợp với tài liệu toàn cầu, rối loạn trầm cảm, rối loạn lo âu và rối loạn '
    'ăn uống phổ biến hơn ở nữ giới so với nam giới tại ASEAN. Những khác biệt này '
    'được liên kết với sự kết hợp của các yếu tố sinh học, tâm lý và môi trường, như '
    'phơi nhiễm với lạm dụng tình dục trẻ em, bạo lực bạn tình (intimate partner '
    'violence) và bất bình đẳng giới cơ cấu. Bạo lực bạn tình đối với phụ nữ là phổ '
    'biến trong ASEAN, với tỷ lệ hiện mắc trong đời dao động từ 15% đến 44%.')

add_p(doc,
    'Tác động ngày càng tăng của rối loạn tâm thần đối với thanh thiếu niên là đặc '
    'biệt đáng lo ngại. Rối loạn tâm thần chiếm khoảng một phần tư tổng gánh nặng '
    'bệnh tật ở nhóm 10\u201319 tuổi tại các quốc gia ASEAN có thu nhập cao, cụ thể '
    'là Brunei và Singapore. Tỷ lệ hiện mắc rối loạn tâm thần theo tuổi đã tăng hơn '
    '10% ở nhóm 15\u201319 tuổi giữa 1990 và 2021, với mức tăng đáng chú ý giữa '
    '2019 và 2021.')

add_p(doc,
    'Tỷ lệ hiện mắc rối loạn tâm thần ở người lớn tuổi là một vấn đề sức khỏe cộng '
    'đồng cấp bách khác trong ASEAN, nơi tỷ lệ dân số trên 60 tuổi được dự báo sẽ '
    'vượt quá 22% vào năm 2050. Kết quả chỉ ra rằng, được thúc đẩy bởi tuổi thọ kỳ '
    'vọng dài hơn và dân số ngày càng già, số trường hợp rối loạn tâm thần ở người '
    'lớn tuổi đã tăng 182,8% trong 30 năm qua, mặc dù tỷ lệ hiện mắc không tăng.')

# ============================================================
# PAGE e490 \u2013 HẠN CHẾ
# ============================================================
add_page_ref(doc, 'e490\u2013e491', JOURNAL, VOL)
add_heading(doc, 'Hạn chế (Limitations)', 3)

add_p(doc,
    'Nghiên cứu này có một số hạn chế. Thứ nhất, do các điều chỉnh cụ thể được thực '
    'hiện để tính đến tác động của COVID-19 trong GBD 2021, hầu hết các thay đổi về '
    'tỷ lệ hiện mắc và gánh nặng rối loạn tâm thần theo thời gian xảy ra giữa 2019 '
    'và 2021, chủ yếu được thúc đẩy bởi mức tăng ước tính trong tỷ lệ hiện mắc và '
    'gánh nặng của rối loạn trầm cảm chủ yếu (major depressive disorder) và rối loạn '
    'lo âu. Thứ hai, các ước tính cho một số quốc gia và nhóm tuổi có độ bất định cao, '
    'vì dữ liệu dịch tễ học đại diện quốc gia về rối loạn tâm thần còn thiếu hoặc '
    'vắng mặt. Thứ ba, sự gia tăng ca rối loạn tâm thần được thúc đẩy bởi nhiều yếu '
    'tố ngoài sự gia tăng thực sự về khởi phát bệnh, bao gồm thay đổi nhân khẩu học, '
    'cải thiện tiếp cận dịch vụ sàng lọc và tiến bộ trong công cụ chẩn đoán. Thứ tư, '
    'trong GBD, rối loạn tâm thần được định nghĩa sử dụng câu trả lời khảo sát về '
    'tần suất triệu chứng hoặc định nghĩa lâm sàng dựa trên tiêu chuẩn DSM và ICD. '
    'Với mức độ hiểu biết sức khỏe tâm thần thấp và tình trạng chẩn đoán thiếu lan '
    'rộng trong khu vực, tỷ lệ hiện mắc và gánh nặng được báo cáo có thể đã đánh '
    'giá thấp hơn tình trạng thực tế. Cuối cùng, gánh nặng tử vong sớm chỉ được ước '
    'tính cho chán ăn tâm thần \u2013 khung GBD hiện tại không tính đến hiệu ứng '
    'trung gian (mediation effect) của tự gây thương tích.')

# ============================================================
# KẾT LUẬN
# ============================================================
add_heading(doc, 'Kết luận', 3)
add_p(doc,
    'Nghiên cứu nhấn mạnh sự gia tăng quan trọng về tỷ lệ hiện mắc rối loạn tâm '
    'thần và gánh nặng liên quan tại các quốc gia ASEAN. Trong khi các tổ chức quốc '
    'gia và quốc tế đã nỗ lực để hạn chế xu hướng này, còn nhiều việc phải làm để '
    'vượt qua các thách thức hiện có ở các cấp độ môi trường, xã hội và cấu trúc. '
    'Sự quan tâm đến sức khỏe tâm thần vẫn đang tăng lên trong các quốc gia ASEAN, '
    'được chứng minh bởi nghiên cứu quốc gia và khu vực ngày càng tăng, hứa hẹn '
    'mang lại sự hiểu biết chính xác hơn về bức tranh sức khỏe tâm thần và việc '
    'hoạch định chính sách dựa trên bằng chứng sau đó.')

# ============================================================
# TÁC GIẢ
# ============================================================
add_page_ref(doc, 'e491', JOURNAL, VOL)
add_heading(doc, 'Nhóm Cộng tác viên GBD 2021 về Rối loạn Tâm thần ASEAN', 3)
add_p(doc,
    'Anna Szucs, Stephanie C C van der Lubbe, Jorge Arias de la Torre, '
    'Jose M Valderas, Simon I Hay, Catherine Bisignano, Brooks W Morgan, '
    'Swetha Acharya, v.v. (danh sách đầy đủ trong bài báo gốc).',
    size=10, italic=True)
add_p(doc,
    'Đồng tác giả cao cấp (Joint senior authors): Damian F Santomauro, Marie Ng.',
    size=10, italic=True)
add_p(doc,
    'Tài trợ: Quỹ Gates (Gates Foundation). Các nhà tài trợ không có vai trò '
    'trong thiết kế nghiên cứu, thu thập dữ liệu, phân tích hoặc viết báo cáo.',
    size=10, italic=True)

# ============================================================
# PHẦN PHẢN BIỆN (RED)
# ============================================================
doc.add_page_break()
add_red_heading(doc, 'PHẦN PHẢN BIỆN / ĐÁNH GIÁ CHUYÊN MÔN')

add_red(doc, 'A. ĐIỂM MẠNH (Strengths)', bold=True)
add_red(doc,
    '1. Phân tích toàn diện đầu tiên: Đây là nghiên cứu đầu tiên khảo sát tỷ lệ '
    'hiện mắc và gánh nặng của rối loạn tâm thần trên toàn bộ 10 quốc gia ASEAN, '
    'sử dụng khung GBD 2021 với phương pháp chuẩn hóa cho phép so sánh xuyên quốc gia.')
add_red(doc,
    '2. Quy mô và độ bao phủ lớn: Phân tích 10 rối loạn tâm thần, từ 1990 đến 2021, '
    'phân tầng theo tuổi, giới tính và quốc gia. Sử dụng DisMod-MR 2.1 (Bayesian '
    'meta-regression) đảm bảo ước tính nhất quán.')
add_red(doc,
    '3. Phương pháp luận nghiêm ngặt: Sử dụng khoảng bất định 95% (95% UIs) '
    'thay vì khoảng tin cậy truyền thống, phản ánh chính xác hơn sự không chắc chắn '
    'trong ước tính. Điều chỉnh cho sai số đo lường (measurement bias) thông qua '
    'phân tích tổng hợp mạng lưới (network meta-analyses).')
add_red(doc,
    '4. Ý nghĩa chính sách rõ ràng: Kết quả được đặt trong bối cảnh Chương trình '
    'Nghị sự Phát triển Y tế ASEAN 2026\u201330, cung cấp bằng chứng kịp thời cho '
    'hoạch định chính sách khu vực.')

add_red(doc, 'B. HẠN CHẾ VÀ ĐIỂM YẾU (Limitations & Weaknesses)', bold=True)
add_red(doc,
    '1. Thiếu dữ liệu gốc: Không có dữ liệu dịch tễ học đại diện quốc gia cho '
    'Brunei và Campuchia; dữ liệu cho các quốc gia khác không liên tục, đặc biệt '
    'trước COVID-19. Điều này làm tăng độ bất định của ước tính.')
add_red(doc,
    '2. Hiệu ứng COVID-19 chi phối: Hầu hết thay đổi theo thời gian tập trung '
    'trong 2019\u20132021 do điều chỉnh đại dịch, khiến khó phân biệt xu hướng '
    'dài hạn thực sự với tác động ngắn hạn của COVID-19.')
add_red(doc,
    '3. Chỉ ước tính tử vong cho chán ăn tâm thần: Khung GBD không tính đến '
    'tử vong trung gian qua tự gây thương tích (self-harm), dẫn đến đánh giá '
    'thấp gánh nặng tử vong thực sự của rối loạn tâm thần.')
add_red(doc,
    '4. Vấn đề văn hóa trong đo lường: Các định nghĩa lâm sàng dựa trên DSM/ICD '
    'có thể không phù hợp hoàn toàn với bối cảnh văn hóa xã hội ASEAN, nơi đau '
    'khổ tâm thần thường được diễn đạt bằng thuật ngữ thể chất hoặc tâm linh.')
add_red(doc,
    '5. Không phân tích yếu tố nguy cơ: Nghiên cứu chỉ mô tả tỷ lệ hiện mắc '
    'và gánh nặng, không xác định các yếu tố nguy cơ hoặc yếu tố bảo vệ cụ thể '
    'cho từng quốc gia ASEAN.')

add_red(doc, 'C. KHOẢNG TRỐNG NGHIÊN CỨU (Research Gaps)', bold=True)
add_red(doc,
    '1. Cần nghiên cứu dọc (longitudinal) tại từng quốc gia ASEAN để theo dõi xu '
    'hướng sức khỏe tâm thần độc lập với GBD, đặc biệt ở trẻ em và thanh thiếu niên.')
add_red(doc,
    '2. Cần dữ liệu về khoảng cách điều trị (treatment gap) cụ thể theo quốc gia '
    'và loại rối loạn, bổ sung cho dữ liệu dịch tễ học.')
add_red(doc,
    '3. Cần nghiên cứu về mối liên hệ giữa rối loạn tâm thần và các bệnh không '
    'lây nhiễm (NCDs) tại ASEAN \u2013 hiện chưa được phân tích trong nghiên cứu này.')
add_red(doc,
    '4. Cần phát triển công cụ sàng lọc sức khỏe tâm thần phù hợp với bối cảnh '
    'văn hóa ASEAN để cải thiện chất lượng dữ liệu.')
add_red(doc,
    '5. Cần nghiên cứu kinh tế y tế (health economics) đánh giá chi phí-hiệu quả '
    'của các can thiệp sức khỏe tâm thần trong bối cảnh tài nguyên hạn chế của '
    'các quốc gia ASEAN thu nhập thấp và trung bình.')

# ============================================================
# BẢNG VIẾT TẮT
# ============================================================
doc.add_page_break()
add_abbreviation_table(doc, [
    ('ADHD', 'Attention-Deficit Hyperactivity Disorder \u2013 '
             'Rối loạn tăng động giảm chú ý'),
    ('ASEAN', 'Association of Southeast Asian Nations \u2013 '
              'Hiệp hội các Quốc gia Đông Nam Á'),
    ('CODEm', 'Cause of Death Ensemble model \u2013 '
              'Mô hình Tập hợp Nguyên nhân Tử vong'),
    ('DALYs', 'Disability-Adjusted Life-Years \u2013 '
              'Số năm sống điều chỉnh theo khuyết tật'),
    ('DisMod-MR', 'Disease Model \u2013 Bayesian Meta-Regression \u2013 '
                  'Mô hình bệnh tật hồi quy siêu phân tích Bayes'),
    ('DSM', 'Diagnostic and Statistical Manual of Mental Disorders \u2013 '
            'Sổ tay Chẩn đoán và Thống kê các Rối loạn Tâm thần'),
    ('GATHER', 'Guidelines for Accurate and Transparent Health Estimates Reporting \u2013 '
               'Hướng dẫn Báo cáo Ước tính Y tế Chính xác và Minh bạch'),
    ('GBD', 'Global Burden of Disease \u2013 Gánh nặng Bệnh tật Toàn cầu'),
    ('HDI', 'Human Development Index \u2013 Chỉ số Phát triển Con người'),
    ('ICD', 'International Classification of Diseases \u2013 '
            'Phân loại Bệnh Quốc tế'),
    ('IHME', 'Institute for Health Metrics and Evaluation \u2013 '
             'Viện Đo lường và Đánh giá Sức khỏe'),
    ('UI', 'Uncertainty Interval \u2013 Khoảng bất định'),
    ('YLDs', 'Years Lived with Disability \u2013 '
             'Số năm sống với khuyết tật'),
    ('YLLs', 'Years of Life Lost \u2013 '
             'Số năm mất đi do tử vong sớm'),
])

# ============================================================
# SAVE
# ============================================================
doc.save(OUT)
print(f'Da tao thanh cong: {OUT}')
