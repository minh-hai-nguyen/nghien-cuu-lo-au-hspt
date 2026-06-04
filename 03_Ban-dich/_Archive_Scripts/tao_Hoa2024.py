# -*- coding: utf-8 -*-
"""
Tạo bản dịch DOCX cho bài báo Hoa et al. (2024)
Frontiers in Public Health, Vol. 12, Article 1232856
"""
import sys, os, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Import template functions
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tao_dich_template import (
    create_doc, add_link_and_qr, add_heading, add_page_ref,
    add_p, add_red, add_red_heading, add_table, add_info_table,
    add_abbreviation_table, add_image
)

JOURNAL = 'Frontiers in Public Health'
VOL = 'Vol. 12, 2024'

def page_ref(doc, pg, total=8):
    add_page_ref(doc, f'{pg}/{total}', JOURNAL, VOL)

doc = create_doc()

# ============================================================
# LINK + QR
# ============================================================
add_link_and_qr(doc, 'https://doi.org/10.3389/fpubh.2024.1232856', 'QR_Hoa2024.png')

# ============================================================
# TITLE
# ============================================================
add_heading(doc, 'BÀI BÁO: Triệu chứng lo âu và chiến lược ứng phó ở học sinh THPT tại Việt Nam sau đại dịch COVID-19', level=1)
add_p(doc, 'Anxiety symptoms and coping strategies among high school students in Vietnam after COVID-19 pandemic: a mixed-method evaluation', size=11, italic=True)

# ============================================================
# INFO TABLE
# ============================================================
add_info_table(doc, [
    ('Tác giả', 'Phạm Thị Thu Hoa, Đỗ Thị Trang*, Nguyễn Thị Liên, Ngô Anh Vinh'),
    ('Đơn vị', '1) Khoa Tâm lý học, ĐH KHXH&NV, ĐHQG Việt Nam, Hà Nội\n'
               '2) Trung tâm NC&ƯD Tâm lý Giáo dục Bình Minh, Hà Nội\n'
               '3) Khoa KHGD, Trường ĐH Giáo dục, ĐHQG Việt Nam, Hà Nội\n'
               '4) Khoa Sức khỏe Vị thành niên, BV Nhi Trung ương, Hà Nội'),
    ('Tạp chí', 'Frontiers in Public Health, 2024, Vol. 12, Article 1232856'),
    ('DOI', '10.3389/fpubh.2024.1232856'),
    ('Loại bài', 'Nghiên cứu gốc (Original Research) — Phương pháp hỗn hợp (Mixed Methods)'),
    ('Mẫu', '3.910 học sinh THPT từ 13 trường tại Hà Nội (khảo sát) + 20 HS phỏng vấn sâu'),
    ('Công cụ', 'GAD-7 (Generalized Anxiety Disorder 7-item Scale, α = 0,916); Phỏng vấn sâu bán cấu trúc'),
    ('Thời gian', 'Tháng 10–11/2021 (giai đoạn phong tỏa COVID-19 tại Hà Nội)'),
    ('Kết quả chính', '40,6% HS có triệu chứng lo âu (23,9% nhẹ, 10,9% trung bình, 5,8% nặng).\n'
                      'Nữ > Nam (M = 1,74 vs 1,50, p < 0,01). Lớp 12 > Lớp 10, 11.'),
])
doc.add_paragraph()

# ============================================================
# PAGE 1 — TÓM TẮT
# ============================================================
page_ref(doc, 1)
add_heading(doc, 'TÓM TẮT (Abstract)', 2)

add_p(doc, 'Giới thiệu: Mục tiêu của nghiên cứu hiện tại là khảo sát tỷ lệ học sinh trung học phổ thông (THPT — Trung học Phổ thông) có nguy cơ mắc rối loạn lo âu (anxiety disorder) trong đại dịch COVID-19 (Coronavirus Disease 2019) tại Việt Nam, cũng như các chiến lược ứng phó (coping strategies) được sử dụng trong nhóm đối tượng này.', bold=False)

add_p(doc, 'Phương pháp: Đánh giá được thực hiện thông qua phương pháp hỗn hợp (mixed methods), bao gồm sự kết hợp giữa nghiên cứu cắt ngang (cross-sectional study) và phỏng vấn sâu (in-depth interviews). Mẫu gồm 3.910 học sinh từ 13 trường THPT tại Hà Nội, Việt Nam. Việc đo lường các triệu chứng rối loạn lo âu được thực hiện thông qua Thang đo Rối loạn Lo âu Tổng quát 7 mục (Generalized Anxiety Disorder 7-Item Scale — GAD-7). Các cuộc phỏng vấn sâu đã được tiến hành để hiểu rõ nguyên nhân cơ bản và cơ chế ứng phó.')

add_p(doc, 'Kết quả: Tỷ lệ hiện mắc (prevalence rate) các triệu chứng rối loạn lo âu ở học sinh là 40,6%. Tỷ lệ triệu chứng lo âu nhẹ (mild), trung bình (moderate) và nặng (severe) lần lượt là 23,9%, 10,9% và 5,8%. Phỏng vấn sâu phát hiện nhiều nguồn gây lo âu: kết quả học tập (academic performance), mối quan hệ xã hội (social interactions), thái độ định kiến (prejudicial attitudes) từ vòng tròn xã hội, và kỳ vọng gia đình (familial expectations). Nhiều chiến lược ứng phó đã được ghi nhận.')

add_p(doc, 'Bàn luận: Nghiên cứu xác định tồn tại mức độ lo âu vừa phải ở học sinh THPT tại Hà Nội trong đợt bùng phát COVID-19. Nghiên cứu đã xác định các chỉ số tiềm năng (potential indicators) để nhận diện cá nhân dễ bị tổn thương và đề xuất phát triển các can thiệp có mục tiêu (targeted interventions).')

add_p(doc, 'Từ khóa: rối loạn lo âu (anxiety disorder), chiến lược ứng phó (coping strategies), vị thành niên (adolescent), trung học phổ thông (high school), đại dịch COVID-19 (COVID-19 pandemic)', italic=True)

# ============================================================
# PAGE 1-2 — GIỚI THIỆU
# ============================================================
page_ref(doc, 1)
add_heading(doc, '1. GIỚI THIỆU (Introduction)', 2)

add_p(doc, 'Tuổi vị thành niên (adolescence) là một giai đoạn trưởng thành đặc biệt được đánh dấu bởi những biến đổi sinh lý (physiological), tâm lý (psychological) và xã hội (social) đáng kể, khiến những cá nhân này có khả năng dễ bị tổn thương hơn trước các tác động có hại của những hoàn cảnh căng thẳng (stressful circumstances) (1, 2). Theo một phân tích tổng hợp (meta-analysis) bao gồm 136 nghiên cứu được thực hiện trên các quần thể đa dạng bị ảnh hưởng bởi đại dịch COVID-19, tối thiểu 15–16% dân số chung bị các triệu chứng liên quan đến lo âu hoặc trầm cảm (depression) (3). Trong đại dịch COVID-19, sự gián đoạn giáo dục trên diện rộng đã có tác động đáng kể đến sức khỏe tâm thần (mental wellbeing) của trẻ em và vị thành niên.')

page_ref(doc, 2)

add_p(doc, 'Trước khi xuất hiện đại dịch, đã có những trường hợp được ghi nhận rõ ràng về sự gia tăng lo âu, trầm cảm, lạm dụng chất gây nghiện (substance abuse) và các thách thức sức khỏe tâm thần ở vị thành niên chịu áp lực học tập và xã hội cao (4). Việc áp dụng học tập từ xa (remote learning), hạn chế tụ tập xã hội, thay đổi/ngừng hoạt động thể thao hoặc câu lạc bộ, và tạm dừng các hoạt động trường học tạo ra trở ngại đáng kể cho sự phát triển trí tuệ và cộng đồng của thanh thiếu niên (5, 6).')

add_p(doc, 'Một nghiên cứu trước đó phát hiện rằng học sinh nhận giáo dục trực tuyến (virtual education) có nhiều ngày không khỏe mạnh về tinh thần hơn, nhiều triệu chứng trầm cảm hơn và xu hướng nghĩ đến tự tử so với bạn cùng trang lứa có các hình thức giáo dục khác (7). Việc cung cấp hỗ trợ xã hội (social support) từ bạn bè đồng trang lứa và nhà giáo dục được coi là yếu tố quan trọng trong việc bảo vệ sức khỏe cảm xúc-xã hội (social-emotional wellbeing) của vị thành niên, đặc biệt trong các giai đoạn căng thẳng (8).')

add_p(doc, 'Tại Việt Nam, trong giai đoạn vị thành niên, các cá nhân trải qua một quá trình chuyển đổi đầy thách thức, khiến họ rất dễ bị tổn thương trước các tác động tiêu cực do đại dịch COVID-19 gây ra. Một nghiên cứu trên quy mô dân số cho thấy trong số sinh viên đại học, 16,2% bị lo âu trong đại dịch COVID-19 (9). Để ứng phó với đại dịch, các cơ sở giáo dục trên toàn quốc đã thực hiện đóng cửa tạm thời và yêu cầu học sinh ở nhà. Tuy nhiên, chỉ có một số lượng hạn chế các nghiên cứu được thực hiện để khám phá các thách thức về sức khỏe tâm lý mà nhóm đối tượng này gặp phải trong đợt bùng phát COVID-19 tại Việt Nam. Mục tiêu của nghiên cứu hiện tại là khảo sát tỷ lệ học sinh THPT có nguy cơ mắc rối loạn lo âu trong đại dịch COVID-19 tại Việt Nam, cũng như các chiến lược ứng phó được sử dụng trong nhóm đối tượng này.')

# ============================================================
# PAGE 2-3 — PHƯƠNG PHÁP
# ============================================================
page_ref(doc, 2)
add_heading(doc, '2. TÀI LIỆU VÀ PHƯƠNG PHÁP (Materials and Methods)', 2)

add_heading(doc, '2.1 Bối cảnh nghiên cứu và đối tượng tham gia', 3)
add_p(doc, 'Một nghiên cứu cắt ngang (cross-sectional study) đã được thực hiện trên học sinh THPT tại Hà Nội, Việt Nam, từ tháng 10 đến tháng 11 năm 2021, nhằm đáp ứng lệnh phong tỏa (lockdown) được áp đặt tại Hà Nội do đại dịch COVID-19. Trong giai đoạn này, học sinh được yêu cầu tham gia học tập từ xa tại nhà.')

add_p(doc, 'Tiêu chuẩn tham gia: (1) độ tuổi từ 14 đến 17; (2) cư trú tại Hà Nội và đang theo học tại một trường THPT được chọn; (3) đồng ý rõ ràng tham gia nghiên cứu.')

add_p(doc, 'Phương pháp sử dụng kỹ thuật lấy mẫu nhiều giai đoạn (multi-stage sampling technique). Danh sách các trường THPT tại Hà Nội được phân tầng (stratified) theo vị trí địa lý (nông thôn/thành thị) và phân loại cơ sở (tư thục/công lập), tạo ra bốn nhóm. Mười ba trường THPT được chọn ngẫu nhiên để đảm bảo mỗi nhóm có ít nhất ba trường. Toàn bộ học sinh đủ điều kiện trong các cơ sở giáo dục này đều được chọn. Một đoàn hệ (cohort) gồm 3.910 học sinh đăng ký từ 13 cơ sở giáo dục trung học tại Hà Nội. Nghiên cứu đã được phê duyệt bởi hội đồng đánh giá thể chế (institutional review board) của Đại học Quốc gia Việt Nam.')

page_ref(doc, 3)
add_heading(doc, '2.2 Đo lường và công cụ (Measurement and instrument)', 3)
add_p(doc, 'Nghiên cứu sử dụng Google Form để xây dựng khảo sát điện tử. Mỗi khảo sát dự kiến mất khoảng 10–15 phút. Nghiên cứu thu thập dữ liệu nhân khẩu học-xã hội (sociodemographic data), dữ liệu về rối loạn lo âu, và thông tin liên quan về đại dịch COVID-19 thông qua bảng hỏi có cấu trúc (structured questionnaire). Bảng hỏi đã trải qua thử nghiệm thí điểm (pilot test) với nhóm 5 vị thành niên trước khi phổ biến trực tuyến.')

add_p(doc, 'Đối tượng tham gia cung cấp dữ liệu về các đặc điểm nhân khẩu học-xã hội, bao gồm giới tính, loại cơ sở giáo dục, bậc học và vị trí địa lý. Thang đo Rối loạn Lo âu Tổng quát 7 mục (GAD-7) được sử dụng như công cụ đánh giá (10, 11). Người trả lời đánh giá tần suất xuất hiện triệu chứng trên thang đánh giá từ 0 đến 3. Hệ thống phân loại gồm bốn mức: không lo âu (0–4), lo âu nhẹ (5–9), lo âu trung bình (10–14) và lo âu nặng (15–21). Hệ số Cronbach\'s alpha đạt giá trị 0,916.')

add_p(doc, 'Trong phân tích định tính, 20 học sinh THPT được chọn ngẫu nhiên từ cơ sở dữ liệu và được liên hệ qua điện thoại để xin chấp thuận có thông tin đầy đủ (informed consent) cho cuộc phỏng vấn ghi âm. Các đối tượng được tuyển cho đến khi đạt điểm bão hòa (saturation).')

add_heading(doc, '2.3 Phân tích thống kê (Statistical analysis)', 3)
add_p(doc, 'SPSS phiên bản 20.0 được sử dụng để phân tích dữ liệu. Mức ý nghĩa (significance level) được đặt ở 5% (p ≤ 0,05). Giá trị trung bình (mean) và độ lệch chuẩn (SD — standard deviation) cho dữ liệu định lượng; tần suất (frequency) và tỷ lệ phần trăm cho biến định tính. Độ tin cậy nhất quán nội tại (internal consistency reliability) được kiểm tra bằng Cronbach\'s alpha (≥ 0,7 là chấp nhận được).')

add_p(doc, 'Phân tích dữ liệu định tính sử dụng phân tích nội dung định tính (qualitative content analysis). Quy trình: xem xét toàn diện các câu trả lời → xác định các phân đoạn có ý nghĩa → mã hóa chủ đề → tổng hợp thành các lớp phân loại (taxonomic classes) → phân tích chủ đề (thematic analysis).')

# ============================================================
# PAGE 3-4 — KẾT QUẢ
# ============================================================
page_ref(doc, 3)
add_heading(doc, '3. KẾT QUẢ (Results)', 2)

# BẢNG 1 — Nhân khẩu học
add_p(doc, 'Bảng 1. Đặc điểm nhân khẩu học của người trả lời (n = 3.910)', bold=True)
add_table(doc,
    ['Đặc điểm', 'Phân nhóm', 'Tần suất (n)', 'Phần trăm (%)'],
    [
        ['Giới tính (Gender)', 'Nam (Male)', '1.851', '47,3'],
        ['', 'Nữ (Female)', '2.009', '51,4'],
        ['', 'Khác (Others)', '50', '1,3'],
        ['Vị trí (Location)', 'Thành thị (Urban)', '2.052', '52,5'],
        ['', 'Nông thôn (Rural)', '1.858', '47,5'],
        ['Loại trường (Type)', 'Công lập (Public)', '2.069', '52,9'],
        ['', 'Tư thục (Private)', '1.841', '47,1'],
        ['Khối lớp (Grade)', 'Lớp 10', '1.381', '35,3'],
        ['', 'Lớp 11', '1.127', '28,8'],
        ['', 'Lớp 12', '1.402', '35,9'],
    ],
    widths=[4, 3.5, 3, 3]
)
add_p(doc, 'Trong số 3.910 học sinh, phần lớn là nữ (51,4%), học ở khu vực thành thị (52,5%), trường công lập (52,9%) và lớp 12 (35,9%).', size=11)
doc.add_paragraph()

page_ref(doc, 4)

# BẢNG 2 — GAD-7 item scores
add_p(doc, 'Bảng 2. Hồ sơ công cụ Rối loạn Lo âu Tổng quát (GAD-7)', bold=True)
add_table(doc,
    ['TT', 'Mục (Items)', 'Hoàn toàn không\nN (%)', 'Vài ngày\nN (%)', 'Hơn nửa số ngày\nN (%)', 'Gần như mỗi ngày\nN (%)', 'TB'],
    [
        ['1', 'Cảm thấy bồn chồn, lo lắng hoặc bất an', '1.574 (40,3)', '1.658 (42,9)', '359 (9,2)', '319 (8,2)', '0,89'],
        ['2', 'Không thể dừng hoặc kiểm soát lo lắng', '2.320 (59,3)', '1.072 (27,4)', '303 (7,7)', '215 (5,5)', '0,85'],
        ['3', 'Lo lắng quá nhiều về những điều khác nhau', '2.029 (51,9)', '1.117 (28,6)', '416 (10,6)', '348 (8,9)', '0,96'],
        ['4', 'Khó thư giãn', '2.111 (54,0)', '1.187 (30,4)', '356 (9,1)', '256 (6,5)', '0,89'],
        ['5', 'Bồn chồn đến mức khó ngồi yên', '2.729 (69,8)', '833 (21,3)', '205 (5,2)', '143 (3,7)', '0,75'],
        ['6', 'Dễ bực bội hoặc cáu kỉnh', '1.754 (44,9)', '1.247 (31,9)', '453 (11,6)', '456 (11,7)', '1,01'],
        ['7', 'Cảm thấy sợ hãi, như thể điều gì đó khủng khiếp có thể xảy ra', '2.637 (67,4)', '824 (21,1)', '244 (6,2)', '205 (5,2)', '0,83'],
    ],
    widths=[1, 4.5, 2.5, 2.5, 2.5, 2.5, 1.5]
)
add_p(doc, 'Triệu chứng phổ biến nhất: "Dễ bực bội hoặc cáu kỉnh" (TB = 1,01), tiếp theo "Lo lắng quá nhiều về những điều khác nhau" (TB = 0,96). Tỷ lệ HS có nguy cơ mắc rối loạn lo âu: 40,6% (nhẹ 23,9%, trung bình 10,9%, nặng 5,8%).', size=11)
doc.add_paragraph()

# BẢNG 3 — Lo âu theo nhân khẩu học
add_p(doc, 'Bảng 3. Điểm rối loạn lo âu theo các đặc điểm nhân khẩu học khác nhau', bold=True)
add_table(doc,
    ['Đặc điểm', 'Phân nhóm', 'n', '%', 'TB (Mean)', 'SD', 'p'],
    [
        ['Giới tính', 'Nam (Male)', '1.851', '47,3', '1,50', '0,83', '<0,01'],
        ['', 'Nữ (Female)', '2.009', '51,4', '1,74', '0,92', ''],
        ['Vị trí', 'Thành thị (Urban)', '2.052', '52,5', '1,68', '0,92', '<0,01'],
        ['', 'Nông thôn (Rural)', '1.858', '47,5', '1,59', '0,86', ''],
        ['Loại trường', 'Công lập (Public)', '2.069', '52,9', '1,61', '0,88', '0,08'],
        ['', 'Tư thục (Private)', '1.841', '47,1', '1,66', '0,91', ''],
        ['Khối lớp', 'Lớp 10', '1.381', '35,3', '1,53', '0,85', '<0,01'],
        ['', 'Lớp 11', '1.127', '28,8', '1,64', '0,87', ''],
        ['', 'Lớp 12', '1.402', '35,9', '1,72', '0,94', ''],
    ],
    widths=[3, 3, 2, 2, 2, 2, 2]
)
add_p(doc, 'Có sự khác biệt có ý nghĩa thống kê trong điểm GAD-7 liên quan đến giới tính (nữ TB = 1,74 vs nam TB = 1,50, p < 0,01), vị trí (thành thị > nông thôn, p < 0,01) và khối lớp (lớp 12 > lớp 10, p < 0,01). Sự khác biệt theo loại trường không có ý nghĩa thống kê (p = 0,08).', size=11)
doc.add_paragraph()

# ============================================================
# PAGE 4-5 — PHÂN TÍCH ĐỊNH TÍNH
# ============================================================
page_ref(doc, 4)
add_heading(doc, '3.1 Phân tích định tính về thách thức, nguyên nhân và chiến lược ứng phó', 3)

add_p(doc, 'Có nhiều nguyên nhân gây rối loạn lo âu ở học sinh THPT tại Hà Nội, trong đó công việc và học tập là nguyên nhân hàng đầu.')

add_p(doc, '"Đợt COVID-19 vừa rồi, em mới vào lớp 10 nên em khá chủ quan và mải chơi, bỏ bê việc học cộng thêm học trực tuyến, trường quản lý ít chặt chẽ hơn so với đi học trực tiếp, nên có vài môn em không thích, như thể bị mất hết. Bây giờ em học lớp 11 và quay lại trường học trực tiếp, mất gốc khiến em khá khó khăn trong việc học và kết quả học tập rất tệ so với mặt bằng chung, em đã chịu rất nhiều áp lực vì điều đó. Vì thế, mẹ em thường có những lời nói nặng nề khiến em càng hoang mang hơn" (Một nam sinh trường công lập).', italic=True, size=11)

page_ref(doc, 5)

add_p(doc, 'Nguyên nhân thứ hai là các mối quan hệ xã hội (social relationships). Học sinh thường cảm thấy tự ti (inferior) so với bạn bè và lo lắng khi bạn bè giỏi hơn mình. Thay vì dùng sự lo lắng như động lực tự cải thiện, một số học sinh coi sự lo lắng này gây mất ngủ, dẫn đến tình trạng bệnh lý (pathological condition) (12).')

add_p(doc, '"Khi em học trực tuyến, mối quan hệ của em với người bạn thân nhất thực sự rất tệ. Chúng em hiếm khi nói chuyện với nhau, thậm chí cả tuần không nói chuyện. Ngay cả sau khi quay lại trường, em bị cô lập và bị nói xấu bởi các bạn cùng lớp" (Một nữ sinh trường THPT nông thôn).', italic=True, size=11)

add_p(doc, 'Một số học sinh có xung đột với giáo viên: "Từ khi chuyển sang học trực tuyến, cô giáo em chửi mắng học sinh rất nhiều. Đó là lý do em cảm thấy rất áp lực mỗi khi học tiết của cô ấy."')

add_p(doc, 'Nguyên nhân thứ ba là gia đình (family). Xung đột giữa cha mẹ như bất hòa vợ chồng (parental discord), ly hôn, hoặc xung đột trong giao tiếp giữa cha mẹ và con cái.')

add_p(doc, '"Bà nội em là người trọng nam khinh nữ, nên gia đình em thường xuyên cãi nhau trong bữa cơm. Em thích đi học hơn nhưng vì dịch COVID-19, cả gia đình phải ở nhà do giãn cách xã hội (social distancing), nên gần như ngày nào em cũng thấy gia đình cãi nhau. Em cảm thấy rất mệt mỏi" (Một nam sinh trường công lập).', italic=True, size=11)

add_p(doc, 'Kỳ vọng của cha mẹ đối với con cái tạo ra áp lực đáng kể. Một nam sinh trường tư thục chia sẻ: "Em muốn vào trường nghệ thuật nhưng bố mẹ em kỳ vọng em thi sư phạm. Em thực sự không muốn điều đó... Em tiêu cực đến mức không muốn làm gì, ngay cả những việc bình thường như ăn, em cũng không muốn ăn."')

add_heading(doc, 'Chiến lược ứng phó (Coping Strategies)', 3)

add_p(doc, 'Tất cả học sinh đều sử dụng chiến lược ứng phó (coping strategy) để tránh suy sụp. Các chiến lược chính bao gồm:')

add_p(doc, '1) Tự tạo động lực (Self-motivation): tận hưởng hoạt động lành mạnh như chơi thể thao, viết nhật ký (journaling), xem phim, nghe nhạc. "Em từng viết nhật ký, nhưng bố mẹ em lục lọi và đọc rồi mắng em. Bây giờ em buồn hay tiêu cực, em chỉ khóc và khóc, em chỉ muốn ra ngoài và ở một mình cho đến khi em ổn lại" (Một nữ sinh trường công lập).', size=11)

add_p(doc, '2) Chơi thể thao: "Khi em cảm thấy mệt, em thường rủ bạn chơi bóng rổ để thư giãn" hoặc "em nghe nhạc hoặc tham gia các hoạt động ngoài trời." Tuy nhiên, một số HS phản ứng bằng hành vi tiêu cực: "chơi game, điều duy nhất có thể khiến mình vui."', size=11)

add_p(doc, '3) Cần có thời gian biểu hợp lý (reasonable timetable): tổ chức thời gian biểu hợp lý giúp HS giảm lo lắng khi kỳ thi đến.', size=11)

add_p(doc, '4) Cha mẹ quan tâm đến cảm xúc: "Với em, cách tốt nhất để đối phó với lo lắng là sự quan tâm của cha mẹ dành cho con cái. Bất cứ khi nào em cảm thấy khó khăn, em thường tâm sự với mẹ."', size=11)

add_p(doc, '5) Gặp tư vấn viên trường học (meeting school counselors): tuy nhiên, nhiều HS không biết về phòng tư vấn trường học và không biết cách tiếp cận dịch vụ tư vấn.', size=11)

# ============================================================
# PAGE 5-7 — BÀN LUẬN
# ============================================================
page_ref(doc, 5)
add_heading(doc, '4. BÀN LUẬN (Discussion)', 2)

add_p(doc, 'Kết quả cho thấy việc điều tra sức khỏe tâm thần hoặc nhận thức (mental or cognitive health) của học sinh THPT là cần thiết để làm sáng tỏ hệ quả tâm lý chính xác do đại dịch COVID-19 gây ra. Đại dịch tạo ra căng thẳng cá nhân mãnh liệt và ảnh hưởng hệ quả đến sức khỏe tâm thần. Xu hướng ứng phó của học sinh Việt Nam bị ảnh hưởng bởi triệu chứng sức khỏe tâm thần và các yếu tố gây căng thẳng xã hội (social stressors) liên quan đến sự bất định xung quanh đại dịch.')

page_ref(doc, 6)

add_p(doc, 'Các phát hiện cho thấy tỷ lệ hiện mắc đáng kể các triệu chứng rối loạn lo âu trong giai đoạn phong tỏa COVID-19. Một khảo sát trên 5.315 học sinh 11–17 tuổi tại Hà Nội cho thấy 7,4% biểu hiện triệu chứng lo âu nặng, 67,9% từ nhẹ đến trung bình (14). Tại Thái Lan, HS học trực tiếp có khả năng báo cáo lo âu trung bình đến nặng thấp hơn 37,8% so với HS học trực tuyến (15).')

add_p(doc, 'Kết quả của chúng tôi (40,6%) cao hơn so với các nghiên cứu khác: 37,4% tại Trung Quốc (16), 36% tại Mỹ (17), 25% (tháng 6/2021) và 16% (tháng 9/2021) tại Đức (18). Một nghiên cứu trên 212 vị thành niên và 662 người trưởng thành trẻ cho thấy tỷ lệ lo âu tăng từ 24,3% lên 28,4% sau COVID-19 (19). Tỷ lệ cao hơn trong nghiên cứu này có thể do: (1) thời gian giãn cách xã hội và phong tỏa kéo dài tại Hà Nội từ tháng 4/2021, (2) học trực tuyến tại Hà Nội không hiệu quả — chỉ 52,6% HS THPT ưa thích học trực tuyến và 54,3% tự nhận thấy tiến bộ (22).')

add_p(doc, 'Nghiên cứu xác định các nhóm dân số phụ dễ bị tổn thương hơn: nữ giới và phi nhị giới (non-binary) có xu hướng gia tăng báo cáo triệu chứng sức khỏe tâm thần (nữ TB = 1,74 vs nam TB = 1,50, p < 0,01). Cá nhân cư trú tại khu vực thành thị và học sinh lớp 12 cũng biểu hiện xu hướng cao hơn rõ rệt trong việc trải nghiệm triệu chứng lo âu.')

page_ref(doc, 7)

add_p(doc, 'Nghiên cứu cũng chứng minh mối tương quan có ý nghĩa giữa cơ chế ứng phó và mức độ lo âu. 40,6% học sinh báo cáo trải nghiệm triệu chứng rối loạn lo âu từ nhẹ đến nặng. Các cá nhân biểu hiện triệu chứng lo âu có xu hướng sử dụng né tránh (avoidance) như cơ chế ứng phó thường xuyên hơn so với chiến lược thay thế (28, 29).')

add_p(doc, 'Việc thực hiện các chính sách COVID-19 (giãn cách xã hội, cách ly, lối sống ít vận động) đã được liên kết với tác động tiêu cực đến sức khỏe tâm thần. Các phát hiện cho thấy rằng tần suất gia tăng của hoạt động giáo dục trực tiếp, khi thiếu can thiệp giảm thiểu nguyên nhân gốc rễ (root causes) của khổ đau tâm lý, có thể không mang lại tác động có lợi.')

add_heading(doc, 'Hạn chế (Limitations)', 3)
add_p(doc, '(1) Khảo sát không phân biệt giữa giáo dục từ xa bắt buộc và tự nguyện, có thể đưa ra yếu tố gây nhiễu (confounding factors). (2) Sử dụng phương pháp lấy mẫu thuận tiện (convenience sampling) thông qua khảo sát trực tuyến, hạn chế khả năng tổng quát hóa. Tuy nhiên, kích thước mẫu (n = 3.910) đáng kể và đại diện cho một nhóm cụ thể HS THPT tại Hà Nội. (3) GAD-7 là công cụ sàng lọc (screening tool) chứ không phải công cụ chẩn đoán (diagnostic tool), không thể đánh giá tỷ lệ hiện mắc thực sự của rối loạn lo âu.')

add_heading(doc, 'Hàm ý (Implications)', 3)
add_p(doc, '(1) Ban quản lý giáo dục cần hiểu biết về chiến lược ứng phó của học sinh và đảm bảo chăm sóc phù hợp cho những HS không sống cùng cha mẹ. (2) Các cơ sở giáo dục nên triển khai can thiệp như nền tảng trực tuyến chia sẻ trải nghiệm cá nhân. (3) Trường THPT nên ưu tiên tích hợp kỹ thuật ứng phó tiếp cận (approach-coping techniques) vào chương trình giáo dục. (4) Cần tích hợp can thiệp ứng phó (coping interventions) như yếu tố cơ bản của chương trình THPT.')

# ============================================================
# PAGE 8 — KẾT LUẬN
# ============================================================
page_ref(doc, 8)
add_heading(doc, '5. KẾT LUẬN (Conclusions)', 2)

add_p(doc, 'Nghiên cứu chứng minh một mức độ đáng kể của triệu chứng lo âu ở học sinh THPT tại Hà Nội, Việt Nam giữa cuộc khủng hoảng toàn cầu COVID-19. Nghiên cứu xác định các yếu tố liên quan có thể hỗ trợ nhận diện các nhóm dễ bị tổn thương (susceptible groups) và xây dựng can thiệp có mục tiêu. Kết quả nghiên cứu có ý nghĩa quan trọng đối với các khu vực tương tự trong Việt Nam và các quốc gia khác về việc giải quyết tình trạng sức khỏe tâm thần của học sinh trung học.')

add_heading(doc, 'Tuyên bố về tính sẵn có của dữ liệu', 3)
add_p(doc, 'Dữ liệu thô hỗ trợ các kết luận của bài viết này sẽ được các tác giả cung cấp mà không có sự bảo lưu quá mức.', size=11)

add_heading(doc, 'Tuyên bố đạo đức', 3)
add_p(doc, 'Nghiên cứu đã được phê duyệt bởi Hội đồng Đánh giá Thể chế, Trường ĐH KHXH&NV. Sự đồng ý có thông tin bằng văn bản được cung cấp bởi người giám hộ hợp pháp.', size=11)

add_heading(doc, 'Tài trợ', 3)
add_p(doc, 'Không có hỗ trợ tài chính cho việc nghiên cứu, quyền tác giả và/hoặc xuất bản bài viết này.', size=11)

# ============================================================
# TÀI LIỆU THAM KHẢO (rút gọn)
# ============================================================
add_heading(doc, 'TÀI LIỆU THAM KHẢO (References)', 2)
refs = [
    '1. Dahl RE. Adolescent brain development: a period of vulnerabilities and opportunities. Ann New York Acad Sci. 2004;1021:1-22.',
    '2. Patton GC, et al. Our future: a Lancet commission on adolescent health and wellbeing. Lancet. 2016;387:2423-78.',
    '3. Cenat JM, et al. Prevalence of symptoms of depression, anxiety, insomnia... Psychiatry Res. 2021;295:113599.',
    '4. Luthar SS, et al. High-achieving schools connote risks for adolescents. Am Psychol. 2020;75:983-95.',
    '5. Hawrilenko M, et al. The association between school closures and child mental health during COVID-19. JAMA Netw Open. 2021;4:e2124092.',
    '6. Golberstein E, et al. COVID-19 and mental health for children and adolescents. JAMA Pediatr. 2020;174:819-20.',
    '7. Hertz MF, et al. Adolescent mental health, connectedness, and mode of school instruction during COVID-19. J Adolesc Health. 2022;70:57-63.',
    '8. Lessard LM, Puhl RM. Adolescent academic worries amid COVID-19. Sch Psychol. 2021;36:285-92.',
    '9. Nguyen LX, et al. Anxiety and associated factors among Vietnamese students during COVID-19 pandemic. Medicine (Baltimore). 2023;102:e33559.',
    '10. Johnson SU, et al. Psychometric properties of the GAD-7 scale. Front Psychol. 2019;10:1713.',
    '11. Mossman SA, et al. The GAD-7 in adolescents with GAD: Signal detection and validation. Ann Clin Psychiat. 2017;29:227-34a.',
    '12. Staner L. Sleep and anxiety disorders. Dialogues Clin Neurosci. 2003;5:249-58.',
    '13. Hossain SFA, et al. Is M-learning a challenge? Int J e-Collaborat. 2019;15:21-37.',
    '14. Tu HL, et al. Economic, social, and mental health difficulties among high school students in Hanoi during COVID-19 lockdown. J Health Dev Stud. 2023;7:9-19.',
    '15. Widyastari DA, et al. Learning methods during school closure and anxiety of Thai students. Front Pediatr. 2022;10:815148.',
    '16. Zhou S-J, et al. Prevalence and socio-demographic correlates of psychological health problems in Chinese adolescents during COVID-19. Eur Child Adolesc Psychiatry. 2020;29:749-58.',
    '17. Yin O, et al. Persistent anxiety among high school students: second year of COVID. PLoS ONE. 2022;17:e0275292.',
    '18. Theuring S, et al. GAD in Berlin school children after third COVID-19 wave. Child Adolesc Psychiatry Ment Health. 2023;17:1.',
    '19. Villanti AC, et al. COVID-related distress, mental health, and substance use. Child Adolesc Ment Health. 2022;27:138-45.',
    '20. Ho TTQ, et al. Academic stress and depression among Vietnamese adolescents. Curr Psychol. 2023;42:27217-27.',
    '21. Tran BX, et al. Impact of COVID-19 on economic well-being and quality of life of Vietnamese. Front Psychol. 2020;11:565153.',
    '22. Dang TTH, et al. Exploring online learning of school students in Vietnam during COVID-19. Viet J Educ Sci. 2022;18:3.',
    '23. Qi H, et al. Prevalence of anxiety for Chinese adolescents during COVID-19. Psychiatry Clin Neurosci. 2020;74:555-7.',
    '24. Truc TT, et al. Validation of the educational stress scale for adolescents (ESSA) in Vietnam. Asia-Pacific J Public Health. 2015;27:Np2112-21.',
    '25. Ho TTQ, et al. Academic stress and depression among Vietnamese adolescents. Curr Psychol. 2022;2022:1-11.',
    '26. Chakraborty R, Samuels F. Impact of Covid-19 on Adolescent Mental Health in Viet Nam and Tanzania. ODI, 2021.',
    '27. Qiu J, et al. A nationwide survey of psychological distress among Chinese people in the COVID-19 epidemic. General Psychiat. 2020;33:e100213.',
    '28. Naff D, et al. Mental health impacts of COVID-19 on PK-12 students. AERA Open. 2022;8:23328584221084722.',
    '29. Wang L, et al. Self-reported anxiety level in senior high school students in China during COVID-19. J Affect Disord. 2022;301:260-7.',
]
for ref in refs:
    add_p(doc, ref, size=10)

# ============================================================
# BẢNG VIẾT TẮT
# ============================================================
add_abbreviation_table(doc, [
    ('COVID-19', 'Coronavirus Disease 2019 — Bệnh do virus Corona 2019'),
    ('GAD-7', 'Generalized Anxiety Disorder 7-item Scale — Thang đo Rối loạn Lo âu Tổng quát 7 mục'),
    ('THPT', 'Trung học Phổ thông (High School / Senior Secondary School)'),
    ('SD', 'Standard Deviation — Độ lệch chuẩn'),
    ('TB / Mean', 'Trung bình (Mean)'),
    ('SPSS', 'Statistical Package for the Social Sciences — Phần mềm thống kê KHXH'),
    ('IRB', 'Institutional Review Board — Hội đồng Đánh giá Thể chế'),
    ('ĐHQG', 'Đại học Quốc gia (Vietnam National University)'),
    ('ĐH KHXH&NV', 'Đại học Khoa học Xã hội và Nhân văn (University of Social Sciences and Humanities)'),
    ('CC BY', 'Creative Commons Attribution License — Giấy phép Ghi công Creative Commons'),
    ('n', 'Cỡ mẫu (sample size)'),
    ('p', 'Giá trị p (p-value) — mức ý nghĩa thống kê'),
    ('α', 'Cronbach\'s alpha — hệ số tin cậy nhất quán nội tại'),
])

# ============================================================
# PHẦN PHÊ BÌNH (RED)
# ============================================================
doc.add_paragraph()
add_red_heading(doc, 'PHẦN PHÊ BÌNH VÀ ĐÁNH GIÁ (Critical Appraisal)')

# Strengths
add_red_heading(doc, 'A. Điểm mạnh (Strengths)')
add_red(doc, '1. Cỡ mẫu lớn (n = 3.910) từ 13 trường THPT đại diện cho cả 4 nhóm phân tầng (thành thị/nông thôn × công lập/tư thục), đảm bảo tính đại diện cao cho bối cảnh Hà Nội.')
add_red(doc, '2. Sử dụng phương pháp hỗn hợp (mixed methods) kết hợp dữ liệu định lượng (GAD-7) và định tính (phỏng vấn sâu 20 HS), giúp hiểu sâu hơn về nguyên nhân và cơ chế ứng phó, không chỉ dừng ở con số thống kê.')
add_red(doc, '3. Công cụ GAD-7 có độ tin cậy cao (Cronbach\'s alpha = 0,916), vượt ngưỡng chấp nhận (0,7), và đã được xác nhận hiệu lực với quần thể vị thành niên quốc tế.')
add_red(doc, '4. Nghiên cứu được thực hiện đúng thời điểm quan trọng (tháng 10–11/2021) khi Hà Nội đang phong tỏa kéo dài, cung cấp dữ liệu thực tế về tác động tâm lý tức thời của đại dịch lên học sinh.')

# Limitations
add_red_heading(doc, 'B. Hạn chế (Limitations)')
add_red(doc, '1. Thiết kế cắt ngang (cross-sectional): không thể xác lập quan hệ nhân quả giữa đại dịch COVID-19 và triệu chứng lo âu, chỉ mô tả tương quan tại một thời điểm.')
add_red(doc, '2. Không có nhóm đối chứng hoặc dữ liệu baseline trước đại dịch: không thể so sánh mức lo âu trước và sau COVID-19 ở cùng nhóm đối tượng, do đó khó kết luận tỷ lệ 40,6% là "tăng" hay "bình thường."')
add_red(doc, '3. Lấy mẫu thuận tiện qua khảo sát trực tuyến (Google Form): học sinh không có thiết bị/internet hoặc không quen sử dụng công nghệ bị loại trừ — chính nhóm này có thể có mức lo âu cao hơn do bất bình đẳng kỹ thuật số (digital divide).')
add_red(doc, '4. Phân tích định tính chỉ với 20 HS phỏng vấn — cỡ mẫu nhỏ, không rõ tiêu chí chọn mẫu mục đích (purposive sampling) hay ngẫu nhiên, và không báo cáo chi tiết khi nào đạt bão hòa (saturation).')
add_red(doc, '5. GAD-7 là công cụ sàng lọc, không phải chẩn đoán lâm sàng. Tỷ lệ 40,6% phản ánh "triệu chứng lo âu" chứ không phải "rối loạn lo âu" được chẩn đoán (diagnosed anxiety disorder). Bài báo đôi khi không phân biệt rõ hai khái niệm này.')
add_red(doc, '6. Không kiểm soát các biến gây nhiễu quan trọng: tình trạng kinh tế gia đình, tiền sử bệnh tâm thần cá nhân/gia đình, mức độ tiếp cận dịch vụ y tế/tâm lý, và thời lượng sử dụng màn hình/mạng xã hội.')

# Research gaps
add_red_heading(doc, 'C. Khoảng trống nghiên cứu có thể khai thác (Research Gaps)')
add_red(doc, '1. Nghiên cứu dọc (longitudinal study) theo dõi diễn biến triệu chứng lo âu từ giai đoạn phong tỏa → mở cửa trường học → giai đoạn hậu đại dịch, để hiểu quỹ đạo phục hồi (recovery trajectory) của HS THPT Việt Nam.')
add_red(doc, '2. So sánh hiệu quả các mô hình can thiệp tâm lý học đường (school-based mental health interventions): tư vấn trực tiếp vs. nền tảng trực tuyến vs. chương trình peer support — đặc biệt cho bối cảnh Việt Nam nơi dịch vụ SKTT học đường còn hạn chế.')
add_red(doc, '3. Phân tích mối quan hệ giữa chiến lược ứng phó cụ thể (tiếp cận vs. né tránh) và mức độ lo âu bằng mô hình hồi quy/SEM, thay vì chỉ mô tả định tính. Cần công cụ đo chiến lược ứng phó có chuẩn hóa (ví dụ: Brief COPE, CERQ).')
add_red(doc, '4. Nghiên cứu vai trò trung gian/điều tiết của các yếu tố bảo vệ (protective factors): hỗ trợ xã hội, mối quan hệ gia đình, khả năng phục hồi (resilience), và vốn tâm lý (psychological capital) trong mối quan hệ giữa căng thẳng đại dịch và lo âu.')
add_red(doc, '5. Mở rộng nghiên cứu sang các tỉnh/thành khác ngoài Hà Nội, đặc biệt vùng nông thôn và vùng sâu vùng xa, nơi học trực tuyến gặp nhiều rào cản hơn và tiếp cận dịch vụ SKTT còn hạn chế hơn.')

# ============================================================
# SAVE
# ============================================================
output = os.path.join(os.path.dirname(os.path.abspath(__file__)), '02_Hoa_2024_Frontiers.docx')
doc.save(output)
print(f'DONE — Saved to: {output}')
