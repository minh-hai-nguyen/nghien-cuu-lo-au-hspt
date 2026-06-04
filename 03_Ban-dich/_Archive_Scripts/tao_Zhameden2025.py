# -*- coding: utf-8 -*-
"""
Generate DOCX for Zhameden 2025 (PLOS ONE) Vietnamese translation
School-based interventions to prevent anxiety and depression in children
and adolescents in LMICs: A systematic review
"""
import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from tao_dich_template import (
    create_doc, add_link_and_qr, add_heading, add_page_ref,
    add_p, add_red, add_red_heading, add_table, add_info_table,
    add_image, add_abbreviation_table
)

JOURNAL = 'PLOS ONE'
VOL = '20(4): e0316825'
URL = 'https://doi.org/10.1371/journal.pone.0316825'
QR = 'QR_Zhameden2025.png'
OUTPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                      '04_Zhameden_2025_PLOSONE.docx')

doc = create_doc()

# ===== LINK + QR =====
add_link_and_qr(doc, URL, QR)

# ===== TITLE =====
add_heading(doc, 'BÀI BÁO: Can thiệp tại trường học phòng ngừa lo âu và trầm cảm ở trẻ em và VTN tại LMIC', level=1)

# ===== Info table =====
add_info_table(doc, [
    ('Tên gốc', 'School-based interventions to prevent anxiety and depression in children and adolescents in low- and middle-income countries: A systematic review'),
    ('Tác giả', 'Sharone Zhameden Dieu Yin, Mei Ken Low, Masuma Pervin Mishu'),
    ('Cơ quan', 'Institute of Epidemiology and Health Care & Institute of Ophthalmology, University College London, UK'),
    ('Tạp chí', f'{JOURNAL}, {VOL}'),
    ('DOI', URL),
    ('Nhận bài', '1/6/2024'),
    ('Chấp nhận', '17/12/2024'),
    ('Xuất bản', '30/4/2025'),
    ('Thiết kế', 'Tổng quan hệ thống (Systematic Review)'),
    ('Cỡ mẫu', '6 RCT, 1.587 học sinh'),
    ('Quốc gia', 'Brazil, Trung Quốc, Ấn Độ, Kenya, Lebanon, Malaysia'),
])

doc.add_paragraph()

# =====================================================================
# PAGE 1 — Abstract
# =====================================================================
add_page_ref(doc, '1/21', JOURNAL, VOL)
add_heading(doc, 'TÓM TẮT (Abstract)', level=2)

add_p(doc, 'Bối cảnh (Background)', bold=True)
add_p(doc, 'Lo âu (anxiety) và trầm cảm (depression) đang gia tăng ở trẻ em và vị thành niên trên toàn cầu. Các nước thu nhập thấp và trung bình (low- and middle-income countries — LMICs) đối mặt với mức độ dễ bị tổn thương cao hơn do nguồn lực hạn chế, tiếp cận dịch vụ sức khỏe tâm thần bị hạn chế, bất bình đẳng kinh tế xã hội và sự kỳ thị sức khỏe tâm thần phổ biến. Trường học cung cấp một môi trường độc đáo và có tiềm năng tác động lớn cho các can thiệp phòng ngừa nhắm vào lo âu và trầm cảm ở giới trẻ.')

add_p(doc, 'Mục tiêu (Aim)', bold=True)
add_p(doc, 'Chúng tôi nhằm xác định các nghiên cứu thực nghiệm để khám phá hiệu quả của các can thiệp tại trường học được thiết kế nhằm phòng ngừa lo âu và trầm cảm ở trẻ em và vị thành niên tại LMICs.')

# =====================================================================
# PAGE 2 — Abstract cont + Intro
# =====================================================================
add_page_ref(doc, '2/21', JOURNAL, VOL)

add_p(doc, 'Phương pháp (Method)', bold=True)
add_p(doc, 'Ovid MEDLINE, Embase, PsycINFO và CENTRAL được tìm kiếm một cách có hệ thống cho các bài báo xuất bản từ năm 2018 đến tháng 7/2023. Các thử nghiệm ngẫu nhiên có đối chứng (randomised controlled trials — RCTs) đánh giá can thiệp tại trường học cho trẻ em và vị thành niên từ 4–18 tuổi tại LMICs được đưa vào. Chỉ các nghiên cứu bằng tiếng Anh được đưa vào. Kết quả chính (primary outcomes) là triệu chứng lo âu và/hoặc trầm cảm. Đánh giá nguy cơ thiên vị (risk-of-bias assessments) đã được thực hiện.')

add_p(doc, 'Kết quả (Results)', bold=True)
add_p(doc, 'Từ 3.863 bài báo được xác định, sáu nghiên cứu bao gồm 1.587 học sinh đáp ứng tiêu chí đưa vào. Trong bốn nghiên cứu đánh giá can thiệp phòng ngừa cả lo âu và trầm cảm, cũng như chỉ lo âu, chỉ một nghiên cứu cho thấy sự giảm triệu chứng lo âu. Trong trường hợp trầm cảm, ba trong bốn nghiên cứu báo cáo cải thiện triệu chứng trầm cảm. Phát hiện này gợi ý hiệu quả tiềm năng của các can thiệp phòng ngừa đối với trầm cảm, nhưng không phải lo âu. Tuy nhiên, phát hiện này cần được diễn giải thận trọng do số lượng nghiên cứu hạn chế. Tất cả các nghiên cứu đều được phân loại có nguy cơ thiên vị cao hoặc có một số lo ngại.')

add_p(doc, 'Kết luận (Conclusion)', bold=True)
add_p(doc, 'Có một số bằng chứng về hiệu quả của can thiệp tại trường học trong việc phòng ngừa lo âu và trầm cảm ở giới trẻ tại LMICs. Cần nghiên cứu thêm để có hiểu biết toàn diện hơn về vấn đề quan trọng này. Trong tương lai, điều thiết yếu là tăng cường và mở rộng các chương trình phòng ngừa tại trường học hiện có tại các quốc gia này, khám phá các chiến lược can thiệp khác nhau phù hợp với các yếu tố bối cảnh cụ thể của họ.')

# =====================================================================
# PAGE 3 — Introduction
# =====================================================================
add_page_ref(doc, '3/21', JOURNAL, VOL)
add_heading(doc, 'GIỚI THIỆU (Introduction)', level=2)

add_p(doc, 'Các vấn đề sức khỏe tâm thần ở trẻ em và vị thành niên là mối quan tâm y tế công cộng lớn trên toàn cầu. Theo Báo cáo Sức khỏe Tâm thần Thế giới (World Mental Health Report) của Tổ chức Y tế Thế giới (WHO) năm 2022, khoảng 8% trẻ nhỏ từ 5–9 tuổi và 14% vị thành niên từ 10–19 tuổi bị ảnh hưởng bởi các rối loạn tâm thần trên toàn cầu [1]. Các rối loạn sức khỏe tâm thần phổ biến nhất là lo âu và trầm cảm, chiếm hơn 40% các trường hợp ở vị thành niên [2]. Kể từ khi bắt đầu đại dịch bệnh coronavirus 2019 (COVID-19), tỷ lệ ước tính các triệu chứng lo âu và trầm cảm ở trẻ em và vị thành niên đã tăng gấp đôi. Một phân tích tổng hợp (meta-analysis) gồm 29 nghiên cứu cho thấy tỷ lệ hiện mắc toàn cầu của triệu chứng lo âu và trầm cảm tăng cao ở mức lâm sàng trong nhóm tuổi này trong năm đầu tiên của đại dịch lần lượt là 20,5% và 25,2% [3]. Đại dịch đã làm trầm trọng thêm các thách thức sức khỏe tâm thần hiện có và đưa vào các yếu tố gây căng thẳng mới cho nhóm dân số này.')

add_p(doc, 'Các yếu tố như sự cô lập, gián đoạn thói quen và giáo dục, áp lực tài chính và tăng cảm giác bất an về tương lai đã đóng góp đáng kể vào mức độ rối nhiễu tâm lý (psychological distress) gia tăng [4–6].')

# =====================================================================
# PAGE 4 — Introduction cont
# =====================================================================
add_page_ref(doc, '4/21', JOURNAL, VOL)

add_p(doc, 'Theo phân tích có hệ thống gần đây nhất về gánh nặng bệnh tật toàn cầu ở giới trẻ từ 10–24 tuổi, tự gây hại (self-harm), một hệ quả nghiêm trọng liên quan mạnh đến sức khỏe tâm thần kém, đã nổi lên là một trong những nguyên nhân hàng đầu gây tử vong ở vị thành niên. Nó đóng góp 8,2% tổng số ca tử vong ở vị thành niên toàn cầu và khoảng 20% ca tử vong trong nhóm tuổi 15–24 [7]. Phần lớn các rối loạn sức khỏe tâm thần bắt đầu trong giai đoạn vị thành niên, với một nửa bắt đầu trước 14 tuổi và ba phần tư trước 25 tuổi [8,9]. Nghiên cứu cho thấy các vấn đề sức khỏe tâm thần trải qua trong thời thơ ấu và vị thành niên có thể kéo dài hoặc tái phát trong suốt cuộc đời [10–12] và có những hệ quả sinh-tâm-xã hội (biopsychosocial) dài hạn [13–15]. Hơn nữa, các rối loạn tâm thần ở trẻ em có tác động kinh tế đáng kể [16–18]. Theo báo cáo Tình trạng Trẻ em Thế giới 2021 của UNICEF, các vấn đề sức khỏe tâm thần ở trẻ em 0–19 tuổi dẫn đến tổn thất ước tính hàng năm là 340,2 tỷ USD, điều chỉnh theo sức mua tương đương, trong vốn nhân lực [2].')

add_p(doc, 'Gánh nặng các vấn đề sức khỏe tâm thần được phân bổ không đều, với tỷ lệ hiện mắc cao hơn ở các nước thu nhập thấp và trung bình (LMICs). Ước tính hơn 225 triệu trẻ em và vị thành niên trên toàn thế giới bị ảnh hưởng bởi các rối loạn tâm thần và 88% trong số này được tìm thấy ở LMICs [19]. Tuy nhiên, dữ liệu về sức khỏe tâm thần, đặc biệt cho trẻ em và vị thành niên, tại các khu vực này còn hạn chế. Mặc dù LMICs là nơi sinh sống của gần 90% dân số vị thành niên toàn cầu, phạm vi bao phủ dữ liệu chỉ có sẵn cho khoảng 2% trẻ em và vị thành niên sống tại các quốc gia này [1,2,20].')

# =====================================================================
# PAGE 5 — Introduction cont (LMIC challenges, prevention)
# =====================================================================
add_page_ref(doc, '5/21', JOURNAL, VOL)

add_p(doc, 'Trẻ em và vị thành niên tại LMICs đối mặt với những thách thức riêng biệt trong việc giải quyết các vấn đề sức khỏe tâm thần và cung cấp dịch vụ chăm sóc sức khỏe tâm thần. Những thách thức này bao gồm nguồn lực con người và tài chính hạn chế, cơ sở hạ tầng sức khỏe tâm thần không đầy đủ và sự khác biệt lớn trong bối cảnh văn hóa-xã hội [21]. Các hệ thống y tế thường thể hiện sự tích hợp không đầy đủ các dịch vụ sức khỏe tâm thần vào các chính sách chăm sóc sức khỏe rộng hơn hoặc trong chăm sóc sức khỏe ban đầu [22–24]. Ngoài ra, nhóm dân số này đối mặt với nguy cơ cao hơn bị phơi nhiễm với các trải nghiệm bất lợi thời thơ ấu (adverse childhood experiences) và các yếu tố nguy cơ khác nhau như lạm dụng thể chất và cảm xúc, bỏ bê mãn tính, nghèo đói, vô gia cư, bạo lực và xung đột [25–30]. Hơn nữa, văn hóa kỳ thị xung quanh sức khỏe tâm thần phổ biến ở LMICs. Những quan niệm sai lầm, sợ hãi và thái độ xã hội về bệnh tâm thần thường dẫn đến loại trừ xã hội, phân biệt đối xử và do dự trong việc tìm kiếm sự hỗ trợ [31–33].')

add_p(doc, 'Xét đến những hệ quả dài hạn của các tình trạng sức khỏe tâm thần không được điều trị, có nhu cầu cấp thiết về các biện pháp phòng ngừa và can thiệp kịp thời. Các chiến lược phòng ngừa sớm là rất quan trọng trong việc giảm tỷ lệ hiện mắc các rối loạn sức khỏe tâm thần ở trẻ em và giới trẻ. Phòng ngừa sơ cấp (primary prevention) đề cập đến các biện pháp được thiết kế để ngăn ngừa sự xuất hiện của bệnh trước khi nó phát triển trong một quần thể hoặc cá nhân có nguy cơ [35]. Các can thiệp nhằm phòng ngừa sơ cấp có thể được phân loại thành phổ quát (universal), chọn lọc (selective) và chỉ định (indicated). Phòng ngừa phổ quát nhằm tiếp cận toàn bộ quần thể không xác định trên cơ sở nguy cơ. Phòng ngừa chọn lọc tập trung vào các cá nhân có nguy cơ cao hơn trung bình. Phòng ngừa chỉ định nhắm vào các cá nhân có nguy cơ cao thể hiện các triệu chứng dưới lâm sàng sớm [36].')

add_p(doc, 'Các chính sách đã ưu tiên phòng ngừa sơ cấp các rối loạn tâm thần ở trẻ em và giới trẻ, với trường học đóng vai trò trung tâm trong việc thực hiện [37]. Trường học là môi trường quan trọng cho học tập xã hội và cảm xúc [38]. Vì giáo dục là bắt buộc ở hầu hết các LMICs [39,40], phần lớn trẻ em và giới trẻ dành phần lớn thời gian tại trường. Do đó, can thiệp tại trường học có tiềm năng tiếp cận một lượng lớn trẻ em và vị thành niên, bất kể nền tảng văn hóa-xã hội. Chúng có thể giúp vượt qua các thách thức liên quan đến nguồn lực hạn chế và tiếp cận dịch vụ sức khỏe tâm thần kém.')

# =====================================================================
# PAGE 6 — Intro cont + Methods start, PRISMA flowchart
# =====================================================================
add_page_ref(doc, '6/21', JOURNAL, VOL)

add_p(doc, 'Trong số các can thiệp, liệu pháp hành vi nhận thức (cognitive-behavioural therapy — CBT) là một trong những can thiệp tâm lý được sử dụng rộng rãi nhất, được hỗ trợ bởi bằng chứng mạnh mẽ chứng minh hiệu quả của nó trong việc giải quyết cả lo âu và trầm cảm [42].')

add_p(doc, 'Một số tổng quan hệ thống đã tìm cách đánh giá hiệu quả của can thiệp phòng ngừa tại trường học trong việc giải quyết lo âu và trầm cảm ở trẻ em và vị thành niên [41,43–46]. Tuy nhiên, phần lớn các can thiệp đã được đánh giá tại các nước thu nhập cao (HICs). Xét đến những thách thức và bất bình đẳng sức khỏe tâm thần riêng biệt tại LMICs, vẫn chưa chắc chắn liệu các can thiệp này có thể được áp dụng hiệu quả trong bối cảnh như vậy. Tổng quan hệ thống này nhằm xác định và khám phá hiệu quả của các can thiệp phòng ngừa lo âu và trầm cảm ở trẻ em và vị thành niên tại LMICs.')

# =====================================================================
# PAGE 7 — Methods
# =====================================================================
add_page_ref(doc, '7/21', JOURNAL, VOL)
add_heading(doc, 'PHƯƠNG PHÁP (Methods)', level=2)

add_p(doc, 'Tổng quan hệ thống này được tiến hành và báo cáo theo hướng dẫn PRISMA (Preferred Reporting Items for Systematic Reviews and Meta-Analyses) [47]. Giao thức (protocol) được đăng ký hồi cứu trên INPLASY: INPLASY2024100063 vào ngày 15 tháng 10, 2024.')

add_p(doc, 'Tiêu chí đưa vào (Inclusion criteria)', bold=True)
add_p(doc, 'Các nghiên cứu với người tham gia từ 4–18 tuổi tại thời điểm tuyển chọn và không có tình trạng sức khỏe thể chất hoặc tâm thần có thể xác định được đưa vào. Các nghiên cứu đủ điều kiện nếu chúng đánh giá can thiệp tại trường học nhằm phòng ngừa lo âu và trầm cảm. Các can thiệp này phải diễn ra trong môi trường trường học hoặc được tích hợp vào chương trình giảng dạy. Các phương pháp phòng ngừa sơ cấp khác nhau, bao gồm phòng ngừa phổ quát, chọn lọc hoặc chỉ định, đều đủ điều kiện. Các thử nghiệm ngẫu nhiên có đối chứng (RCTs) bao gồm cả RCT cá nhân và RCT theo cụm báo cáo triệu chứng lo âu và/hoặc trầm cảm làm kết quả được đưa vào. Chỉ các nghiên cứu được tiến hành tại LMICs, dựa trên danh sách DAC về các nước nhận viện trợ phát triển chính thức (ODA), được đưa vào. Các nghiên cứu xuất bản từ năm 2018 trở đi được đưa vào. Chỉ các nghiên cứu bằng tiếng Anh được đưa vào.')

add_p(doc, 'Tiêu chí loại trừ (Exclusion criteria)', bold=True)
add_p(doc, 'Các nghiên cứu tập trung vào nâng cao nhận thức sức khỏe tâm thần, sức khỏe cảm xúc (emotional well-being) và tâm lý học tích cực (positive psychology) không đủ điều kiện, trừ khi mục tiêu là phòng ngừa lo âu và trầm cảm. Các can thiệp nhằm giải quyết các vấn đề có thể dẫn đến rối loạn tâm thần (stress, bắt nạt, lạm dụng chất) bị loại trừ. Tương tự, các can thiệp nhằm giúp trẻ em và giới trẻ đối phó với các sự kiện hoặc hoàn cảnh cụ thể như ly hôn của cha mẹ, thiên tai và xung đột, bị loại trừ.')

# =====================================================================
# PAGE 8 — Search strategy, Data extraction
# =====================================================================
add_page_ref(doc, '8/21', JOURNAL, VOL)

add_p(doc, 'Chiến lược tìm kiếm và lựa chọn nghiên cứu (Search strategy and study selection)', bold=True)
add_p(doc, 'Ovid MEDLINE, Embase, PsycINFO và CENTRAL được tìm kiếm vào ngày 7 tháng 7, 2023. Các thuật ngữ tìm kiếm được phát triển sử dụng chiến lược tìm kiếm của Caldwell và cộng sự [41] và được mở rộng thêm để bao gồm LMICs được liệt kê trong danh sách DAC [48].')

add_p(doc, 'Trích xuất dữ liệu (Data extraction)', bold=True)
add_p(doc, 'Một mẫu trích xuất dữ liệu tiêu chuẩn hóa được sử dụng để trích xuất dữ liệu từ các nghiên cứu được đưa vào, bao gồm: tác giả, năm xuất bản, quốc gia, thiết kế nghiên cứu, cỡ mẫu, đặc điểm dân số (tuổi, giới tính), loại can thiệp và nhóm đối chứng, kết quả chính, công cụ đo lường, và kết quả.')

add_p(doc, 'Đánh giá nguy cơ thiên vị (Risk of bias assessment)', bold=True)
add_p(doc, 'Đánh giá nguy cơ thiên vị được thực hiện sử dụng công cụ Cochrane Risk of Bias 2 (RoB 2) cho các RCT cá nhân [50] và phiên bản RoB 2 cho các thử nghiệm ngẫu nhiên theo cụm [51].')

add_p(doc, 'Đánh giá chất lượng bằng chứng (Quality of evidence)', bold=True)
add_p(doc, 'Phương pháp GRADE (Grading of Recommendations, Assessment, Development and Evaluations) được sử dụng để đánh giá chất lượng bằng chứng [52].')

# =====================================================================
# PAGE 9 — PRISMA flowchart + Results: Study selection
# =====================================================================
add_page_ref(doc, '9/21', JOURNAL, VOL)
add_heading(doc, 'KẾT QUẢ (Results)', level=2)

add_p(doc, 'Lựa chọn nghiên cứu (Study selection)', bold=True)
add_p(doc, 'Tổng cộng 3.863 bài báo được xác định. Sau khi loại bỏ trùng lặp (n = 1.232), 2.631 bài được sàng lọc tiêu đề và tóm tắt. 34 bài toàn văn được đánh giá. Cuối cùng, sáu nghiên cứu đáp ứng tiêu chí đưa vào.')

add_image(doc, 'Zhameden2025_fig1_prisma.png', width_cm=14,
          caption='Hình 1. Sơ đồ PRISMA — Quy trình lựa chọn nghiên cứu (Fig 1. PRISMA flow diagram)')
doc.add_paragraph()

# =====================================================================
# PAGE 10 — Study characteristics table
# =====================================================================
add_page_ref(doc, '10/21', JOURNAL, VOL)

add_p(doc, 'Đặc điểm nghiên cứu (Study characteristics)', bold=True)
add_p(doc, 'Sáu nghiên cứu bao gồm tổng cộng 1.587 người tham gia. Các nghiên cứu được tiến hành tại: Brazil (thu nhập trung bình cao), Trung Quốc (thu nhập trung bình cao), Ấn Độ (thu nhập trung bình thấp), Kenya (thu nhập trung bình thấp), Lebanon (thu nhập trung bình cao) và Malaysia (thu nhập trung bình cao). Không có nghiên cứu nào được tiến hành tại các nước thu nhập thấp.')

# Study characteristics table
add_heading(doc, 'Bảng 1. Đặc điểm các nghiên cứu được đưa vào', level=3)
study_headers = ['Tác giả (Năm)', 'Quốc gia', 'Thiết kế', 'N', 'Tuổi', 'Can thiệp', 'Kết quả chính']
study_rows = [
    ['Ab Ghaffar và cs. (2019)', 'Malaysia', 'RCT', '58', '16–17', 'CBT trực tuyến (e-CBT)', 'Lo âu (DASS-21)'],
    ['Bella-Awusah và cs. (2021)', 'Kenya', 'Cluster RCT', '231', '11–14', 'CBT nhóm', 'Lo âu + Trầm cảm'],
    ['Desan và cs. (2022)', 'Lebanon', 'Cluster RCT thí điểm', '580', '10–14', 'Giáo dục tâm lý tích cực', 'Trầm cảm (CES-D)'],
    ['Li và cs. (2019)', 'Trung Quốc', 'Cluster RCT', '345', '13–15', 'CBT tại trường', 'Lo âu + Trầm cảm'],
    ['Rivero và cs. (2022)', 'Brazil', 'RCT', '253', '8–12', 'FRIENDS phổ quát', 'Lo âu + Trầm cảm'],
    ['Singhal và cs. (2018)', 'Ấn Độ', 'RCT', '120', '13–18', 'CBT chỉ định (Coping Skills)', 'Trầm cảm (CDI, CES-DC)'],
]
add_table(doc, study_headers, study_rows, widths=[3.0, 1.8, 2.0, 1.0, 1.2, 3.0, 3.0])
doc.add_paragraph()

# =====================================================================
# PAGE 11 — Results: Anxiety + Depression outcomes
# =====================================================================
add_page_ref(doc, '11/21', JOURNAL, VOL)

add_heading(doc, 'Kết quả về Lo âu (Anxiety outcomes)', level=3)
add_p(doc, 'Trong bốn nghiên cứu đánh giá can thiệp cho lo âu:')

add_p(doc, 'Ab Ghaffar và cộng sự (Malaysia): Can thiệp CBT trực tuyến cho 58 học sinh 16–17 tuổi. Kết quả: giảm nhỏ nhưng có ý nghĩa thống kê trong triệu chứng lo âu (nhóm can thiệp: trước M=7.48, SD=4.45; sau M=5.79, SD=4.08; p<0.05, d=0.25).')
add_p(doc, 'Rivero và cộng sự (Brazil): Chương trình phòng ngừa phổ quát FRIENDS cho 253 học sinh 8–12 tuổi. Kết quả: giảm triệu chứng lo âu ở cả hai nhóm nhưng sự khác biệt không đạt ý nghĩa thống kê (ITT1: F=0.14, p=0.86; ITT2: F=0.70, p=0.47).')

add_p(doc, 'Tóm tắt: Cả hai nghiên cứu quan sát thấy sự giảm triệu chứng lo âu. Tuy nhiên, chỉ kết quả của Ab Ghaffar và cộng sự đạt ý nghĩa thống kê.', bold=True)

add_heading(doc, 'Kết quả về Trầm cảm (Depression outcomes)', level=3)
add_p(doc, 'Hai RCT đánh giá can thiệp tập trung riêng vào trầm cảm:')
add_p(doc, 'Desan và cộng sự (Lebanon): RCT thí điểm theo cụm, 580 học sinh. Chương trình giáo dục tâm lý tích cực. Cả hai nhóm đều tăng điểm CES-D. Hiệu ứng can thiệp không có ý nghĩa thống kê (M=-0.73, SE=0.44, p>0.05, 95% CI [-1.58, 0.13]).')
add_p(doc, 'Singhal và cộng sự (Ấn Độ): Can thiệp CBT chỉ định (Chương trình Kỹ năng Ứng phó — Coping Skills), 120 học sinh 13–18 tuổi có trầm cảm dưới lâm sàng. Kết quả: giảm có ý nghĩa thống kê ở cả CDI (F(1,90)=234.2, p<0.001) và CES-DC (F(1,90)=132.5, p<0.001).')

# =====================================================================
# PAGE 12 — Risk of bias figure
# =====================================================================
add_page_ref(doc, '12/21', JOURNAL, VOL)

add_heading(doc, 'Kết quả kết hợp Lo âu và Trầm cảm', level=3)
add_p(doc, 'Bốn nghiên cứu đánh giá can thiệp cho cả lo âu và trầm cảm: Bella-Awusah và cs. (Kenya), Li và cs. (Trung Quốc), Ab Ghaffar và cs. (Malaysia), Rivero và cs. (Brazil). Kết quả: 3/4 nghiên cứu cho thấy cải thiện triệu chứng trầm cảm, nhưng chỉ 1/4 cho thấy giảm lo âu có ý nghĩa.')

add_heading(doc, 'Đánh giá nguy cơ thiên vị (Risk of bias)', level=3)
add_p(doc, 'Bốn nghiên cứu được đánh giá có nguy cơ thiên vị cao (high risk of bias) và hai nghiên cứu có một số lo ngại (some concerns). Các lý do chính bao gồm: thiếu mù hóa (blinding) trong đánh giá kết quả, thiếu bằng chứng về dữ liệu thiếu (missing data), sai lệch đáng kể so với can thiệp dự định, và vấn đề với quy trình ngẫu nhiên hóa.')

add_image(doc, 'Zhameden2025_fig2_rob.png', width_cm=14,
          caption='Hình 2. Đánh giá nguy cơ thiên vị — Risk of bias summary (Fig 2)')
doc.add_paragraph()

# =====================================================================
# PAGE 13 — Forest plots (Fig 3 & 4)
# =====================================================================
add_page_ref(doc, '13/21', JOURNAL, VOL)

add_image(doc, 'Zhameden2025_fig3_4.png', width_cm=14,
          caption='Hình 3. Forest plot — Hiệu quả can thiệp đối với lo âu (Fig 3)')
doc.add_paragraph()

add_image(doc, 'Zhameden2025_fig3_4.png', width_cm=14,
          caption='Hình 4. Forest plot — Hiệu quả can thiệp đối với trầm cảm (Fig 4)')
doc.add_paragraph()

# =====================================================================
# PAGE 14 — GRADE assessment
# =====================================================================
add_page_ref(doc, '14/21', JOURNAL, VOL)

add_heading(doc, 'Đánh giá chất lượng bằng chứng GRADE', level=3)
add_p(doc, 'Theo đánh giá GRADE, chất lượng bằng chứng tổng thể cho tất cả các lĩnh vực kết quả được đánh giá là rất thấp (very low). Một số lo ngại về chất lượng phát sinh từ nguy cơ thiên vị, tính gián tiếp (indirectness) và tính không chính xác (imprecision).')

add_heading(doc, 'Bảng 2. Đánh giá chất lượng bằng chứng GRADE', level=3)
grade_headers = ['Kết quả', 'Số NC', 'Nguy cơ thiên vị', 'Không nhất quán', 'Gián tiếp', 'Không chính xác', 'Chất lượng']
grade_rows = [
    ['Lo âu', '4', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng', 'Rất thấp'],
    ['Trầm cảm', '4', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng', 'Rất thấp'],
]
add_table(doc, grade_headers, grade_rows, widths=[2.0, 1.2, 2.5, 2.5, 2.0, 2.5, 2.0])
doc.add_paragraph()

# =====================================================================
# PAGE 15–16 — Discussion: Anxiety
# =====================================================================
add_page_ref(doc, '15/21', JOURNAL, VOL)
add_heading(doc, 'THẢO LUẬN (Discussion)', level=2)

add_p(doc, 'Tổng cộng, sáu nghiên cứu liên quan về can thiệp tại trường học phòng ngừa lo âu và trầm cảm ở trẻ em và vị thành niên tại LMICs đã được xác định. Các phát hiện gợi ý hiệu quả tiềm năng của can thiệp phòng ngừa đối với trầm cảm, trong khi điều tương tự không thể kết luận cho lo âu.')

add_p(doc, 'Can thiệp cho lo âu', bold=True)
add_p(doc, 'Trong số các nghiên cứu đánh giá can thiệp cho lo âu, chỉ một nghiên cứu cho thấy giảm có ý nghĩa thống kê. Kích thước hiệu ứng (effect size) nhỏ (d=0.25) gợi ý rằng mặc dù can thiệp có hiệu quả, tác động tổng thể lên triệu chứng lo âu là khiêm tốn. Phát hiện này phù hợp với các tổng quan trước đây cho thấy can thiệp tại trường học nhìn chung có hiệu quả khiêm tốn trong việc giảm triệu chứng lo âu.')

add_page_ref(doc, '16/21', JOURNAL, VOL)

add_p(doc, 'Một lý do khả thi cho hiệu quả hạn chế đối với lo âu so với trầm cảm có thể liên quan đến bản chất của các rối loạn. Lo âu thường bao gồm các phản ứng sinh lý (physiological responses) ngoài các thành phần nhận thức và hành vi. CBT, mặc dù hiệu quả cho các thành phần nhận thức, có thể ít tác động hơn đến các phản ứng sinh lý liên quan đến lo âu.')

# =====================================================================
# PAGE 17 — Discussion: Depression
# =====================================================================
add_page_ref(doc, '17/21', JOURNAL, VOL)

add_p(doc, 'Can thiệp cho trầm cảm', bold=True)
add_p(doc, 'Kết quả cho trầm cảm khả quan hơn, với 3/4 nghiên cứu cho thấy cải thiện. Đặc biệt, nghiên cứu của Singhal và cộng sự sử dụng can thiệp chỉ định (indicated prevention) — nhắm vào học sinh đã có triệu chứng dưới lâm sàng — cho thấy kết quả mạnh nhất. Điều này gợi ý rằng can thiệp chỉ định có thể hiệu quả hơn so với can thiệp phổ quát trong bối cảnh LMIC.')

# =====================================================================
# PAGE 18 — Discussion: Gaps
# =====================================================================
add_page_ref(doc, '18/21', JOURNAL, VOL)

add_p(doc, 'Khoảng trống nghiên cứu', bold=True)
add_p(doc, 'Một phát hiện đáng chú ý là không có nghiên cứu nào từ các nước thu nhập thấp (low-income countries). Tất cả sáu nghiên cứu đều từ các nước thu nhập trung bình. Điều này tạo ra khoảng trống bằng chứng đáng kể, vì các nước thu nhập thấp có thể đối mặt với những thách thức khác biệt đáng kể.')

add_p(doc, 'Tuổi người tham gia dao động từ 4–18 tuổi. Bốn nghiên cứu sử dụng can thiệp dựa trên CBT và hai nghiên cứu sử dụng chương trình giáo dục tâm lý (psychoeducational programs). Các phương thức triển khai bao gồm trực tiếp (face-to-face) và kỹ thuật số (digital), được thực hiện bởi giáo viên, chuyên gia sức khỏe tâm thần hoặc nhà nghiên cứu.')

# =====================================================================
# PAGE 19 — Limitations
# =====================================================================
add_page_ref(doc, '19/21', JOURNAL, VOL)

add_p(doc, 'Hạn chế', bold=True)
add_p(doc, '• Số lượng nghiên cứu rất ít (n=6)')
add_p(doc, '• Chất lượng bằng chứng "rất thấp" (GRADE)')
add_p(doc, '• Tính không đồng nhất (heterogeneity) cao giữa các nghiên cứu')
add_p(doc, '• Thiếu dữ liệu theo dõi dài hạn (>3 tháng)')
add_p(doc, '• Phần lớn can thiệp dùng CBT — thiếu đa dạng phương pháp')

# =====================================================================
# PAGE 20 — Conclusion
# =====================================================================
add_page_ref(doc, '20/21', JOURNAL, VOL)
add_heading(doc, 'KẾT LUẬN (Conclusion)', level=2)

add_p(doc, 'Có một số bằng chứng về hiệu quả của can thiệp tại trường học trong việc phòng ngừa lo âu và trầm cảm ở giới trẻ tại LMICs, mặc dù bằng chứng còn hạn chế. Can thiệp cho trầm cảm có vẻ hiệu quả hơn so với lo âu. Cần nghiên cứu bổ sung tại các nước thu nhập thấp, với thời gian theo dõi dài hơn, công cụ đo lường chuẩn hóa, và can thiệp được điều chỉnh phù hợp văn hóa. Điều quan trọng là phát triển các mô hình tiết kiệm chi phí, phù hợp với nguồn lực hạn chế tại LMICs.')

# =====================================================================
# PAGE 21 — References
# =====================================================================
add_page_ref(doc, '21/21', JOURNAL, VOL)
add_heading(doc, 'TÀI LIỆU THAM KHẢO (References)', level=2)
add_p(doc, '[Giữ nguyên 76 tài liệu tham khảo tiếng Anh gốc — xem bản PDF gốc]', italic=True)

# =====================================================================
# ABBREVIATION TABLE
# =====================================================================
doc.add_paragraph()
add_abbreviation_table(doc, [
    ('LMIC', 'Low- and Middle-Income Countries — Các nước thu nhập thấp và trung bình'),
    ('HIC', 'High-Income Countries — Các nước thu nhập cao'),
    ('RCT', 'Randomised Controlled Trial — Thử nghiệm ngẫu nhiên có đối chứng'),
    ('CBT', 'Cognitive-Behavioural Therapy — Liệu pháp hành vi nhận thức'),
    ('PRISMA', 'Preferred Reporting Items for Systematic Reviews and Meta-Analyses'),
    ('GRADE', 'Grading of Recommendations, Assessment, Development and Evaluations'),
    ('CES-D', 'Center for Epidemiologic Studies Depression Scale'),
    ('CES-DC', 'Center for Epidemiologic Studies Depression Scale for Children'),
    ('CDI', 'Children\'s Depression Inventory — Thang đo trầm cảm trẻ em'),
    ('DASS-21', 'Depression, Anxiety and Stress Scales — 21 items'),
    ('RoB 2', 'Cochrane Risk of Bias 2 — Công cụ đánh giá nguy cơ thiên vị'),
    ('DAC', 'Development Assistance Committee — Ủy ban Hỗ trợ Phát triển'),
    ('ODA', 'Official Development Assistance — Viện trợ phát triển chính thức'),
    ('ITT', 'Intention-to-Treat — Phân tích theo ý định điều trị'),
    ('WHO', 'World Health Organization — Tổ chức Y tế Thế giới'),
    ('UNICEF', 'United Nations Children\'s Fund — Quỹ Nhi đồng Liên Hợp Quốc'),
    ('COVID-19', 'Coronavirus Disease 2019 — Bệnh coronavirus 2019'),
    ('VTN', 'Vị thành niên'),
])

# =====================================================================
# RED CRITIQUE SECTION
# =====================================================================
doc.add_paragraph()
add_red_heading(doc, 'PHẢN BIỆN VÀ ĐÁNH GIÁ (Critique)')

add_red(doc, 'ĐIỂM MẠNH (Strengths):', bold=True)
add_red(doc, '1. Chủ đề có ý nghĩa thực tiễn cao: Tập trung vào LMICs — nơi chiếm 88% trẻ em có rối loạn tâm thần toàn cầu nhưng thiếu nghiên cứu nghiêm trọng.')
add_red(doc, '2. Phương pháp luận chặt chẽ: Tuân thủ PRISMA, đăng ký giao thức trên INPLASY, sử dụng RoB 2 và GRADE — đảm bảo tính minh bạch và tái lập.')
add_red(doc, '3. Chỉ đưa vào RCT: Tiêu chuẩn cao nhất cho bằng chứng can thiệp, loại trừ các thiết kế yếu hơn.')
add_red(doc, '4. Phân biệt rõ ràng giữa hiệu quả trên lo âu vs. trầm cảm: Không gộp chung kết quả, giúp nhận diện mô hình khác biệt.')
add_red(doc, '5. Thảo luận sâu về bối cảnh LMIC: Nhấn mạnh các yếu tố văn hóa, kinh tế, kỳ thị ảnh hưởng đến khả năng triển khai.')

add_red(doc, '')
add_red(doc, 'HẠN CHẾ (Limitations):', bold=True)
add_red(doc, '1. Cỡ mẫu nghiên cứu rất nhỏ (n=6): Chỉ 6 RCT từ 3.863 bài sàng lọc. Kết luận mang tính gợi ý hơn là khẳng định.')
add_red(doc, '2. Chất lượng bằng chứng "rất thấp" (GRADE): Tất cả kết quả đều ở mức very low — hạn chế nghiêm trọng độ tin cậy.')
add_red(doc, '3. Không có nghiên cứu từ nước thu nhập thấp: Dù tiêu đề ghi "LMIC", thực tế chỉ có nước thu nhập trung bình — tính đại diện bị giới hạn.')
add_red(doc, '4. Giới hạn ngôn ngữ (chỉ tiếng Anh): Có thể bỏ sót các nghiên cứu quan trọng xuất bản bằng ngôn ngữ bản địa tại LMIC.')
add_red(doc, '5. Thiếu phân tích tổng hợp (meta-analysis): Tính không đồng nhất cao ngăn cản meta-analysis, nhưng điều này hạn chế khả năng ước lượng kích thước hiệu ứng tổng hợp.')
add_red(doc, '6. Thời gian tìm kiếm hẹp (2018–7/2023): Chỉ 5 năm, có thể bỏ sót các NC quan trọng trước 2018.')
add_red(doc, '7. Không đánh giá chi phí-hiệu quả: Yếu tố quan trọng cho khả năng nhân rộng tại LMIC.')

add_red(doc, '')
add_red(doc, 'KHOẢNG TRỐNG NGHIÊN CỨU (Research gaps):', bold=True)
add_red(doc, '1. Cần RCT chất lượng cao tại các nước thu nhập thấp (LIC) — hiện hoàn toàn vắng mặt.')
add_red(doc, '2. Cần đánh giá hiệu quả dài hạn (follow-up >6 tháng) — phần lớn NC chỉ đo ngay sau can thiệp.')
add_red(doc, '3. Cần đa dạng hóa can thiệp ngoài CBT: mindfulness, kỹ năng xã hội-cảm xúc (SEL), can thiệp kỹ thuật số.')
add_red(doc, '4. Cần nghiên cứu về chi phí-hiệu quả và tính bền vững (sustainability) trong bối cảnh nguồn lực hạn chế.')
add_red(doc, '5. Cần điều chỉnh văn hóa (cultural adaptation) — các công cụ đo lường và nội dung can thiệp phải phù hợp bối cảnh.')
add_red(doc, '6. Thiếu dữ liệu về yếu tố điều hòa (moderators): giới tính, tuổi, mức thu nhập, khu vực thành thị/nông thôn.')
add_red(doc, '7. Cần vai trò rõ hơn của giáo viên vs. chuyên gia sức khỏe tâm thần trong triển khai — liên quan đến khả năng nhân rộng (scalability).')

# =====================================================================
# SAVE
# =====================================================================
doc.save(OUTPUT)
print(f'OK — Saved: {OUTPUT}')
