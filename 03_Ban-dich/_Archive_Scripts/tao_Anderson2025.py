# -*- coding: utf-8 -*-
"""
Tao ban dich DOCX: Anderson et al. (2025)
Contributing Factors to the Rise in Adolescent Anxiety and Associated Mental Health Disorders
J Child Adolesc Psychiatric Nursing, 38, e70009
"""
import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from tao_dich_template import (
    create_doc, add_link_and_qr, add_heading, add_page_ref,
    add_p, add_red, add_red_heading, add_info_table, add_table,
    add_abbreviation_table
)

JOURNAL = 'J Child Adolesc Psychiatr Nurs'
VOL = '38, e70009 (2025)'
URL = 'https://doi.org/10.1111/jcap.70009'
QR = 'QR_Anderson2025.png'
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '05_Anderson_2025_Wiley.docx')

doc = create_doc()

# ============================================================
# LINK + QR
# ============================================================
add_link_and_qr(doc, URL, QR)

# ============================================================
# TITLE
# ============================================================
add_heading(doc, 'BÀI BÁO: Các yếu tố góp phần vào sự gia tăng lo âu ở vị thành niên và các rối loạn SKTT liên quan: Tổng quan tường thuật', level=1)

# ============================================================
# INFO TABLE
# ============================================================
info = [
    ('Tên bài báo gốc', 'Contributing Factors to the Rise in Adolescent Anxiety and Associated Mental Health Disorders: A Narrative Review of Current Literature'),
    ('Tác giả', 'Thea L. Anderson, Rasa Valiauga, Christian Tallo, Catriona Blythe Hong, Shaminy Manoranjithan, Catherine Domingo, Manasvi Paudel, Ana Untaroiu, Samantha Barr, Kate Goldhaber'),
    ('Tạp chí', 'Journal of Child and Adolescent Psychiatric Nursing'),
    ('Năm xuất bản', '2025 (online 26/12/2024)'),
    ('Volume/Issue', '38, e70009'),
    ('DOI', '10.1111/jcap.70009'),
    ('Loại bài', 'Tổng quan tường thuật (Narrative Review)'),
    ('Số trang', '10 trang (e70009, trang 1-10)'),
    ('Số bài được tổng quan', '56 bài báo bình duyệt + dữ liệu thăm dò ý kiến'),
    ('Từ khóa', 'vị thành niên (adolescence), lo âu (anxiety), sức khỏe tâm thần (mental health)'),
]
add_info_table(doc, info)
doc.add_paragraph()

# ============================================================
# PAGE 1 -- ABSTRACT
# ============================================================
add_page_ref(doc, 1, JOURNAL, VOL)

add_heading(doc, 'TÓM TẮT (Abstract)', level=2)

add_p(doc, 'Bối cảnh: Tỷ lệ lo âu (anxiety) ở vị thành niên (adolescents) đã có sự gia tăng đáng kể trong những năm gần đây, trở thành một mối quan ngại lớn về sức khỏe cộng đồng (public health). Trên thực tế, lo âu phổ biến hơn đáng kể ở Thế hệ Z (Generation Z) (những cá nhân sinh từ 1997 đến 2012) so với bất kỳ thế hệ nào trong ba thế hệ trước đó. Chúng tôi nhằm mục đích xem xét những yếu tố nào góp phần vào sự gia tăng tỷ lệ lo âu ở thanh thiếu niên và xác định các điểm can thiệp (intervention).', bold=False)

add_p(doc, 'Phương pháp: Nghiên cứu này sử dụng phương pháp tổng quan tường thuật (narrative review). Chúng tôi đã thực hiện tìm kiếm tài liệu trên các cơ sở dữ liệu PubMed, ScienceDirect và Medline, và xác định các bài báo nghiên cứu gốc và bài tổng quan thảo luận về sự gia tăng lo âu và các rối loạn sức khỏe tâm thần (mental health disorders) khác ở Thế hệ Z.', bold=False)

add_p(doc, 'Kết quả: Chúng tôi cung cấp một cái nhìn tổng quan toàn diện về các yếu tố góp phần vào sự gia tăng tỷ lệ lo âu ở vị thành niên, bao gồm áp lực học tập (academic pressures), ảnh hưởng của mạng xã hội (social media influence), động lực gia đình (family dynamics) và các yếu tố gây căng thẳng xã hội rộng lớn hơn (broader societal stressors).', bold=False)

add_p(doc, 'Kết luận: Trong tổng quan tường thuật này, chúng tôi xem xét bản chất đa chiều (multifaceted nature) của lo âu vị thành niên, xác định các yếu tố góp phần. Ngoài ra, chúng tôi thảo luận về các can thiệp tiềm năng dựa trên lâm sàng (clinical), giáo dục (educational) và cộng đồng (community-based) để phòng ngừa và điều trị lo âu vị thành niên. Bằng cách hiểu và giải quyết các nguyên nhân cơ bản của lo âu, có thể giảm thiểu tác động của nó và thúc đẩy quỹ đạo phát triển lành mạnh hơn cho các cá nhân trẻ.', bold=False)

# ============================================================
# PAGE 1-2 -- INTRODUCTION
# ============================================================
add_heading(doc, '1 | Giới thiệu (Introduction)', level=2)

add_p(doc, 'Lo âu vị thành niên (adolescent anxiety) và các rối loạn sức khỏe tâm thần liên quan đã nổi lên như một mối quan ngại cấp bách trong những năm gần đây, với ngày càng nhiều bằng chứng cho thấy sự gia tăng đáng kể về tỷ lệ mắc (prevalence rates). Đại dịch COVID-19 đã dẫn đến nghiên cứu đáng kể liên quan đến tác động của đại dịch và sự cô lập xã hội (social isolation) kết quả đối với sức khỏe tâm thần vị thành niên. Một tổng quan hệ thống (systematic review) gồm 61 bài báo đã tìm thấy sự gia tăng đáng kể về tỷ lệ lo âu và trầm cảm (depression) ở trẻ em và vị thành niên trên toàn thế giới so với mức trước đại dịch (Panchal và cộng sự 2023). Các yếu tố nguy cơ (risk factors) chính được xác định bởi các nghiên cứu trong tổng quan bao gồm các vấn đề sức khỏe tâm thần trước đại dịch và tiếp xúc quá mức với phương tiện truyền thông (excessive media exposure), trong khi giao tiếp gia đình mạnh mẽ (strong family communication) và hỗ trợ xã hội (social support) có tác dụng bảo vệ chống lại sự phát triển của bệnh tâm thần. Ngoài ra, một số nghiên cứu phát hiện rằng các cá nhân có rối loạn phát triển thần kinh (neurodevelopmental disorders) và nhu cầu giáo dục đặc biệt biểu hiện nhiều vấn đề cảm xúc hơn so với các bạn đồng trang lứa phát triển thần kinh bình thường (neurotypical peers) (Nonweiler và cộng sự 2020; Waite và cộng sự 2021). Các bất bình đẳng có sẵn đã bị đại dịch làm trầm trọng thêm, tác động không cân xứng đến thanh thiếu niên từ các cộng đồng thiểu số (minority backgrounds) (Fortuna và cộng sự 2023). Một nghiên cứu đánh giá những thay đổi về tỷ lệ trầm cảm, lo âu và nguy cơ tự tử (suicide risk) ở thanh thiếu niên từ 8-20 tuổi trong giai đoạn 2015-2022 phát hiện rằng sự gia tăng lớn nhất về trầm cảm và lo âu là ở nữ giới gốc Tây Ban Nha (Hispanic) và châu Á (Asian), trong khi nguy cơ tự tử lớn nhất được quan sát ở nữ giới gốc châu Á và nữ giới da đen (Black) (Prichett và cộng sự 2024). Các báo cáo ngành gần đây nhấn mạnh các xu hướng tương đương trong quần thể Thế hệ Z.')

add_page_ref(doc, 2, JOURNAL, VOL)

add_p(doc, 'Báo cáo Tiếng nói của Thế hệ Z của Gallup-Walton Family Foundation (Gallup-Walton Family Foundation Voices of Gen Z Report), dựa trên khảo sát thực hiện tháng 4-5 năm 2023, nêu bật tình trạng đáng báo động về sức khỏe tâm thần ở người trẻ (Walton Family Foundation 2024). Báo cáo phát hiện rằng chỉ 47% thành viên Thế hệ Z (từ 12-26 tuổi) tự coi mình đang phát triển tốt (thriving), so với 59% thế hệ thiên niên kỷ (millennials), 57% Thế hệ X (Gen X), và 52% thế hệ bùng nổ dân số (baby boomers). Một phần của con số đáng kinh ngạc này có thể được giải thích bởi phát hiện rằng Thế hệ Z có nhiều khả năng báo cáo các mối lo ngại về sức khỏe tâm thần hơn (Bethune 2019). Tuy nhiên, sức khỏe tâm thần là đa yếu tố (multifactorial), và những con số gia tăng này cũng có thể được quy cho các ảnh hưởng văn hóa khác nhau, bao gồm các yếu tố sinh học (biological factors), mạng xã hội (social media), động lực gia đình hạt nhân (nuclear family dynamics), áp lực học tập (academic pressures), hoạt động ngoại khóa (extracurriculars) và sự bất ổn toàn cầu (global uncertainty) (McKinsey & Company 2024).')

add_p(doc, 'Gánh nặng của việc sống với bệnh tâm thần là nghiêm trọng, và các phương pháp đối phó có hại (detrimental coping methods) có thể làm trầm trọng thêm những thách thức mà cá nhân phải đối mặt. Theo một nghiên cứu xem xét những thay đổi trong việc sử dụng thuốc từ tháng 1 năm 2016 đến tháng 12 năm 2022, tỷ lệ cấp phát thuốc chống trầm cảm (antidepressant) hàng tháng cho trẻ em, vị thành niên và thanh niên từ 12-25 tuổi đã tăng 66.3% trong khoảng thời gian xác định (Chua và cộng sự 2024). Mặc dù các loại thuốc này có thể cực kỳ hiệu quả trong việc quản lý lo âu và trầm cảm, chúng có thể gây ra các tác dụng phụ đáng chú ý vì thuốc chống trầm cảm mang cảnh báo hộp đen (black box warning) của Cục Quản lý Thực phẩm và Dược phẩm Liên bang (Federal Drug Administration) về khả năng tăng nguy cơ tự tử (suicidality) ở những người dưới 25 tuổi (Reeves và Ladner 2010). Khoảng 30% vị thành niên sẽ trải qua một rối loạn liên quan đến lo âu (anxiety-related disorder), một thống kê tiếp tục tăng, và tỷ lệ nhập viện của thanh thiếu niên có ý định tự tử đã tăng gấp đôi trong thập kỷ qua (Chiu, Falk, và Walkup 2016; Plemmons và cộng sự 2018). Điều này phản ánh tầm quan trọng của việc xác định và hiểu các yếu tố bên ngoài và bên trong góp phần vào sức khỏe tâm thần và thúc đẩy các cơ chế đối phó lành mạnh (healthy coping mechanisms).')

add_p(doc, 'Sự gia tăng tỷ lệ lo âu ở vị thành niên song hành với nhu cầu ngày càng tăng đối với các chuyên gia sức khỏe tâm thần (mental health professionals), nêu bật một khoảng trống quan trọng trong dịch vụ sức khỏe tâm thần (mental health services). Một phân tích tổng hợp (meta-analysis) phát hiện rằng tỷ lệ điều trị kết hợp cho bất kỳ rối loạn sức khỏe tâm thần nào ở trẻ em và vị thành niên chỉ là 38%, nhấn mạnh nhu cầu chăm sóc dễ tiếp cận và toàn diện hơn (Wang và cộng sự 2023). Điều dưỡng chuyên khoa tâm thần (psychiatric nurse practitioners) đóng vai trò then chốt trong chẩn đoán, điều trị và kê đơn cho bệnh nhân trong cả bối cảnh nội trú và ngoại trú (Muench và Fraze 2022). Với tình trạng thiếu hụt nghiêm trọng bác sĩ tâm thần trẻ em và vị thành niên (child and adolescent psychiatrists), điều dưỡng chuyên khoa tâm thần có thể thu hẹp khoảng cách trong tiếp cận dịch vụ sức khỏe tâm thần, cung cấp hỗ trợ quan trọng cho nhóm dân số dễ bị tổn thương này (Yang, Idzik, và Evans 2021).')

add_page_ref(doc, 3, JOURNAL, VOL)

add_p(doc, 'Sự gia tăng lo âu vị thành niên có những hàm ý sâu rộng không chỉ đối với sức khỏe cá nhân mà còn đối với động lực gia đình và chức năng xã hội. Hiểu các yếu tố cơ bản liên quan đến hiện tượng này là thiết yếu để phát triển các chiến lược phòng ngừa (prevention) và can thiệp (intervention) hiệu quả. Tổng quan tường thuật này tổng hợp tài liệu hiện tại về các yếu tố góp phần vào sự gia tăng lo âu vị thành niên và các rối loạn sức khỏe tâm thần liên quan. Bằng cách xem xét một loạt các ảnh hưởng tiềm năng -- cá nhân, gia đình, xã hội và môi trường -- chúng tôi nhằm cung cấp một cái nhìn tổng quan toàn diện về sự tương tác phức tạp ảnh hưởng đến sức khỏe tâm thần vị thành niên. Thông qua việc xem xét này, chúng tôi hy vọng đưa ra những hiểu biết trao quyền cho các chuyên gia sức khỏe tâm thần để giải quyết tốt hơn nhu cầu ngày càng tăng của việc chăm sóc vị thành niên. Tổng quan của chúng tôi nhằm trang bị cho điều dưỡng chuyên khoa tâm thần kiến thức cần thiết để cung cấp chăm sóc phù hợp và hiệu quả hơn.')

# ============================================================
# PAGE 3 -- METHODS
# ============================================================
add_heading(doc, '2 | Phương pháp (Methods)', level=2)

add_p(doc, 'Nghiên cứu hiện tại sử dụng phương pháp tổng quan tường thuật (narrative review). Chúng tôi tìm kiếm các bài báo nghiên cứu sử dụng các cơ sở dữ liệu PubMed, ScienceDirect và Medline. Các thuật ngữ tìm kiếm chính bao gồm: (lo âu vị thành niên HOẶC lo âu HOẶC sức khỏe tâm thần) VÀ (di truyền HOẶC môi trường HOẶC công nghệ HOẶC mạng xã hội HOẶC gia đình HOẶC áp lực học tập HOẶC căng thẳng học tập HOẶC hoạt động ngoại khóa HOẶC chính trị HOẶC biến đổi khí hậu VÀ lo âu vị thành niên). Tiêu chí đưa vào (inclusion criteria) của chúng tôi như sau: (1) các bài báo viết bằng tiếng Anh và xuất bản từ 2000-2023, (2) thảo luận về các vấn đề sức khỏe tâm thần trải nghiệm bởi vị thành niên và thanh niên từ 10-21 tuổi (từ trung học cơ sở đến đại học), và (3) toàn văn bài báo có sẵn. Bất kỳ bài báo nào không đáp ứng cả ba tiêu chí đều bị loại khỏi tổng quan. Tìm kiếm này cho ra 56 bài báo được bình duyệt (peer-reviewed articles), được tổng hợp và trình bày trong bản thảo này. Chúng tôi cũng bao gồm thông tin gần đây từ các tổ chức thăm dò ý kiến tại Hoa Kỳ để cung cấp dữ liệu đại diện cho dân số rộng hơn.')

# ============================================================
# 2.1 Biological Factors
# ============================================================
add_heading(doc, '2.1 | Các Yếu Tố Sinh Học (Biological Factors)', level=3)

add_p(doc, 'Nguyên nhân học (etiology) của các rối loạn lo âu (anxiety disorders) là sự tương tác phức tạp giữa khuynh hướng di truyền (genetic predisposition) và các yếu tố môi trường (environmental factors), với nghiên cứu mới nổi điều tra các sắc thái của mối quan hệ này. Tính di truyền (heritability) 30%-40% của các rối loạn lo âu nhấn mạnh tầm quan trọng của đóng góp di truyền (Hettema, Neale, và Kendler 2001). Bằng chứng từ các nghiên cứu gia đình và sinh đôi (family and twin studies) nêu bật nguy cơ tăng từ bốn đến sáu lần ở người thân bậc một (first-degree relatives). Tuy nhiên, bản chất đa gen (polygenic nature) của lo âu khiến việc xác định chính xác các gen cụ thể trở nên khó khăn. Các nghiên cứu liên kết toàn bộ hệ gen (genome-wide association studies) gần đây đã xác định gen PDE4B là một vị trí nguy cơ mạnh mẽ (robust risk locus) liên quan đến các rối loạn lo âu và liên quan đến căng thẳng, cùng với mối tương quan với các đặc điểm tâm thần khác, bao gồm trầm cảm (depression), tính thần kinh (neuroticism) và tâm thần phân liệt (schizophrenia) (Meier và Deckert 2019; Meier và cộng sự 2019).')

add_page_ref(doc, 4, JOURNAL, VOL)

add_p(doc, 'Song song với các yếu tố di truyền, các yếu tố môi trường, đáng chú ý là ô nhiễm (pollution) và chất độc (toxins), đã thu hút sự chú ý trong những năm gần đây. Các chất ô nhiễm phổ biến bao gồm hạt bụi mịn (particulate matter) và nitơ dioxide (nitrogen dioxide). Sự tiếp xúc ngày càng tăng với các chất độc môi trường, bao gồm các hóa chất gây rối loạn nội tiết (endocrine-disrupting chemicals) như bisphenol A và phthalate, đã là hướng nghiên cứu vì chúng đã được chứng minh ảnh hưởng đến sự phát triển và hành vi trước và sau sinh (pre- and post-natal development) (Bakoyiannis, Kitraki, và Stamatakis 2021). Hơn nữa, nghiên cứu cho thấy ô nhiễm không khí (air pollution) đóng vai trò trong sự phát triển của các rối loạn lo âu và trầm cảm, với khoảng 73% các nghiên cứu chỉ ra mối tương quan dương (positive correlation) (Zundel và cộng sự 2022). Mối tương quan đáng kể này nhấn mạnh tác động của ô nhiễm lên não, bao gồm tăng viêm (increased inflammation), stress oxy hóa (oxidative stress) và mất cân bằng chất dẫn truyền thần kinh (neurotransmitter imbalance).')

add_p(doc, 'Mối quan hệ giữa khuynh hướng di truyền và ảnh hưởng môi trường được làm sáng tỏ thêm bởi các cân nhắc biểu sinh (epigenetic considerations), trong đó siêu methyl hóa DNA (DNA hypermethylation) trong các con đường phát triển thần kinh (neurodevelopment pathways) đã được xác định ở trẻ em và vị thành niên có lo âu. Ngoài ra, một nghiên cứu năm 2021 đã nêu bật quá nhạy cảm tự chủ (autonomic hypersensitivity) với kích thích adrenergic (adrenergic stimulation) và rối loạn điều hòa hoạt động vỏ não trước trán bụng giữa (ventromedial prefrontal cortex) ở những người có rối loạn lo âu tổng quát (generalized anxiety disorder) (Teed và cộng sự 2022). Nhìn chung, những hiểu biết này tổng hợp cho thấy các mục tiêu điều trị tiềm năng đồng thời nhấn mạnh nhu cầu hiểu biết toàn diện hơn về nền tảng sinh học thần kinh (neurobiological) và di truyền (genetic) của lo âu.')

# ============================================================
# 2.2 Digital Technology and Social Media
# ============================================================
add_heading(doc, '2.2 | Công Nghệ Số Và Mạng Xã Hội (Digital Technology and Social Media)', level=3)

add_p(doc, 'Tác động của mạng xã hội (social media) đối với sức khỏe tâm thần vị thành niên là một vấn đề phức tạp đã thu hút sự chú ý đáng kể từ các nhà nghiên cứu, nhà giáo dục và nhà hoạch định chính sách (policymakers). Sự xuất hiện của các nền tảng mạng xã hội đã cách mạng hóa cách vị thành niên tương tác với bạn đồng trang lứa (peers) và nhận thức về bản thân, mang lại cả những hàm ý tích cực và tiêu cực cho sức khỏe tâm thần. Một cuộc khảo sát năm 2023 của Trung tâm Nghiên cứu Pew (Pew Research Center) phát hiện rằng 95% thanh thiếu niên từ 13-17 tuổi có điện thoại thông minh (smartphones) và 96% sử dụng internet hàng ngày. Ngoài ra, 1 trong 5 thanh thiếu niên báo cáo sử dụng YouTube và TikTok gần như liên tục (Auxier 2020). Sự phổ biến của điện thoại thông minh cho phép thanh thiếu niên truy cập thông tin liên tục và tạo điều kiện cho đa nhiệm (multitasking) trong các hoạt động hàng ngày. Tuy nhiên, việc sử dụng công nghệ số tăng lên làm dấy lên lo ngại về các tác động tiêu cực tiềm ẩn đối với sức khỏe tâm thần, chẳng hạn như các vấn đề liên quan đến thời gian sử dụng màn hình (screen time) và tiếp xúc với nội dung trực tuyến tiêu cực (Orben 2020).')

add_page_ref(doc, 5, JOURNAL, VOL)

add_p(doc, 'Việc theo đuổi giáo dục cũng bị ảnh hưởng bởi các công cụ kỹ thuật số. Tại Hoa Kỳ, 6 trong 10 học sinh lớp 8 dựa vào máy tính và internet cho nghiên cứu, học trực tuyến và các dự án hợp tác hàng ngày (Auxier 2020). Tiềm năng gây nghiện (addictive potential) của mạng xã hội, được đặc trưng bởi sử dụng cưỡng chế (compulsive use) và bận tâm với các tương tác trực tuyến, đã được liên kết với các kết quả sức khỏe tâm thần có hại (Shannon và cộng sự 2022). Tuy nhiên, mối quan hệ này rất phức tạp, bị ảnh hưởng bởi chất lượng tương tác trực tuyến và mạng lưới hỗ trợ xã hội ngoại tuyến (offline social support networks). Sự phân biệt giữa sử dụng mạng xã hội chủ động (active) và thụ động (passive) càng làm phức tạp bức tranh này, với việc tham gia chủ động (đăng bài, bình luận, v.v.) có khả năng mang lại lợi ích bảo vệ chống lại sự cô đơn (loneliness) và trầm cảm, trái ngược với tiêu thụ thụ động như lướt qua các bảng tin (scrolling through feeds) (Vaingankar và cộng sự 2022).')

add_p(doc, 'Do đó, điều quan trọng là phải thừa nhận các khía cạnh tích cực của mạng xã hội, đặc biệt về nhận thức sức khỏe tâm thần (mental health awareness) và hỗ trợ trực tuyến. Tính ẩn danh (anonymity) và khả năng tiếp cận (accessibility) của các nền tảng trực tuyến có thể tạo điều kiện hình thành các cộng đồng hỗ trợ, cung cấp phao cứu sinh cho những vị thành niên có thể không có mạng lưới hỗ trợ mạnh mẽ trong cuộc sống ngoại tuyến (Naslund và cộng sự 2020). Các nghiên cứu cung cấp bằng chứng về tác động tích cực của mạng xã hội trong việc giảm triệu chứng trầm cảm và tăng hiểu biết về sức khỏe tâm thần (mental health literacy) ở vị thành niên (Hassen và cộng sự 2022). Những phát hiện này nhấn mạnh tiềm năng của mạng xã hội như một công cụ có giá trị trong thúc đẩy sức khỏe tâm thần, nếu người dùng tham gia vào các hành vi trực tuyến lành mạnh, hỗ trợ.')

add_p(doc, 'Tác động của mạng xã hội đối với sức khỏe tâm thần vị thành niên được định hình sâu sắc bởi sự tương tác giữa các tính dễ tổn thương cá nhân (individual vulnerabilities), bản chất của sự tham gia mạng xã hội và bối cảnh xã hội rộng hơn (Fassi và cộng sự 2024). Điều này nhấn mạnh tầm quan trọng của việc áp dụng cách tiếp cận đa chiều (multidimensional approach) để hiểu và giải quyết các hàm ý sức khỏe tâm thần của việc sử dụng mạng xã hội. Nghiên cứu và can thiệp trong tương lai không chỉ nên nhằm giảm thiểu rủi ro liên quan đến việc sử dụng mạng xã hội mà còn tận dụng tiềm năng của nó như một nguồn lực để thúc đẩy sức khỏe tâm thần và khả năng phục hồi (resilience) ở vị thành niên. Cách tiếp cận cân bằng này đòi hỏi nỗ lực hợp tác giữa các nhà nghiên cứu, nhà giáo dục và nhà hoạch định chính sách để hỗ trợ các chuyên gia sức khỏe tâm thần trong việc phát triển các chiến lược dựa trên bằng chứng (evidence-based strategies) trao quyền cho vị thành niên điều hướng bối cảnh kỹ thuật số theo cách thúc đẩy sức khỏe tâm thần và hạnh phúc của họ.')

# ============================================================
# 2.3 Nuclear Family Dynamics
# ============================================================
add_heading(doc, '2.3 | Động Lực Gia Đình Hạt Nhân (Nuclear Family Dynamics)', level=3)

add_p(doc, 'Nghiên cứu đã ghi nhận những thay đổi trong cấu trúc gia đình (family structure) và các hàm ý tiềm năng của nó đối với kết quả sức khỏe tâm thần của trẻ em và vị thành niên (An và cộng sự 2024). Một thay đổi đáng kể là sự suy giảm trong thực hành "nuôi dạy chung" (alloparenting), bao gồm sự chăm sóc và hỗ trợ trẻ em bởi các cá nhân khác ngoài cha mẹ ruột, chẳng hạn như ông bà, cô dì, chú bác và các thành viên cộng đồng. Hỗ trợ gia đình mở rộng (extended family support) đã là một thực hành phổ biến trong nhiều nền văn hóa trong suốt lịch sử, và các nghiên cứu cho thấy nó có thể cung cấp nguồn lực bổ sung cho cả cha mẹ và trẻ em (Herlosky và Crittenden 2021; Kenkel, Perkeybile, và Carter 2017). Trách nhiệm chia sẻ trong việc nuôi dạy trẻ có thể cung cấp vùng đệm chống lại căng thẳng và yêu cầu của việc làm cha mẹ, cho phép phân bổ khối lượng công việc hợp lý hơn và tăng hỗ trợ về cảm xúc và thực tiễn.')

add_page_ref(doc, 5, JOURNAL, VOL)

add_p(doc, 'Tuy nhiên, khi xã hội trở nên cá nhân hóa hơn (individualistic) và phân tán về mặt địa lý, vai trò của gia đình mở rộng và cộng đồng trong việc nuôi dạy trẻ đã giảm đi (Itao và Kaneko 2021). Sự thay đổi này đã đặt gánh nặng lớn hơn lên cha mẹ, đặc biệt là những người trong các hộ gia đình đơn thân (single-parent) hoặc cả hai vợ chồng đều đi làm (dual-working households), dẫn đến tăng căng thẳng trong các mối quan hệ gia đình. Những thách thức mà cha mẹ đơn thân phải đối mặt, chẳng hạn như căng thẳng tài chính (financial stress), kỳ thị xã hội (social stigma) và hạn chế tiếp cận mạng lưới hỗ trợ, có thể bị trầm trọng thêm bởi sự vắng mặt của hỗ trợ nuôi dạy chung (Taylor và Conger 2017). Sự kết hợp các yếu tố này có thể góp phần vào tỷ lệ tâm bệnh học của cha mẹ (parental psychopathology) tăng cao và các hành vi nuôi dạy tiêu cực (negative parenting behaviors), từ đó làm tăng nguy cơ lo âu và các vấn đề sức khỏe tâm thần khác ở trẻ em (Daryanani và cộng sự 2016; Gruning Parache và cộng sự 2023; Maitri và Venus 2022).')

add_p(doc, 'Sự suy giảm trong thực hành nuôi dạy chung có thể được quy một phần cho sự gia tăng tính di động địa lý (geographic mobility), khi các cá nhân và gia đình có nhiều khả năng sống xa hơn mạng lưới họ hàng mở rộng của họ (Sadruddin và cộng sự 2019). Mặc dù tính di động này có thể cung cấp quyền tiếp cận các cơ hội và nguồn lực mới, nó cũng có thể dẫn đến sự cô lập gia tăng và thiếu hỗ trợ xã hội. Sự vắng mặt của hỗ trợ nuôi dạy chung cũng có thể có tác động trực tiếp đến sự phát triển cảm xúc của trẻ em. Một nghiên cứu năm 2020 phát hiện rằng trẻ em nhận được sự chăm sóc và hỗ trợ nhất quán từ các thành viên gia đình mở rộng có mức độ lo âu và trầm cảm thấp hơn so với những trẻ không nhận được (Van IJzendoorn và cộng sự 2020). Điều này cho thấy sự hiện diện của nhiều người chăm sóc hỗ trợ có thể cung cấp cho trẻ em cảm giác an toàn (sense of security) và khả năng phục hồi cảm xúc (emotional resilience), có thể giúp bảo vệ chống lại sự phát triển của các rối loạn lo âu.')

add_page_ref(doc, 6, JOURNAL, VOL)

add_p(doc, 'Khi các xã hội hiện đại tiếp tục vật lộn với những thách thức của việc cân bằng công việc và cuộc sống gia đình, nhận ra tầm quan trọng của hỗ trợ gia đình mở rộng và tìm cách thúc đẩy các mạng lưới gia đình mở rộng và cộng đồng hỗ trợ có thể rất quan trọng trong việc thúc đẩy sức khỏe tâm thần của cả cha mẹ và trẻ em. Điều này đặc biệt quan trọng đối với các gia đình đơn thân, những người có thể phải đối mặt với các rào cản bổ sung để tiếp cận hỗ trợ. Nghiên cứu trong tương lai nên điều tra các cơ chế cụ thể mà cấu trúc gia đình tác động đến lo âu vị thành niên và phát triển các can thiệp dựa trên bằng chứng có tính đến các cấu hình gia đình đa dạng.')

# ============================================================
# 2.4 Academic Pressure
# ============================================================
add_heading(doc, '2.4 | Áp Lực Học Tập (Academic Pressure)', level=3)

add_p(doc, 'Khi việc theo đuổi giáo dục đại học (higher education) tiếp tục gắn liền với cơ hội nghề nghiệp và thành công tài chính, và với sự cạnh tranh ngày càng tăng để được nhận vào các trường đại học hàng đầu cho các chương trình cử nhân và sau đại học, sinh viên chịu áp lực hơn bao giờ hết để xuất sắc về mặt học thuật. Yêu cầu học tập chắc chắn đã góp phần vào số lượng thanh thiếu niên ngày càng tăng trải qua lo âu và các vấn đề sức khỏe tâm thần khác, cả ở Hoa Kỳ và nước ngoài (Pascoe, Hetrick, và Parker 2020). Một tổng quan hệ thống được thực hiện gần đây về các nghiên cứu trên toàn thế giới xem xét áp lực học tập và sức khỏe tâm thần vị thành niên phát hiện rằng 48 trong số 52 nghiên cứu cho thấy mối tương quan dương giữa áp lực học tập và kết quả sức khỏe tâm thần vị thành niên kém (Steare và cộng sự 2023). Ngoài ra, một nghiên cứu cắt ngang (cross-sectional study) nhỏ (n = 399) trên vị thành niên Hy Lạp từ 12-18 tuổi phát hiện rằng tuổi lớn hơn, giới tính nữ, hộ gia đình đơn thân, nghề nghiệp của cha mẹ và số giờ dành cho học tập có tương quan dương với căng thẳng (stress) và lo âu (Moustaka và cộng sự 2023).')

add_p(doc, 'Tương tự, hai nghiên cứu dọc (longitudinal studies) được thực hiện trên vị thành niên ở Trung Quốc và Ấn Độ đều cho thấy sinh viên chịu áp lực học tập nặng nề có nhiều khả năng mắc rối loạn lo âu hơn (Hua và cộng sự 2023; Trevethan và cộng sự 2022). Tại Hoa Kỳ, Đại học Franciscan ở Ohio đã tìm kiếm lời giải thích cho sự gia tăng mạnh mẽ lượt đến trung tâm tư vấn của trường và phát hiện rằng kết quả học tập và áp lực thành công là những mối quan tâm hàng đầu trong số 374 sinh viên đại học được khảo sát, cho thấy áp lực học tập có thể là một phần nguyên nhân cho các vấn đề sức khỏe tâm thần gia tăng ở sinh viên (Beiter và cộng sự 2015). Hơn nữa, một nghiên cứu điều tra tác động sức khỏe tâm thần của Chương trình Nhập học Đại học Sớm (Early Entrance to College Programs) phát hiện rằng áp lực học tập bổ sung trong các chương trình tăng tốc này đã tăng tính dễ tổn thương của sinh viên đối với việc phát triển các rối loạn sức khỏe tâm thần, bao gồm lo âu và trầm cảm (Singh và cộng sự 2021).')

add_p(doc, 'Các nghiên cứu nói trên đều cho thấy rằng yêu cầu học tập cao đang có tác động tiêu cực đến sức khỏe tâm thần của thanh thiếu niên, góp phần vào sự gia tăng tỷ lệ lo âu ở thanh thiếu niên xuyên văn hóa (cross-culturally). Cần thêm nghiên cứu sử dụng nhóm thuần tập lớn hơn (larger cohorts) để xác định quan hệ nhân quả (causality) và hiểu các can thiệp có thể phù hợp ở đâu để khuyến khích sự xuất sắc học thuật đồng thời giảm thiểu căng thẳng không cần thiết.')

# ============================================================
# 2.5 Extracurricular Activities
# ============================================================
add_page_ref(doc, 7, JOURNAL, VOL)

add_heading(doc, '2.5 | Hoạt Động Ngoại Khóa (Extracurricular Activities)', level=3)

add_p(doc, 'Ngoài thành tích trên lớp, tham gia vào các hoạt động ngoại khóa (extracurricular activities) thường được coi là thành phần quan trọng của thành công học tập. Từ một góc độ, hoạt động ngoại khóa chuyên sâu có thể dẫn đến "lịch trình quá tải" (overscheduling) và gia tăng cảm giác lo âu, như được chứng minh trong một nghiên cứu năm 2007 đánh giá thời gian dành cho hoạt động ngoại khóa ở học sinh trung học tại New York. Mặc dù lo âu được báo cáo ở mức cao hơn ở những sinh viên dành nhiều thời gian hơn cho hoạt động ngoại khóa, kết quả tương tự không được quan sát thấy đối với trầm cảm hoặc rối loạn cơ thể hóa (somatization) (Melman, Little, và Akin-Little 2007). Tuy nhiên, hầu hết nghiên cứu cho thấy rằng tham gia vào hoạt động ngoại khóa mang lại tác dụng bảo vệ (protective effect) đối với sức khỏe tâm thần vị thành niên. Một nghiên cứu xem xét 332 vị thành niên tại các cuộc hẹn y tế thường quy đã đánh giá vai trò của nhiều yếu tố xã hội đối với tỷ lệ triệu chứng trầm cảm (depressive symptoms), và phát hiện rằng những người có nhiều hoạt động ngoại khóa hơn (câu lạc bộ, đội thể thao có tổ chức, v.v.) và mối quan hệ gia đình bình đẳng cao có triệu chứng trầm cảm thấp hơn đáng kể (Mason và cộng sự 2009). Tương tự, một nghiên cứu về học sinh trung học phổ thông ở Hy Lạp phát hiện rằng những người tham gia >11 giờ mỗi tuần hoạt động ngoại khóa báo cáo mức lo âu thấp hơn, trong khi nhiều thời gian hơn dành cho các hoạt động liên quan đến trường học được liên kết với mức lo âu cao hơn (Lazaratou và cộng sự 2013). Một nghiên cứu dọc điều tra vai trò của hoạt động sau giờ học đối với kết quả phát triển phát hiện rằng tham gia thể thao có liên quan đến triệu chứng lo âu và trầm cảm thấp hơn (Fauth, Roth, và Brooks-Gunn 2007). Tuy nhiên, nhiều nghiên cứu lưu ý tầm quan trọng của việc xem xét các loại hoạt động ngoại khóa khác nhau, vì các phát hiện cho đến nay cho thấy các hàm ý khác nhau đối với lo âu, trầm cảm, căng thẳng, khả năng phục hồi (resiliency), sử dụng chất (substance use) và hành vi phạm pháp (delinquency) (Farb và Matjasko 2012).')

add_p(doc, 'Thú vị thay, một nghiên cứu gần đây xem xét thời gian sử dụng màn hình so với tham gia hoạt động ngoại khóa ở học sinh lớp 7 Canada phát hiện rằng vị thành niên tham gia hoạt động ngoại khóa báo cáo ít thời gian sử dụng màn hình giải trí hơn đáng kể, và điều này có liên quan đến sức khỏe tâm thần tốt hơn ở cả nam và nữ (Oberle và cộng sự 2020). Ngược lại, hai giờ hoặc hơn sử dụng màn hình sau giờ học được liên kết với sức khỏe tâm thần kém hơn, đặc biệt đối với nữ, cho thấy rằng tham gia tích cực vào hoạt động ngoại khóa có thể lấp đầy thời gian rảnh mà nếu không có thể được sử dụng cho các hoạt động cô lập hơn và/hoặc có hại hơn như lướt mạng xã hội.')

add_page_ref(doc, 8, JOURNAL, VOL)

add_p(doc, 'Đã được xác lập rõ ràng rằng một số hoạt động ngoại khóa, cụ thể là thể thao có tổ chức (organized sports), có thể có tác dụng bảo vệ đối với triệu chứng lo âu và trầm cảm và thúc đẩy khả năng phục hồi (Binsinger, Laure, và Ambard 2006; Panza và cộng sự 2020; Ruvalcaba và cộng sự 2017). Một phân tích tổng hợp xem xét 29 bài báo và với cỡ mẫu tích lũy 122,056 người tham gia phát hiện hiệu ứng nhỏ nhưng có ý nghĩa thống kê giữa vị thành niên tham gia thể thao và tỷ lệ lo âu và trầm cảm thấp hơn; tuy nhiên, không thể thiết lập hiệu ứng nhân quả (causal effect) (Panza và cộng sự 2020). Tương tự, một nghiên cứu năm 2006 từ Pháp phát hiện rằng thực hành thể thao ngoại khóa thường xuyên có liên quan đến mức tự trọng (self-esteem) và lo âu đặc điểm (trait anxiety) tốt hơn ở vị thành niên trẻ.')

add_p(doc, 'Sự khác biệt giới tính (gender differences) xuất hiện trong tác dụng bảo vệ của tập thể dục ngoại khóa. Tập thể dục ngoại khóa ở nữ có tác dụng bảo vệ nhiều hơn chống lại biến động lớn trong điểm tự trọng Rosenberg (Rosenberg\'s self-esteem scores), trong khi nam không cho thấy sự bảo vệ đáng kể chống lại biến động vừa phải hoặc lớn trong điểm tự trọng hoặc lo âu đặc điểm được đo bằng thang lo âu Spielberger (Spielberger\'s anxiety scale) (Binsinger, Laure, và Ambard 2006). Tổng hợp lại, những phát hiện này cho thấy sự tham gia ngoại khóa phần lớn có tác dụng bảo vệ đối với sức khỏe tâm thần vị thành niên, nhưng cần thêm nghiên cứu để thiết lập quan hệ nhân quả và đặc trưng hóa chi tiết về loại hình tham gia và cường độ thời gian. Cũng cần nghiên cứu thêm về mối quan hệ giữa sự tham gia chuyên sâu vào hoạt động ngoại khóa với mục đích thành công học tập so với tham gia vì giải trí hoặc niềm vui và các tác động tương ứng đối với lo âu vị thành niên.')

# ============================================================
# 2.6 Political and Environmental Uncertainty
# ============================================================
add_heading(doc, '2.6 | Bất Ổn Chính Trị Và Môi Trường (Political and Environmental Uncertainty)', level=3)

add_p(doc, 'Các chính sách chính trị và môi trường đan xen làm trầm trọng thêm căng thẳng, thúc đẩy việc xem xét kỹ lưỡng các yếu tố này trong lo âu vị thành niên. Trong một nghiên cứu được thực hiện trên hơn 10,000 thanh niên từ 16-25 tuổi từ 10 quốc gia, hầu hết người tham gia đều cực kỳ hoặc vừa phải lo lắng về biến đổi khí hậu (climate change), và sự lo âu này có tương quan với nhận thức về phản ứng không đầy đủ của chính phủ hoặc cảm giác bị phản bội (Hickman và cộng sự 2021). Thanh niên ở các nước nghèo hơn đối mặt với tác động lớn hơn từ biến đổi khí hậu có nhiều khả năng báo cáo tần suất lo lắng cao hơn và cảm thấy điều này ảnh hưởng đến cuộc sống hàng ngày. Các hàm ý dài hạn của biến đổi khí hậu, chẳng hạn như di dời (displacement), mất an ninh lương thực (food insecurity) và sụp đổ hệ sinh thái (ecosystem collapse), có thể góp phần vào cảm giác tuyệt vọng và thất vọng ở vị thành niên (Cianconi, Betro, và Janiri 2020). Lo âu liên quan đến bất ổn môi trường cũng có thể tác động tiêu cực đến sự tham gia học tập (academic engagement) và hành vi vì xã hội (prosocial behavior), các yếu tố ảnh hưởng mạnh mẽ đến tương lai của vị thành niên (Kong và Zeng 2023).')

add_p(doc, 'Nhiều yếu tố khác nhau, chẳng hạn như chính sách môi trường, bầu cử và lo ngại về an toàn, có thể dẫn đến bất ổn chính trị (political uncertainty). Một nghiên cứu phát hiện rằng hơn 50% trẻ em có lo âu về ít nhất một vấn đề chính trị, với môi trường và bạo lực súng đạn (gun violence) là hai vấn đề có mức lo ngại cao nhất, trong khi một nghiên cứu khác cho thấy 54% những người sử dụng mạng xã hội trong một cuộc bầu cử gần đây xác định đó là nguồn gây căng thẳng đáng kể (Caporino, Exley, và Latzman 2020; Dejonckheere, Fisher, và Chang 2018). Cùng với bất ổn chính trị, bất ổn xã hội (social unrest) đã được xác định là yếu tố đồng góp phần vào lo âu vị thành niên. Tiếp xúc với bạo lực chính trị (political violence), trực tiếp hoặc thông qua truyền thông, đã làm tăng tỷ lệ rối loạn stress sau sang chấn (posttraumatic stress disorder), trầm cảm và lo âu ở thanh thiếu niên (Slone và Mann 2016). Bản chất lan tỏa của mạng xã hội và chu kỳ tin tức 24 giờ có thể khuếch đại tác động của những yếu tố gây căng thẳng này, khi vị thành niên liên tục tiếp xúc với hình ảnh và câu chuyện đau buồn (Neria và Sullivan 2011). Hơn nữa, sự phân cực (polarization) của diễn ngôn chính trị và sự lan truyền thông tin sai lệch (misinformation) có thể góp phần vào cảm giác bất ổn và mất lòng tin, càng làm trầm trọng thêm triệu chứng lo âu (Strasser, Sumner, và Meyer 2022).')

add_page_ref(doc, 9, JOURNAL, VOL)

add_p(doc, 'Ngoài biến đổi khí hậu và bất ổn chính trị, đại dịch COVID-19 đã ảnh hưởng tiêu cực đến sức khỏe tâm thần thanh thiếu niên. Một báo cáo của Tổng Y sĩ Hoa Kỳ (United States Surgeon General) đã xác định một số yếu tố góp phần vào lo âu gia tăng ở người trẻ trong đại dịch, bao gồm có thành viên gia đình tuyến đầu (frontline family member), sống ở khu vực thành thị với các đợt bùng phát nghiêm trọng hơn, gián đoạn thói quen hàng ngày, và trải qua bất ổn về nhà ở hoặc thực phẩm (Office of the Surgeon General 2021). Một số nhóm, như LGBTQ+, người khuyết tật, và cá nhân từ nền tảng kinh tế xã hội thấp hơn (lower socioeconomic backgrounds), phải đối mặt với thách thức bổ sung do thiếu nguồn lực hỗ trợ và tăng nguy cơ kết quả sức khỏe tâm thần tiêu cực tác động không cân xứng đến vị thành niên da đen (Black) và gốc Tây Ban Nha (Hispanic) (Brooks và cộng sự 2022).')

add_p(doc, 'Khi hậu quả dài hạn của biến đổi khí hậu, bất ổn chính trị và đại dịch COVID-19 tiếp tục bộc lộ, điều quan trọng là ưu tiên nhu cầu sức khỏe tâm thần của vị thành niên. Trao quyền cho vị thành niên với các công cụ và nguồn lực để xây dựng khả năng phục hồi, tham gia vào hoạt động vận động (activism) và nuôi dưỡng ý thức cộng đồng có thể giúp giảm thiểu tác động tâm lý của những yếu tố gây căng thẳng này (Sanson, Van Hoorn, và Burke 2019). Hơn nữa, thanh thiếu niên có thể được hỗ trợ bởi cha mẹ hoặc bạn thân, giúp họ lưu tâm đến lượng thông tin truyền thông tiếp nhận đồng thời hỗ trợ thông tin họ xử lý. Cùng với các biện pháp này, thanh thiếu niên có thể thực hành các thói quen tự chăm sóc lành mạnh (healthy self-care habits) và tìm kiếm sự giúp đỡ chuyên nghiệp. Cần hiểu tác động tâm thần của thay đổi chính trị và môi trường đối với thanh thiếu niên và tạo không gian an toàn (safe spaces) cho trẻ em và vị thành niên để giải quyết nỗi sợ và mối lo ngại của họ.')

# ============================================================
# DISCUSSION
# ============================================================
add_heading(doc, '3 | Thảo Luận (Discussion)', level=2)

add_p(doc, 'Sự gia tăng lo âu vị thành niên có thể được quy cho nhiều yếu tố. Chúng tôi đã xác định sáu xu hướng hiện tại góp phần vào lo âu ở thanh thiếu niên, bao gồm các khía cạnh sinh học (biological aspects), công nghệ số và mạng xã hội (digital technology and social media), động lực gia đình hạt nhân (nuclear family dynamics), học tập (academia), hoạt động ngoại khóa (extracurriculars) và bất ổn chính trị và môi trường (political and environmental uncertainty). Mặc dù các yếu tố góp phần khác tồn tại, các biến số này được tìm thấy là có ảnh hưởng nhất quán, đặc biệt khi đi sâu vào các tính dễ tổn thương cụ thể của vị thành niên. Vị thành niên Thế hệ Z đặc biệt dễ bị tổn thương khi so sánh với các thế hệ khác do sự tiếp xúc và tiêu thụ kiến thức qua mạng xã hội và các yếu tố gây căng thẳng gia tăng từ các khía cạnh khác của cuộc sống. Chúng tôi phát hiện rằng các yếu tố được thảo luận trong tổng quan này bao gồm những thách thức độc đáo mà vị thành niên Thế hệ Z phải đối mặt so với các nhóm tuổi khác, có khả năng giải thích sự khác biệt này trong lo âu và sự xuất hiện của các bệnh sức khỏe tâm thần khác. Lo âu có tỷ lệ phổ biến lớn hơn gấp bốn lần ở Thế hệ Z so với thế hệ Bùng nổ Dân số (Baby Boomer) và khoảng gấp hai lần so với Thế hệ X (Grelle và cộng sự 2023). Những con số này tương tự cho các bệnh tâm thần khác mà Thế hệ Z phải đối mặt, bao gồm rối loạn cơ thể hóa (somatization disorder) và rối loạn trầm cảm chủ yếu (major depressive disorder), nhấn mạnh sự phức tạp của chủ đề này. Hiểu nguồn gốc của cuộc khủng hoảng sức khỏe tâm thần vị thành niên hiện tại sẽ cho phép xác định các lựa chọn điều trị và phòng ngừa tiềm năng.')

add_p(doc, 'Giải quyết lo âu vị thành niên đòi hỏi phải điều hướng giao lộ phức tạp của các yếu tố góp phần, mỗi yếu tố cung cấp cơ hội can thiệp độc đáo. Cách tiếp cận lo âu vị thành niên đòi hỏi kế hoạch điều trị cá nhân hóa (personalized treatment plan), vì lo âu cá nhân có khả năng bắt nguồn từ chỉ một vài yếu tố được mô tả. Thú vị thay, mặc dù có các liệu pháp đã được xác nhận và phê duyệt cho điều trị lo âu và trầm cảm vị thành niên, không có hướng dẫn thiết lập cho phòng ngừa (prevention) (Walter và cộng sự 2020). Cả di truyền và môi trường đều đóng vai trò đáng kể trong sự phát triển của lo âu, với một số biến thể di truyền (genetic variants) khiến cá nhân dễ bị tổn thương hơn trước các tác nhân môi trường (Ask và cộng sự 2021; Gross và Hen 2004; Walter và cộng sự 2023). Do đó, thận trọng khi xem xét các ảnh hưởng sinh học nhưng hiểu rằng thao tác các biến sinh học có thể không luôn là lựa chọn đầu tiên, vì xét nghiệm di truyền (genetic testing) tốn kém, và việc xác định một số dấu ấn di truyền (genetic markers) nhất định có thể không thay đổi kế hoạch điều trị. Do đó, các chuyên gia sức khỏe tâm thần, bao gồm điều dưỡng chuyên khoa tâm thần, nên xem xét các yếu tố nguy cơ có thể thay đổi được (modifiable risk factors) như các điểm can thiệp.')

add_page_ref(doc, 9, JOURNAL, VOL)

add_p(doc, 'Vì lo âu vị thành niên không tồn tại một cách cô lập, điều quan trọng là phải có sự tham gia của người chăm sóc (caregivers) trong quá trình điều trị. Phong cách nuôi dạy (parenting styles) có thể có tác động lớn đến triệu chứng lo âu ở vị thành niên (Romero-Acosta và cộng sự 2021). Hỗ trợ nuôi dạy con lành mạnh và khuyến khích giao tiếp lành mạnh giữa vị thành niên và người chăm sóc có thể cung cấp an toàn cảm xúc (emotional security) và giảm bớt lo âu. Lạm dụng công nghệ số và mạng xã hội, có tương quan mạnh với tỷ lệ lo âu tăng ở thanh thiếu niên, là một lĩnh vực khác có thể hưởng lợi từ can thiệp gia đình (Hamatani và cộng sự 2022). Điều dưỡng chuyên khoa tâm thần có thể xem xét tư vấn người chăm sóc về việc hạn chế thời gian sử dụng màn hình và có thể hỗ trợ người chăm sóc trong việc đề cập chủ đề với vị thành niên. Ngược lại, các chuyên gia có thể chọn tư vấn trực tiếp bệnh nhân để tránh xung đột cha mẹ-con cái (parent-child conflict) có thể làm trầm trọng thêm triệu chứng lo âu ở vị thành niên. Một giải pháp thay thế cho thời gian sử dụng màn hình là tham gia hoạt động ngoại khóa, vì đã được giả thuyết rằng hoạt động ngoại khóa thúc đẩy tương tác xã hội đồng thời giảm tiếp xúc với mạng xã hội và công nghệ số. Do đó, các nhà cung cấp dịch vụ chăm sóc sức khỏe tâm thần có thể xem xét giáo dục gia đình về tầm quan trọng của tham gia hoạt động ngoại khóa, đồng thời vận động cho các trường học và trung tâm cộng đồng tạo và thúc đẩy các nguồn lực này.')

add_p(doc, 'Có một số yếu tố, chẳng hạn như cuộc khủng hoảng môi trường (environmental crisis) và bất ổn toàn cầu (global uncertainty), gây lo âu ở nhiều vị thành niên nhưng nằm ngoài tầm kiểm soát của gia đình. Những cảm xúc này có thể được xoa dịu một phần thông qua giáo dục, vì các nghiên cứu cho thấy giáo dục chính trị và tư tưởng (political and ideological education) có liên quan đến giảm triệu chứng lo âu ở sinh viên (Zhang và Liu 2023). Phát hiện này cũng nhấn mạnh tầm quan trọng của việc không bỏ qua thành phần này của lo âu vị thành niên, vì thanh thiếu niên tiếp xúc với nội dung chính trị toàn cầu hàng ngày thông qua việc sử dụng mạng xã hội. Để chống lại thông tin sai lệch (false information) và giảm bớt lo âu, các trường học có thể xem xét dạy sinh viên cách diễn giải sự bất đồng chính trị và khoa học. Điều này có thể giúp quản lý lo âu nếu được thực hiện một cách nhân từ và không thiên vị. Áp dụng ý tưởng này để giải quyết lo âu liên quan đến bất ổn toàn cầu là phức tạp nhưng mang lại tiềm năng phòng ngừa chủ động, quy mô lớn thay vì quản lý triệu chứng phản ứng.')

add_p(doc, 'Các cách bổ sung mà trường học có thể hỗ trợ giảm bớt lo âu của sinh viên là thông qua các sáng kiến sức khỏe (wellness initiatives) khác nhau. Một bước quan trọng là bình thường hóa cuộc trò chuyện về sức khỏe tâm thần và hỏi sinh viên họ cần gì để cảm thấy được hỗ trợ. Hơn nữa, trường học có thể xem xét tạo không gian an toàn nơi sinh viên có thể bày tỏ cảm xúc, dù cá nhân hay trong nhóm với bạn đồng trang lứa. Khuyến khích giao tiếp cởi mở với giáo viên, cố vấn (counselors) và phụ huynh có thể giúp xác định những mối quan ngại đáng kể có thể hưởng lợi từ điều trị. Các can thiệp dựa trên trường học (school-based interventions) cho vị thành niên có lo âu đã được khám phá rộng rãi trong tài liệu và tập trung chủ yếu vào cung cấp liệu pháp nhận thức hành vi (cognitive behavioral therapy) (Neil và Christensen 2009). Tuy nhiên, mặc dù các phát hiện chung là hứa hẹn về tính hữu ích của việc giảm triệu chứng lo âu, một phân tích tổng hợp năm 2020 gần đây phát hiện rằng triệu chứng lo âu chỉ giảm tạm thời sau can thiệp (transiently post-intervention), và các can thiệp hiện tại được thực hiện bởi trường học không được hỗ trợ bởi bằng chứng hiện có (Gee và cộng sự 2020).')

add_page_ref(doc, 10, JOURNAL, VOL)

add_p(doc, 'Phát hiện sớm (early detection) các vấn đề sức khỏe tâm thần này và giới thiệu đánh giá và điều trị bởi nhà cung cấp dịch vụ chăm sóc sức khỏe tâm thần như điều dưỡng chuyên khoa tâm thần có thể cải thiện đáng kể sức khỏe tâm thần thanh thiếu niên và ngăn ngừa sự phát triển của các vấn đề sức khỏe nghiêm trọng hơn. Ngoài ra, các nhà cung cấp dịch vụ y tế có thể đóng vai trò có giá trị trong việc nuôi dưỡng lòng tự trọng (self-esteem), đây là nền tảng quan trọng cho sức khỏe tâm thần (Fernandes, Newton, và Essau 2022). Các phương pháp điều trị nhấn mạnh phát triển lòng tự trọng có thể mang lại kết quả tích cực (Henriksen và cộng sự 2017).')

add_p(doc, 'Tiếp xúc với thiên nhiên (exposure to nature) đã được đề xuất như một cơ chế tiềm năng mà lo âu có thể được giảm bớt. Một số nghiên cứu gần đây và tổng quan tài liệu chứng minh rằng tiếp xúc với thiên nhiên có thể có vai trò bảo vệ đáng kể đối với sức khỏe tâm thần (Browning và cộng sự 2023; Jimenez và cộng sự 2021; Kotera, Richardson, và Sheffield 2022; Li và cộng sự 2018; Moll và cộng sự 2022). Một tổng quan tường thuật rộng rãi được thực hiện năm 2021 đánh giá tác động của tiếp xúc với thiên nhiên đối với nhiều kết quả sức khỏe phát hiện rằng tăng tiếp xúc với thiên nhiên và tiếp cận không gian xanh (greenspace) dẫn đến cải thiện nhất quán về sức khỏe thể chất và tâm thần (Jimenez và cộng sự 2021). Tương tự, một nghiên cứu đánh giá vai trò của mô phỏng thiên nhiên ảo (virtual nature simulation) đối với lo âu ở sinh viên đại học phát hiện rằng lo lắng dự đoán (anxious apprehension/worry) giảm đáng kể trong mẫu 40 sinh viên tham gia nghiên cứu, và sự giảm lo âu và kích thích lo âu (anxious arousal/panic) tiếp cận mức có ý nghĩa (Browning và cộng sự 2023). Những phát hiện này làm sáng tỏ một hướng điều trị tiềm năng khác cho vị thành niên có lo âu. Thúc đẩy thời gian ngoài trời, dù thông qua các sự kiện có cấu trúc, chẳng hạn như thể thao, hoặc các hoạt động không có cấu trúc, có thể cung cấp một lối thoát tiềm năng cho lo âu mà phần lớn thanh thiếu niên có thể tiếp cận. Cần thêm nghiên cứu về khung thời gian cụ thể mà vị thành niên nên dành ngoài trời để đạt được lợi ích này.')

add_p(doc, 'Điều dưỡng chuyên khoa sức khỏe tâm thần (psychiatric mental health nurse practitioners) là những nhà cung cấp quan trọng của dịch vụ sức khỏe tâm thần cho vị thành niên và tham gia nhiều vào chẩn đoán và quản lý lo âu (Yang, Idzik, và Evans 2021). Điều dưỡng tâm thần (psychiatric nurses) đóng vai trò quan trọng trong chăm sóc sức khỏe tâm thần, với số lượng lượt kê đơn tăng đáng kể từ 2011 đến 2019 (Cai và cộng sự 2022). Hiểu các nguyên nhân gốc rễ mới nổi của lo âu trong quần thể Thế hệ Z sẽ cho phép điều dưỡng tính đến các yếu tố này trong giai đoạn đánh giá (assessment phase) của điều trị (Sampaio và cộng sự 2021). Ngoài ra, thông tin này có thể cải thiện điều trị bằng cách tăng cường liên minh trị liệu (therapeutic alliance) giữa bệnh nhân và đội ngũ chăm sóc (Hartley và cộng sự 2020).')

# ============================================================
# CONCLUSIONS
# ============================================================
add_heading(doc, '4 | Kết Luận (Conclusions)', level=2)

add_p(doc, 'Lo âu vị thành niên bị ảnh hưởng bởi nhiều yếu tố bao gồm các lĩnh vực sinh học (biological), môi trường (environmental) và xã hội (social). Hiểu các yếu tố này có thể hỗ trợ trong việc xác định thanh thiếu niên có nguy cơ (at-risk youth) và phát triển các can thiệp và chiến lược điều trị hiệu quả, từ đó tạo điều kiện cho sự chăm sóc được cung cấp bởi điều dưỡng tâm thần.')

# ============================================================
# AUTHOR CONTRIBUTIONS
# ============================================================
add_heading(doc, 'Đóng Góp Của Tác Giả (Author Contributions)', level=2)

add_p(doc, 'Thea L. Anderson và Rasa Valiauga đã hình thành ý tưởng cho bài tổng quan và dẫn dắt dự án. Thea L. Anderson, Rasa Valiauga, Christian Tallo, Catriona Blythe Hong, Sham Manoranjithan, Catherine Domingo, Manasvi Paudel, Ana Untaroiu và Samantha Barr đã thực hiện tìm kiếm tài liệu và phân tích dữ liệu. Kate Goldhaber đã đóng góp vào thiết kế khung tổng quan và cung cấp phản hồi quan trọng về bản thảo. Tất cả tác giả đã đóng góp vào việc viết và sửa đổi bản thảo, phê duyệt phiên bản cuối cùng và đồng ý chịu trách nhiệm cho tất cả các khía cạnh của công trình.', size=10)

add_p(doc, 'Tuyên bố đạo đức: Các tác giả không có gì để báo cáo.', size=10, italic=True)
add_p(doc, 'Xung đột lợi ích: Các tác giả tuyên bố không có xung đột lợi ích.', size=10, italic=True)
add_p(doc, 'Dữ liệu: Các tác giả không có gì để báo cáo.', size=10, italic=True)

# ============================================================
# KEY NUMBERS TABLE
# ============================================================
add_heading(doc, 'CÁC SỐ LIỆU QUAN TRỌNG', level=2)

key_numbers = [
    ('Số liệu', 'Chi tiết'),
]
key_data = [
    ('61 bài báo', 'Tổng quan hệ thống (Panchal 2023) tìm thấy sự gia tăng đáng kể lo âu/trầm cảm ở trẻ em và VTN so với trước đại dịch'),
    ('56 bài báo', 'Số bài báo bình duyệt được tổng hợp trong tổng quan tường thuật này'),
    ('47% Gen Z', 'Chỉ 47% Thế hệ Z tự coi mình đang phát triển tốt (thriving), so với 59% millennials, 57% Gen X'),
    ('30%-40%', 'Tính di truyền (heritability) của các rối loạn lo âu'),
    ('4-6 lần', 'Nguy cơ tăng ở người thân bậc một (first-degree relatives)'),
    ('73%', 'Các nghiên cứu chỉ ra mối tương quan dương giữa ô nhiễm không khí và lo âu/trầm cảm'),
    ('95% VTN', '95% thanh thiếu niên 13-17 tuổi có smartphone; 96% sử dụng internet hàng ngày'),
    ('48/52 nghiên cứu', 'Cho thấy mối tương quan dương giữa áp lực học tập và SKTT VTN kém'),
    ('66.3%', 'Tăng tỷ lệ cấp phát thuốc chống trầm cảm cho nhóm 12-25 tuổi (2016-2022)'),
    ('~30% VTN', 'Sẽ trải qua một rối loạn liên quan đến lo âu'),
    ('38%', 'Tỷ lệ điều trị kết hợp cho bất kỳ rối loạn SKTT nào ở trẻ em và VTN'),
    ('122,056', 'Cỡ mẫu tích lũy trong phân tích tổng hợp 29 bài báo về thể thao và lo âu/trầm cảm VTN'),
    ('4x', 'Lo âu phổ biến gấp 4 lần ở Gen Z so với Baby Boomer'),
]
add_table(doc, ['Số liệu', 'Chi tiết'], key_data, widths=[3.5, 12.5])

# ============================================================
# ABBREVIATIONS
# ============================================================
abbrevs = [
    ('SKTT', 'Sức khỏe Tâm thần (Mental Health)'),
    ('VTN', 'Vị thành niên (Adolescent)'),
    ('Gen Z', 'Thế hệ Z (Generation Z, sinh 1997-2012)'),
    ('CBT', 'Liệu pháp nhận thức hành vi (Cognitive Behavioral Therapy)'),
    ('SEL', 'Học tập cảm xúc xã hội (Social-Emotional Learning)'),
    ('ADHD', 'Rối loạn tăng động giảm chú ý (Attention Deficit Hyperactivity Disorder)'),
    ('PTSD', 'Rối loạn stress sau sang chấn (Posttraumatic Stress Disorder)'),
    ('GWAS', 'Nghiên cứu liên kết toàn bộ hệ gen (Genome-Wide Association Studies)'),
    ('GAD', 'Rối loạn lo âu tổng quát (Generalized Anxiety Disorder)'),
    ('FDA', 'Cục Quản lý Thực phẩm và Dược phẩm Liên bang Hoa Kỳ (Federal Drug Administration)'),
    ('LGBTQ+', 'Cộng đồng đa dạng giới tính và tình dục'),
    ('COVID-19', 'Bệnh do vi-rút SARS-CoV-2 (Coronavirus Disease 2019)'),
    ('DNA', 'Axit deoxyribonucleic'),
    ('PDE4B', 'Phosphodiesterase 4B (gen nguy cơ lo âu)'),
]
add_abbreviation_table(doc, abbrevs)

# ============================================================
# REFERENCES (selected)
# ============================================================
add_heading(doc, 'TÀI LIỆU THAM KHẢO (References)', level=2)

refs = [
    'An, J., X. Zhu, Z. Shi, và J. An. 2024. "A Serial Mediating Effect of Perceived Family Support on Psychological Well-Being." BMC Public Health 24(1): 940.',
    'Ask, H., R. Cheesman, E. S. Jami, et al. 2021. "Genetic Contributions to Anxiety Disorders." Psychological Medicine 51(13): 2231-2246.',
    'Bakoyiannis, I., E. Kitraki, và A. Stamatakis. 2021. "Endocrine-Disrupting Chemicals and Behaviour." Best Practice & Research Clinical Endocrinology & Metabolism 35(5): 101517.',
    'Beiter, R., R. Nash, M. McCrady, et al. 2015. "The Prevalence and Correlates of Depression, Anxiety, and Stress in a Sample of College Students." Journal of Affective Disorders 173: 90-96.',
    'Browning, M. H. E. M., et al. 2023. "Daily Exposure to Virtual Nature Reduces Symptoms of Anxiety in College Students." Scientific Reports 13(1): 1239.',
    'Chua, K. P., et al. 2024. "Antidepressant Dispensing to US Adolescents and Young Adults: 2016-2022." Pediatrics 153(3): e2023064245.',
    'Fassi, L., et al. 2024. "Social Media Use and Internalizing Symptoms in Clinical and Community Adolescent Samples." JAMA Pediatrics 178(8): 814-822.',
    'Fortuna, L. R., et al. 2023. "The Impact of COVID-19 on Anxiety Disorders in Youth." Child and Adolescent Psychiatric Clinics of North America 32(3): 531-542.',
    'Grelle, K., et al. 2023. "The Generation Gap Revisited: Generational Differences in Mental Health." Journal of Adult Development 16: 1-12.',
    'Hickman, C., et al. 2021. "Climate Anxiety in Children and Young People." The Lancet Planetary Health 5(12): e863-e873.',
    'Hua, Y., et al. 2023. "Association Between Academic Pressure, NR3C1 Gene Methylation, and Anxiety Symptoms Among Chinese Adolescents." BMC Psychiatry 23(1): 376.',
    'Meier, S. M., và J. Deckert. 2019. "Genetics of Anxiety Disorders." Current Psychiatry Reports 21(3): 16.',
    'Oberle, E., et al. 2020. "Screen Time and Extracurricular Activities as Risk and Protective Factors for Mental Health in Adolescence." Preventive Medicine 141: 106291.',
    'Panchal, U., et al. 2023. "The Impact of COVID-19 Lockdown on Child and Adolescent Mental Health: Systematic Review." European Child & Adolescent Psychiatry 32(7): 1151-1177.',
    'Panza, M. J., et al. 2020. "Adolescent Sport Participation and Symptoms of Anxiety and Depression." Journal of Sport and Exercise Psychology 42(3): 201-218.',
    'Prichett, L. M., et al. 2024. "COVID-19 and Youth Mental Health Disparities." Academic Pediatrics 24(5): 837-847.',
    'Steare, T., et al. 2023. "The Association Between Academic Pressure and Adolescent Mental Health Problems." Journal of Affective Disorders 339: 302-317.',
    'Wang, S., et al. 2023. "Treatment Rates for Mental Disorders Among Children and Adolescents." JAMA Network Open 6(10): e2338174.',
    'Walter, H. J., et al. 2020. "Clinical Practice Guideline for the Assessment and Treatment of Children and Adolescents With Anxiety Disorders." JAACAP 59(10): 1107-1124.',
    'Zundel, C. G., et al. 2022. "Air Pollution, Depressive and Anxiety Disorders, and Brain Effects." Neurotoxicology 93: 272-300.',
]

for ref in refs:
    add_p(doc, ref, size=10)

add_p(doc, '(Xem danh sách đầy đủ 56+ tài liệu tham khảo trong bài báo gốc)', size=10, italic=True)

# ============================================================
# RED CRITIQUE
# ============================================================
add_red_heading(doc, 'NHẬN XÉT PHẢN BIỆN (Critical Appraisal)')

add_red(doc, 'ĐIỂM MẠNH (Strengths):', bold=True)
add_red(doc, '1. Phạm vi rộng và toàn diện: Bài tổng quan xem xét sáu nhóm yếu tố góp phần (sinh học, mạng xã hội, gia đình, học tập, ngoại khóa, chính trị/môi trường), cung cấp cái nhìn đa chiều về lo âu VTN thay vì tập trung vào một yếu tố đơn lẻ.')
add_red(doc, '2. Tính thời sự: Tổng hợp dữ liệu từ 2000-2023 bao gồm các nghiên cứu hậu COVID-19, phản ánh bối cảnh SKTT hiện tại của Thế hệ Z.')
add_red(doc, '3. Định hướng thực hành: Bài báo không chỉ mô tả vấn đề mà còn thảo luận các can thiệp tiềm năng cho điều dưỡng chuyên khoa tâm thần, tạo giá trị ứng dụng lâm sàng.')
add_red(doc, '4. Đa nguồn dữ liệu: Kết hợp 56 bài báo bình duyệt với dữ liệu khảo sát dân số (Gallup, Pew Research) để cung cấp bức tranh toàn diện hơn.')

add_red(doc, 'HẠN CHẾ (Limitations):', bold=True)
add_red(doc, '1. Phương pháp tổng quan tường thuật (narrative review) không có quy trình chọn lọc có hệ thống, không đánh giá chất lượng bằng chứng (quality assessment) hay đánh giá nguy cơ sai lệch (risk of bias) -- hạn chế khả năng tổng hợp kết luận mạnh mẽ so với tổng quan hệ thống (systematic review).')
add_red(doc, '2. Thiếu phân tích định lượng: Không có phân tích tổng hợp (meta-analysis) hoặc tổng hợp dữ liệu thống kê, khó đánh giá mức độ ảnh hưởng (effect size) của từng yếu tố.')
add_red(doc, '3. Thiên lệch địa lý: Phần lớn nghiên cứu được trích dẫn từ Hoa Kỳ và các nước phương Tây; hạn chế khả năng tổng quát hóa cho VTN ở các quốc gia đang phát triển, đặc biệt Đông Nam Á hoặc Việt Nam.')
add_red(doc, '4. Không phân biệt rõ ràng quan hệ nhân quả (causation) và tương quan (correlation): Nhiều yếu tố được trình bày dưới dạng tương quan mà không thiết lập được cơ chế nhân quả.')
add_red(doc, '5. Thiếu dữ liệu cập nhật: Tiêu chí tìm kiếm chỉ đến 2023, bỏ sót các nghiên cứu 2024 quan trọng có thể bổ sung thêm bằng chứng.')
add_red(doc, '6. Chủ đề rộng nhưng thiếu chiều sâu: Với 6 nhóm yếu tố trong 10 trang, mỗi yếu tố chỉ được phân tích ở mức bề mặt, thiếu chi tiết về cơ chế sinh học thần kinh hoặc mô hình can thiệp cụ thể.')

add_red(doc, 'KHOẢNG TRỐNG NGHIÊN CỨU (Research Gaps):', bold=True)
add_red(doc, '1. Thiếu nghiên cứu dọc (longitudinal) quy mô lớn theo dõi VTN qua nhiều giai đoạn phát triển để xác lập quan hệ nhân quả giữa các yếu tố và lo âu.')
add_red(doc, '2. Thiếu dữ liệu về tương tác giữa các yếu tố (interaction effects): Các yếu tố sinh học, mạng xã hội, gia đình và học tập có thể tương tác và khuếch đại lẫn nhau nhưng chưa được phân tích.')
add_red(doc, '3. Cần nghiên cứu về hiệu quả can thiệp phòng ngừa (prevention interventions) -- bài báo chỉ ra không có hướng dẫn phòng ngừa nhưng không đề xuất khung nghiên cứu cụ thể.')
add_red(doc, '4. Thiếu dữ liệu từ các nước có thu nhập thấp và trung bình (LMIC), đặc biệt Đông Nam Á, nơi VTN đối mặt với bối cảnh văn hóa và kinh tế xã hội khác biệt đáng kể.')
add_red(doc, '5. Cần nghiên cứu sâu hơn về vai trò của epigenetics (biểu sinh học) như cầu nối giữa yếu tố di truyền và môi trường trong lo âu VTN.')
add_red(doc, '6. Thiếu đánh giá hiệu quả chi phí (cost-effectiveness) của các mô hình can thiệp dựa trên trường học vs. lâm sàng cho VTN có lo âu.')

# ============================================================
# SAVE
# ============================================================
doc.save(OUT)
print(f'Da luu thanh cong: {OUT}')
