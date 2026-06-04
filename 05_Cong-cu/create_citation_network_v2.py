# -*- coding: utf-8 -*-
"""Citation network v2: Clean tooltips (no HTML) + usage guide"""
from pyvis.network import Network
import os, json

BASE = r"c:\Users\OS\OneDrive\read_books\Lo-au"

CORE = {
    "Jenkins2023": {
        "label": "Jenkins (2023)\nUSA, N=75",
        "title": "Jenkins et al. (2023)\nDepression and anxiety among multiethnic middle school students:\nAge, gender, and sociocultural environment\n\n📖 Tạp chí: International Journal of Social Psychiatry, Q2, IF~2.2\n🔬 N=75 học sinh THCS đa sắc tộc (tuổi TB 11.2), Nam California\n🛠 Công cụ: PHQ-9A + GAD-10 + 1-8 phỏng vấn dân tộc học/học sinh\n\n📊 KẾT QUẢ CHÍNH:\n• Lo âu (GAD-10): 50.6% có triệu chứng nhẹ-nặng\n• Trầm cảm (PHQ-9A): 44% có triệu chứng nhẹ-nặng\n• Nữ > nam: trầm cảm p=0.002, lo âu p=0.016\n• PHQ-9A và GAD-10 tương quan mạnh: r²=0.76\n\n⭐ ĐIỂM NỔI BẬT:\n• Nghiên cứu duy nhất dùng phương pháp hỗn hợp (mixed-methods)\n• Dân tộc học phát hiện 3 yếu tố mà sàng lọc không bắt được:\n  (1) Bạo lực trên cơ sở giới tại trường và nhà\n  (2) Lo ngại kinh tế COVID-19, mất việc, bệnh tật gia đình\n  (3) Phân biệt chủng tộc, chính sách nhập cư (gần biên giới Mỹ-Mexico)\n• Nhóm tuổi sớm nhất trong 11 bài (10-11 tuổi)\n\n⚠ HẠN CHẾ: Cỡ mẫu rất nhỏ (N=75), snowball sampling,\nkhông đại diện. GAD-10 sửa đổi thiếu xác thực.\n\n💡 Ý NGHĨA: Cần kết hợp phương pháp định tính bên cạnh\ncông cụ sàng lọc tiêu chuẩn để hiểu bối cảnh văn hóa xã hội.",
        "color": "#E91E63", "size": 30,
    },
    "Saikia2023": {
        "label": "Saikia (2023)\nIndia, N=360",
        "title": "Saikia et al. (2023)\nMental Health Morbidities among Adolescents in Kamrup (Metro), Assam\n\n📖 Tạp chí: Indian J Community Medicine, Q3, IF~0.8\n🔬 N=360 (14-17 tuổi), 10 trường THCS Kamrup Metro, Assam\n🛠 Công cụ: DASS-21 (dịch tiếng Assam) + Kuppuswamy\n\n📊 KẾT QUẢ CHÍNH:\n• Trầm cảm: 22.2% (n=80/360)\n• Lo âu: 24.4% (n=88/360)\n• Căng thẳng: 6.9% (n=25/360)\n\n⭐ ĐIỂM NỔI BẬT:\n• NAM có lo âu CAO HƠN nữ: 30.0% vs 18.9% (P=0.049)\n  → Trái ngược hầu hết y văn quốc tế\n  → Phù hợp với Wen (2020) và Xu (2021) ở TQ\n  → Có thể do vai trò xã hội đặt nam vào áp lực hơn\n• Cha mẹ đơn thân: 52.6% trầm cảm vs 20.5% (P=0.001)\n• Cha mẹ uống rượu: trầm cảm 38.3% vs 16.5% (P<0.001)\n• Chơi game: trầm cảm 27.9% vs 10.8% (P<0.001)\n• Bị lưu ban: trầm cảm 58.6% vs 19.0% (P<0.001)\n• Nghiên cứu ĐẦU TIÊN về SKTT thanh thiếu niên ở Đông Bắc Ấn Độ\n\n⚠ HẠN CHẾ: Chỉ dùng chi-square (đơn biến), thiếu hồi quy đa biến.\nDASS-21 tiếng Assam thiếu báo cáo Cronbach alpha.\nThu thập trùng COVID-19 (4/2019-6/2020).\n\n💡 Ý NGHĨA: Cần mở rộng nghiên cứu SKTT ở Đông Bắc Ấn Độ.\nCần khám phá sâu hơn hiện tượng nam > nữ về lo âu.",
        "color": "#FF9800", "size": 28,
    },
    "Mandaknalli2021": {
        "label": "Mandaknalli (2021)\nIndia, N=450",
        "title": "Mandaknalli & Malusare (2021)\nPrevalence of anxiety among municipality school area\n\n📖 MedPulse Int J Psychology, Vol 18(3), pp 19-22. Không IF, không PubMed.\n🔬 N=450 (lớp 9-12), Mahadevappa Rampure Medical College, Ấn Độ\n🛠 GAD-7, SPSS 23.0\n\n📊 KẾT QUẢ:\n• Nhẹ 49.4%, TB 43.3%, Nặng 7.3% (tổng ~100%)\n• Điểm TB: 16.90 ± 9.02\n• Nữ lo âu nặng: 10.9% vs nam 3.8%\n• Thiếu vận động (P=0.025), giấc ngủ kém (P=0.018), hút thuốc liên quan\n\n⭐ NỔI BẬT: Gần 100% có lo âu → điểm cắt có thể quá thấp.\nSo với Chen (2023) chỉ 13.9% — chênh lệch rất lớn.\n\n⚠ HẠN CHẾ: Tạp chí uy tín thấp. Thiếu thông tin phương pháp.\nKhông có chấp thuận đạo đức. Thiếu phân tích đa biến.",
        "color": "#FF9800", "size": 20,
    },
    "NSCH2020": {
        "label": "NSCH (2020)\nUSA, N=55,162",
        "title": "National Survey of Children's Health (NSCH) 2020\nDữ liệu quốc gia Hoa Kỳ về sức khỏe trẻ em\n\n📖 HRSA/Census Bureau — nguồn dữ liệu chính thức liên bang. Uy tín rất cao.\n🔬 N=55,162 trẻ (18,397 thanh thiếu niên 12-17 tuổi)\n🛠 Khảo sát qua đường bưu điện + online, dữ liệu từ cha mẹ\n\n📊 KẾT QUẢ (dữ liệu 2023):\n• 20.3% thanh thiếu niên có rối loạn tâm thần/hành vi được chẩn đoán\n• Lo âu: 16.1% (phổ biến nhất)\n• Trầm cảm: 8.4%\n• Nữ lo âu 20.1% vs nam 12.3%\n\n⭐ NỔI BẬT:\n• Lo âu TĂNG 61% từ 2016 đến 2023 (10.0% → 16.1%)\n• Trầm cảm tăng 45% (5.8% → 8.4%)\n• 61% gặp khó khăn tiếp cận điều trị (tăng 35% từ 2018)\n• Thanh thiếu niên bị ảnh hưởng: nghỉ học gấp 5 lần, khó kết bạn gấp 10 lần\n\n⚠ HẠN CHẾ: Dữ liệu từ cha mẹ (không trực tiếp từ trẻ) → có thể underestimate.\nDựa trên chẩn đoán → phụ thuộc tiếp cận dịch vụ y tế.\nXu hướng tăng có thể do giảm kỳ thị hơn là tăng bệnh thực sự.\n\n💡 Nguồn tham khảo hàng đầu cho so sánh quốc tế.",
        "color": "#E91E63", "size": 35,
    },
    "Alharbi2019": {
        "label": "Alharbi (2019)\nSaudi, N=1,245",
        "title": "Alharbi et al. (2019)\nDepression and anxiety among high school student at Qassim Region\n\n📖 J Family Med Prim Care, Q3, IF~1.5. Open access.\n🔬 N=1,245 (13-19 tuổi, 55.6% nữ), Qassim, Ả Rập Saudi\n🛠 PHQ-9 (trầm cảm) + GAD-7 (lo âu)\n\n📊 KẾT QUẢ:\n• Trầm cảm: 74% (nhẹ 34%, TB 24.6%, nặng vừa 10.4%, nặng 5%)\n• Lo âu: 63.5% (nhẹ 34.1%, TB 19.5%, nặng 9.8%)\n• Nữ > nam cả 2 tình trạng (P<0.001)\n• Nơi cư trú Buraidah chiếm ưu thế (P<0.001)\n• Hệ thống giáo dục Syllabus liên quan lo âu (P=0.022)\n\n⭐ NỔI BẬT:\n• Tỷ lệ CỰC CAO — 74% trầm cảm (so Chen 23%, NSCH 8.4%)\n• Bài ĐẦU TIÊN ở Ả Rập Saudi dùng PHQ-9/GAD-7\n  (các bài trước dùng DASS)\n• Được trích dẫn rộng rãi: Pakistan, Sudan, Qassim COVID...\n\n⚠ HẠN CHẾ: Ngưỡng PHQ-9 ≥5 (bao gồm nhẹ) → phóng đại tỷ lệ.\nChỉ dùng chi-square, thiếu hồi quy đa biến.\nPHQ-9/GAD-7 phương Tây chưa xác thực cho bối cảnh Ả Rập.",
        "color": "#4CAF50", "size": 30,
    },
    "Nakie2022": {
        "label": "Nakie (2022)\nEthiopia, N=849",
        "title": "Nakie et al. (2022)\nPrevalence and associated factors of depression, anxiety, stress\namong high school students in Northwest Ethiopia\n\n📖 BMC Psychiatry, Q1, IF~4.4. Open access.\n🔬 N=849 (tỷ lệ phản hồi 96.1%, phân tích N=810)\n   Tuổi TB 18.59±1.79, 52.7% nữ, 70.5% nông thôn\n🛠 DASS-21 + SPIN + Oslo-3 + ASSIST\n\n📊 KẾT QUẢ:\n• Trầm cảm: 41.4% (KTC 95%: 38.0-44.8%)\n• Lo âu: 66.7% (KTC 95%: 66.4-66.9%)\n• Căng thẳng: 52.2% (KTC 95%: 49.1-56.0%)\n• Trầm cảm: nữ AOR=1.304, khat AOR=5.595, ám ảnh xã hội AOR=1.416\n• Lo âu: bệnh mãn tính AOR=2.099, tiền sử gia đình AOR=1.777,\n  hút thuốc AOR=4.777\n• Stress: rượu AOR=1.828, nông thôn AOR=1.395, hỗ trợ thấp AOR=1.739\n\n⭐ NỔI BẬT:\n• ĐẦU TIÊN ở châu Phi đánh giá đồng thời trầm cảm + lo âu + stress\n• Sử dụng chất (khat, rượu, thuốc) là yếu tố nguy cơ mạnh\n• 46.5% đã dùng rượu, 10.1% khat, 7.5% thuốc lá\n• Bối cảnh xung đột Ethiopia 4/2021 có thể phóng đại tỷ lệ\n\n⚠ HẠN CHẾ: CI hút thuốc-lo âu 1.407-7304 (có thể lỗi in ấn).\nPhạm vi tuổi rộng 15-25 (bao gồm người trưởng thành).\nDASS-21 tiếng Amharic thiếu thông tin xác thực.",
        "color": "#9C27B0", "size": 32,
    },
    "Chen2023": {
        "label": "Chen (2023)\nChina, N=63,205",
        "title": "Chen et al. (2023)\nPrevalence and associated factors of depressive and anxiety symptoms\namong Chinese secondary school students\n\n📖 BMC Psychiatry, Q1, IF~4.4. Open access.\n🔬 N=63,205 (tỷ lệ phản hồi 99.6%), Tự Cống, miền Tây TQ\n   Tuổi TB 14.3, 49.36% nam, 68.62% THCS\n🛠 PHQ-9 + GAD-7 + MPVS + PSQI + IGDS9-SF\n\n📊 KẾT QUẢ (tỷ lệ có trọng số):\n• Trầm cảm: 23.0% (KTC 95%: 19.6-27.0%)\n• Lo âu: 13.9% (KTC 95%: 11.2-17.0%)\n• ~25% có triệu chứng trầm cảm hoặc lo âu\n• Nữ: OR=1.55. Gia đình không hạt nhân: OR=1.31\n• Rối loạn giấc ngủ: OR=6.99. IGD: OR=5.00\n• Bắt nạt lời nói: OR=1.70. Thao túng xã hội: OR=1.97\n\n⭐ NỔI BẬT:\n• Nghiên cứu LỚN NHẤT từ miền Tây Trung Quốc\n• 5 công cụ đo lường → đánh giá toàn diện nhất\n• Giấc ngủ kém và nghiện game là yếu tố nguy cơ MẠNH NHẤT\n• Phân tích riêng THCS vs THPT: yếu tố khác nhau\n\n⚠ HẠN CHẾ: 1 thành phố duy nhất (Tự Cống).\nCI khá rộng dù N lớn → hiệu ứng cụm.\nSàng lọc, không phải chẩn đoán.",
        "color": "#FFC107", "size": 35,
    },
    "Wen2020": {
        "label": "Wen (2020)\nChina, N=900",
        "title": "Wen et al. (2020)\nLatent Profile Analysis of Anxiety among Junior High School Students\nin Less Developed Rural Regions of China\n\n📖 Int J Environ Res Public Health, Q2, IF~4.6. Open access.\n🔬 N=900 THCS nông thôn tỉnh Giang Tây\n🛠 MHT (Mental Health Test) + LPA (Mplus 7.4)\n\n📊 KẾT QUẢ:\n• 3 phân nhóm lo âu qua LPA:\n  - Thấp (nhóm tham chiếu)\n  - Trung bình\n  - Nặng: 223 em (24.78%)\n• NAM có nguy cơ lo âu TB/nặng CAO HƠN nữ\n• Trường có hỗ trợ SKTT → giảm lo âu nặng\n• Tự tin học tập → giảm lo âu trung bình\n• Áp lực học tập cực kỳ lớn → tăng lo âu đáng kể\n\n⭐ NỔI BẬT:\n• ĐẦU TIÊN dùng LPA (person-centered) cho lo âu nông thôn TQ\n• NAM > nữ — phù hợp Saikia (2023), Xu (2021)\n• Hỗ trợ SKTT tại trường là yếu tố bảo vệ → gợi ý chính sách\n• Nông thôn kém phát triển: trẻ bị bỏ lại, thiếu dịch vụ y tế\n\n⚠ HẠN CHẾ: Công cụ MHT không phổ biến quốc tế.\nThiếu báo cáo chỉ số LPA (BIC, AIC, entropy).\nN=900 từ 1 tỉnh — hạn chế khái quát hóa.",
        "color": "#FFC107", "size": 28,
    },
    "Qiu2022": {
        "label": "Qiu (2022)\nChina, N=2,079",
        "title": "Qiu et al. (2022)\nAssociations of Parenting Style and Resilience\nWith Depression and Anxiety Symptoms in Chinese Middle School Students\n\n📖 Frontiers in Psychology, Q2, IF~3.8. Open access.\n🔬 N=2,079 (tuyển 2,936; hợp lệ 2,879), Hợp Phì, An Huy\n   Tuổi TB 16.7±1.8, 63.4% nam\n🛠 EMBU + CES-D + SAS + SRSMSS (Cronbach α=0.931)\n\n📊 KẾT QUẢ:\n• 3 nhóm phong cách nuôi dạy (LPA):\n  - Tích cực: 58.6% → trầm cảm OR=0.30, lo âu OR=0.32\n  - Tiêu cực: 9.2% → trầm cảm OR=1.82, lo âu OR=2.01\n  - Trung bình: 32.2% (nhóm so sánh)\n• Phục hồi thấp: trầm cảm OR=6.74, lo âu OR=2.80\n• Tương tác nuôi dạy × phục hồi: KHÔNG có ý nghĩa thống kê\n\n⭐ NỔI BẬT:\n• ĐẦU TIÊN kiểm tra tương tác nuôi dạy-phục hồi ở TN\n• Phục hồi thấp có OR trầm cảm CAO NHẤT (6.74) trong 11 bài\n• Giả thuyết chính không được hỗ trợ → giảm giá trị đóng góp mới\n\n⚠ HẠN CHẾ: EMBU đo từ quan điểm TN (thiên lệch nhận thức).\n63.4% nam — mất cân đối giới. CES-D/SAS khác PHQ-9/GAD-7.",
        "color": "#FFC107", "size": 28,
    },
    "Xu2021": {
        "label": "Xu (2021)\nChina, N=373,216",
        "title": "Xu et al. (2021)\nPrevalence and risk factors for anxiety symptoms during COVID-19:\nA large survey among 373,216 students in China\n\n📖 J Affective Disorders, Q1, IF~6.6\n🔬 N=373,216 (THCS 244,193 + THPT 129,023), Hà Nam\n   Thu thập online 4-12/02/2020 (đỉnh COVID-19)\n🛠 GAD-7, SPSS 21.0\n\n📊 KẾT QUẢ:\n• Tỷ lệ lo âu tổng: 9.89% (GAD-7 ≥10)\n• Nhẹ-nặng: 38.42%\n• NAM: 10.11% vs NỮ: 9.66% (OR nữ=0.84, bảo vệ)\n• Nông thôn: 12.80% > Thành phố: 8.77%\n• Lo lắng cao: OR giảm 40-60% so với thấp\n• Hành vi sai hoàn toàn: OR=2.72-2.93\n\n⭐ NỔI BẬT:\n• NGHIÊN CỨU LỚN NHẤT TOÀN CẦU về lo âu TN\n• NAM > nữ — phù hợp Saikia (2023), Wen (2020)\n  Nhưng chênh lệch nhỏ (0.45%) dù P<0.001 → ý nghĩa lâm sàng?\n• Tỷ lệ 9.89% THẤP hơn Racine meta-analysis (25%)\n  → có thể do bảo vệ gia đình trong văn hóa TQ\n• Hà Nam giáp Hồ Bắc (tâm dịch) → bối cảnh đặc thù\n\n⚠ HẠN CHẾ: Thu thập 8 ngày (lo âu tình huống, không phải rối loạn).\nBiến worry/fear có thể = triệu chứng lo âu (tautology).\nChỉ 1 tỉnh Hà Nam.",
        "color": "#FFC107", "size": 40,
    },
    "Bhardwaj2020": {
        "label": "Bhardwaj (2020)\nIndia, N=288",
        "title": "Bhardwaj et al. (2020)\nDepression, Anxiety & Stress among higher secondary students\nof Government schools of Chandigarh, India\n\n📖 J IPHA Chandigarh State Branch — tạp chí địa phương, không IF/PMID\n🔬 N=288, 5 trường công lập Chandigarh, PPS sampling\n🛠 DASS-21\n\n📊 KẾT QUẢ:\n• Trầm cảm: 64.9% (nhẹ 18.8%, TB 29.9%, nặng 13.9%, cực nặng 2.4%)\n• Lo âu: 73.3% (TB 26.5%, nặng 25.3%, cực nặng 21.5%)\n• Căng thẳng: 74.7% (nhẹ 21.5%, TB 44.8%, nặng 8.3%)\n• Lo âu nặng + cực nặng: 46.8%\n\n⭐ NỔI BẬT:\n• TỶ LỆ CAO NHẤT trong 11 bài (lo âu 73.3%)\n• So với Saikia (2023) cùng DASS-21 cùng Ấn Độ: 24.4% → gấp 3 lần!\n• Phân bố lo âu bất thường (không có nhóm nhẹ)\n• Trường công lập, học sinh kinh tế thấp → dễ tổn thương\n\n⚠ HẠN CHẾ: Chỉ thống kê mô tả, KHÔNG phân tích yếu tố liên quan.\nTạp chí địa phương, không peer review rõ ràng.\nN=288 nhỏ. Thiếu STROBE checklist.",
        "color": "#FF9800", "size": 22,
    },
}

EXTERNAL = {
    "Racine2021": {
        "label": "Racine (2021)\nJAMA Pediatrics",
        "title": "Racine et al. (2021)\nGlobal Prevalence of Depression/Anxiety in Children During COVID-19\n\nTạp chí: JAMA Pediatrics, Q1, IF~26.0\nMeta-analysis: 29 nghiên cứu, >80,000 trẻ\n1/4 có trầm cảm, 1/5 có lo âu trong COVID-19",
        "color": "#64B5F6", "size": 25,
    },
    "Polanczyk2015": {
        "label": "Polanczyk (2015)\nJ Child Psychol",
        "title": "Polanczyk et al. (2015)\nWorldwide prevalence of mental disorders in children: Meta-analysis\n\nTạp chí: J Child Psychol Psychiatry, Q1\nTỷ lệ rối loạn tâm thần toàn cầu ở trẻ em: 13.4%",
        "color": "#64B5F6", "size": 22,
    },
    "Spitzer2006": {
        "label": "Spitzer (2006)\nGAD-7 gốc",
        "title": "🛠 CÔNG CỤ: GAD-7 (Generalized Anxiety Disorder 7-item)\n\nSpitzer et al. (2006). Arch Internal Medicine, Q1\nĐược trích dẫn >15,000 lần\n\nMÔ TẢ:\n• 7 câu hỏi, thang Likert 0-3\n• Tổng điểm: 0-21\n• Ngưỡng: 5 nhẹ, 10 trung bình, 15 nặng\n• Thời gian: 2 phút\n• Xác thực: sensitivity 89%, specificity 82%\n\nSỬ DỤNG TRONG 11 BÀI:\n• Alharbi (2019): GAD-7 → lo âu 63.5%\n• Chen (2023): GAD-7 → lo âu 13.9%\n• Xu (2021): GAD-7 → lo âu 9.89%\n• Mandaknalli (2021): GAD-7\n→ Cùng công cụ nhưng tỷ lệ chênh lệch lớn (9.89%-63.5%)\n  do ngưỡng cắt và bối cảnh khác nhau",
        "color": "#B0BEC5", "size": 30,
    },
    "Kroenke2001": {
        "label": "Kroenke (2001)\nPHQ-9 gốc",
        "title": "🛠 CÔNG CỤ: PHQ-9 (Patient Health Questionnaire 9-item)\n\nKroenke et al. (2001). J Gen Intern Med, Q1\nĐược trích dẫn >20,000 lần\n\nMÔ TẢ:\n• 9 câu hỏi dựa trên DSM-IV\n• Thang Likert 0-3, tổng 0-27\n• Ngưỡng: 5 nhẹ, 10 TB, 15 nặng vừa, 20 nặng\n• Thời gian: 2-3 phút\n• Xác thực: sensitivity 88%, specificity 88%\n\nSỬ DỤNG TRONG 11 BÀI:\n• Jenkins (2023): PHQ-9A (phiên bản TN) → trầm cảm 44%\n• Alharbi (2019): PHQ-9 → trầm cảm 74% (ngưỡng ≥5)\n• Chen (2023): PHQ-9 → trầm cảm 23% (ngưỡng ≥10)\n→ Chênh lệch 74% vs 23% phần lớn do ngưỡng cắt (≥5 vs ≥10)",
        "color": "#B0BEC5", "size": 30,
    },
    "Lovibond1995": {
        "label": "Lovibond (1995)\nDASS-21 gốc",
        "title": "🛠 CÔNG CỤ: DASS-21 (Depression Anxiety Stress Scale 21-item)\n\nLovibond & Lovibond (1995). Behav Res Ther, Q1\nĐược trích dẫn >10,000 lần\n\nMÔ TẢ:\n• 21 mục: 7 trầm cảm + 7 lo âu + 7 căng thẳng\n• Thang Likert 0-3, nhân 2 → 0-42 mỗi thang\n• Ngưỡng lo âu: 0-7 bình thường, 8-9 nhẹ, 10-14 TB, 15-19 nặng, ≥20 cực nặng\n• Thời gian: 5-10 phút\n• Ưu điểm: đo đồng thời 3 tình trạng\n\nSỬ DỤNG TRONG 11 BÀI:\n• Saikia (2023): lo âu 24.4%, stress 6.9%\n• Nakie (2022): lo âu 66.7%, stress 52.2%\n• Bhardwaj (2020): lo âu 73.3%, stress 74.7%\n→ Cùng DASS-21 cùng Ấn Độ: Saikia 24.4% vs Bhardwaj 73.3% = gấp 3 lần!\n  Cần xem xét bối cảnh, cách lấy mẫu, thời điểm",
        "color": "#B0BEC5", "size": 28,
    },
    "GBD2020": {
        "label": "GBD Study (2020)\nLancet Psychiatry",
        "title": "Global Burden of Disease Study (2020)\nMental disorders burden globally\n\nTạp chí: Lancet Psychiatry, Q1, IF~65\nGánh nặng bệnh tâm thần toàn cầu 1990-2021",
        "color": "#64B5F6", "size": 25,
    },
    "Sandal2017": {
        "label": "Sandal (2017)\nChandigarh",
        "title": "Sandal et al. (2017)\nDepression, anxiety, stress — adolescents Chandigarh\n\nTạp chí: J Family Med Prim Care\nDASS-21: trầm cảm 65.53%, lo âu 80.85%\nĐược Saikia và Bhardwaj trích dẫn",
        "color": "#64B5F6", "size": 18,
    },
    "PLOSONE2025": {
        "label": "PLOS ONE (2025)\nDAS Secondary",
        "title": "Prevalence and determinants of DAS among secondary school students (2025)\n\nTạp chí: PLOS ONE, Q1\nTrích dẫn: Jenkins 2023, Nakie 2022",
        "color": "#A5D6A7", "size": 18,
    },
    "QassimCOVID2023": {
        "label": "Qassim COVID\n(2023)",
        "title": "Depression/Anxiety among Qassim Univ Students During COVID-19 (2023)\n\nTạp chí: PMC\nTrích dẫn: Alharbi 2019",
        "color": "#A5D6A7", "size": 16,
    },
    # ===== BÀI VIỆT NAM =====
    "VNAMHS2023": {
        "label": "V-NAMHS (2023)\n🇻🇳 Vietnam\nN=9,781",
        "title": "V-NAMHS — Khảo sát Quốc gia về Sức khỏe Tâm thần\nThanh thiếu niên Việt Nam (2023)\n\n📖 Báo cáo UNICEF + ĐH Queensland + Bộ LĐTBXH\n🔬 N=9,781, đại diện quốc gia, DASS-21\n📊 Trầm cảm nhẹ-nặng: 10%, Lo âu nhẹ-nặng: 15.6%\n\n⭐ NGHIÊN CỨU ĐẠI DIỆN QUỐC GIA ĐẦU TIÊN của VN\nCỡ mẫu lớn nhất ĐNA về SKTT thanh thiếu niên.\nSử dụng DASS-21 — so sánh trực tiếp với Saikia, Nakie, Bhardwaj.\nTỷ lệ thấp hơn nhiều nước (10-15.6% vs 24-73%)\n→ Có thể do phương pháp đại diện quốc gia vs mẫu thuận tiện.\n\n⚠ Báo cáo chính phủ, không phải bài báo peer-reviewed.\nPhương pháp chi tiết cần xem báo cáo đầy đủ.",
        "color": "#FF0000", "size": 35,
    },
    "VN_COVID2023": {
        "label": "VN COVID\n(2023) 🇻🇳\nN=8,473",
        "title": "Levels of Stress, Anxiety, and Depression in Adolescents\nduring and after COVID-19 Pandemic in Vietnam (2023)\n\n📖 American J Psychiatric Rehabilitation\n🔬 N=8,473 thanh thiếu niên, 6 tỉnh VN, DASS-21\n📊 Trong COVID: stress 65.5%, lo âu 41.5%, trầm cảm 34.2%\nSau COVID: stress 55.4%, lo âu 25.4%, trầm cảm 20.1%\n\n⭐ So sánh TRONG vs SAU COVID-19 ở VN.\nLo âu giảm 16 điểm % sau đại dịch.\nCỡ mẫu lớn (8,473) từ nhiều tỉnh.\n\n⚠ Tạp chí ít được biết đến quốc tế.",
        "color": "#FF0000", "size": 30,
    },
    "VN_HCM2022": {
        "label": "HCM City\n(2022) 🇻🇳\nN=384",
        "title": "Prevalence and Factors Associated with Depression\namong High School Students in Ho Chi Minh City, Vietnam (2022)\n\n📖 National J Community Medicine\n🔬 N=384 (165 nam, 219 nữ), THPT TP.HCM, CES-D\n📊 Trầm cảm: 57.6% (CES-D mean=17.3)\n\n⭐ Tỷ lệ cao ở TP.HCM — đô thị lớn nhất VN.\nSử dụng CES-D (cùng công cụ với Qiu 2022).\nNữ > nam. Áp lực học tập là yếu tố chính.\n\n⚠ Cỡ mẫu nhỏ (N=384). Một trường duy nhất.",
        "color": "#FF0000", "size": 22,
    },
    "VN_Stress2022": {
        "label": "VN Academic\nStress (2022) 🇻🇳",
        "title": "Academic stress and depression among Vietnamese adolescents:\nA moderated mediation model (2022)\n\n📖 Current Psychology, Springer, Q2\n🔬 Thanh thiếu niên Việt Nam, mô hình trung gian\n📊 Áp lực học tập → trầm cảm, điều tiết bởi hài lòng cuộc sống\nvà khả năng phục hồi\n\n⭐ Cùng chủ đề phục hồi (resilience) với Qiu (2022).\nChứng minh vai trò trung gian của hài lòng cuộc sống.\nBối cảnh giáo dục VN — áp lực thi cử cao.\n\n⚠ Không báo cáo tỷ lệ trầm cảm/lo âu cụ thể.",
        "color": "#FF0000", "size": 22,
    },
    "VN_Ethnic2024": {
        "label": "VN Ethnic\nMinority (2024) 🇻🇳",
        "title": "Mental health among ethnic minority adolescents in Vietnam\nand correlated factors (2024)\n\n📖 ScienceDirect (Mental Health & Prevention)\n🔬 Thanh thiếu niên dân tộc thiểu số, trường nội trú VN\n📊 Stress: 24.7%, Lo âu: 54.4%, Trầm cảm: 59.0%\n\n⭐ Tỷ lệ RẤT CAO ở nhóm dân tộc thiểu số.\nLo âu 54.4% — cao hơn nhiều so với V-NAMHS (15.6%).\nCho thấy bất bình đẳng SKTT giữa Kinh và dân tộc thiểu số.\n\n⚠ Nhóm đặc thù (trường nội trú) — không đại diện toàn quốc.",
        "color": "#FF0000", "size": 25,
    },
    # ===== 10 BÀI MỚI THÊM =====
    "Nepal2023": {
        "label": "Nepal (2023)\nN=453",
        "title": "Depression, anxiety and stress among high school students:\nA cross-sectional study in Kathmandu, Nepal (2023)\n\n📖 PLOS Global Public Health\n🔬 N=453, THPT Tokha Municipality, Kathmandu\n🛠 DASS-21\n📊 Trầm cảm 34.2%, Lo âu 40.0%, Stress 22.7%\n\n⭐ Cùng DASS-21 với Saikia, Nakie, Bhardwaj.\nĐông Nam Á — gần Việt Nam về bối cảnh.\nNữ > nam cho cả 3 tình trạng.",
        "color": "#00BCD4", "size": 22,
    },
    "SSAfrica2021": {
        "label": "Sub-Saharan\nReview (2021)",
        "title": "Prevalence of mental health problems in sub-Saharan adolescents:\nSystematic Review (2021)\n\n📖 PLOS ONE, Q1\n🔬 Review 72 nghiên cứu, châu Phi hạ Sahara\n📊 Trầm cảm: 27%, Lo âu: 30% ở TN 10-19 tuổi\n\n⭐ Cung cấp baseline cho châu Phi.\nNakie (2022) lo âu 66.7% >> trung bình châu Phi 30%\n→ Bối cảnh xung đột Ethiopia có thể phóng đại.",
        "color": "#64B5F6", "size": 22,
    },
    "GBD2021Lancet": {
        "label": "GBD 2021\nTransl Psychiatry",
        "title": "Global burden of mental disorders among adolescents\nand young adults, 1990-2021 (2025)\n\n📖 Translational Psychiatry (Nature), Q1\n🔬 GBD Study 2021 — phân tích toàn cầu\n📊 Trầm cảm và lo âu tăng >50% từ 1990-2021\nDALY tâm thần chiếm 15% gánh nặng bệnh toàn cầu ở TN\n\n⭐ Nguồn tham khảo hàng đầu cho bối cảnh toàn cầu.\nXác nhận xu hướng tăng mà NSCH (2020) ghi nhận.",
        "color": "#64B5F6", "size": 25,
    },
    "LMIC2025": {
        "label": "LMIC Review\n(2025)",
        "title": "Prevalence of Anxiety and Depression Among Children/Adolescents\nin Low- and Middle-Income Countries (2025)\n\n📖 Psychiatric Research and Clinical Practice (APA)\n🔬 Systematic review, N=208,842 (1-20 tuổi)\n📊 Tổng hợp tỷ lệ lo âu/trầm cảm ở các nước thu nhập thấp-TB\n\n⭐ Bao gồm bối cảnh giống Ấn Độ, Ethiopia, TQ nông thôn.\nCung cấp benchmark cho so sánh 11 bài.",
        "color": "#64B5F6", "size": 22,
    },
    "Uganda2025": {
        "label": "Uganda (2025)\nN=2,845",
        "title": "Depression/anxiety among school-going adolescents\nin poverty and conflict-affected settings, Uganda (2025)\n\n📖 BMC Psychiatry, Q1\n🔬 N=2,845 (14-17 tuổi), hậu xung đột\n📊 So sánh vùng xung đột vs không xung đột\n\n⭐ Cùng bối cảnh xung đột như Nakie (2022) Ethiopia.\nBMC Psychiatry Q1 — cùng tạp chí với Nakie và Chen.",
        "color": "#A5D6A7", "size": 20,
    },
    "SouthAfrica2022": {
        "label": "South Africa\n(2022)",
        "title": "Validity of PHQ-9 and GAD-7 among adolescents in South Africa (2022)\n\n📖 J Adolescent Health, Q1\n🔬 N=302, isiXhosa translation, 10-19 tuổi\n📊 Xác thực PHQ-9 và GAD-7 cho bối cảnh châu Phi\n\n⭐ Quan trọng cho phản biện: xác thực công cụ phương Tây\nở các nước đang phát triển (liên quan Alharbi, Nakie).",
        "color": "#A5D6A7", "size": 18,
    },
    "TangMeta2019": {
        "label": "Tang Meta\n(2019)",
        "title": "Tang et al. (2019)\nPrevalence of depressive symptoms among adolescents in secondary school\nin mainland China: a systematic review and meta-analysis\n\n📖 J Affect Disord, Q1, IF~6.6\n🔬 51 nghiên cứu, TQ đại lục\n📊 Tỷ lệ trầm cảm: 24.3%\n\n⭐ Chen (2023) báo cáo 23.0% — rất gần meta-analysis này.\nXác nhận tỷ lệ ~1/4 học sinh TQ có triệu chứng trầm cảm.",
        "color": "#64B5F6", "size": 22,
    },
    "KumarImphal2017": {
        "label": "Kumar (2017)\nManipur",
        "title": "Kumar KS, Akoijam BS (2017)\nDepression, anxiety, stress among higher secondary school students\nof Imphal, Manipur\n\n📖 Indian J Community Med\n🔬 DASS-21 tại Manipur, Đông Bắc Ấn Độ\n📊 Trầm cảm 19.5%, Lo âu 24.4%, Stress 21.1%\n\n⭐ Cùng Đông Bắc Ấn Độ như Saikia (2023) — lo âu 24.4% giống hệt!\nĐược Saikia và Nakie trích dẫn.",
        "color": "#64B5F6", "size": 18,
    },
    "ZhouCOVID2020": {
        "label": "Zhou (2020)\nCOVID China",
        "title": "Zhou et al. (2020)\nPrevalence and socio-demographic correlates of psychological health\nproblems in Chinese adolescents during COVID-19\n\n📖 Eur Child Adolesc Psychiatry, Q1\n🔬 N=8,079, TQ, COVID-19\n📊 Trầm cảm 43.7%, Lo âu 37.4%\n\n⭐ Tỷ lệ cao hơn nhiều so với Xu (2021) 9.89%\n→ Khác biệt công cụ và thời điểm thu thập.\nĐược Chen (2023) trích dẫn.",
        "color": "#64B5F6", "size": 20,
    },
    "Daya2018": {
        "label": "Daya (2018)\nTamil Nadu",
        "title": "Daya PA, Karthikeyan G (2018)\nDepression, anxiety, stress and its correlates among urban\nschool going adolescents in Tamil Nadu, India\n\n📖 Int J Res Med Sci\n🔬 DASS-21 tại Tamil Nadu\n📊 Lo âu cao hơn Saikia (2023) — biến thiên vùng Ấn Độ\n\n⭐ Được Saikia trích dẫn.\nCho thấy tỷ lệ SKTT khác nhau đáng kể giữa các bang Ấn Độ.",
        "color": "#64B5F6", "size": 16,
    },
}

EDGES = [
    ("Nakie2022", "Alharbi2019", "Nakie trích dẫn Alharbi"),
    ("Saikia2023", "Sandal2017", "Saikia trích dẫn Sandal"),
    ("Bhardwaj2020", "Sandal2017", "Bhardwaj trích dẫn Sandal"),
    ("Alharbi2019", "Spitzer2006", "Sử dụng GAD-7"),
    ("Alharbi2019", "Kroenke2001", "Sử dụng PHQ-9"),
    ("Chen2023", "Spitzer2006", "Sử dụng GAD-7"),
    ("Chen2023", "Kroenke2001", "Sử dụng PHQ-9"),
    ("Xu2021", "Spitzer2006", "Sử dụng GAD-7"),
    ("Mandaknalli2021", "Spitzer2006", "Sử dụng GAD-7"),
    ("Saikia2023", "Lovibond1995", "Sử dụng DASS-21"),
    ("Nakie2022", "Lovibond1995", "Sử dụng DASS-21"),
    ("Bhardwaj2020", "Lovibond1995", "Sử dụng DASS-21"),
    ("Saikia2023", "GBD2020", "Trích dẫn GBD"),
    ("Nakie2022", "GBD2020", "Trích dẫn GBD"),
    ("Mandaknalli2021", "Polanczyk2015", "Trích dẫn meta-analysis"),
    ("Xu2021", "Racine2021", "Cùng chủ đề COVID-19"),
    ("Jenkins2023", "Racine2021", "Bối cảnh COVID-19"),
    ("PLOSONE2025", "Jenkins2023", "Trích dẫn Jenkins"),
    ("PLOSONE2025", "Nakie2022", "Trích dẫn Nakie"),
    ("QassimCOVID2023", "Alharbi2019", "Trích dẫn Alharbi"),
    ("Chen2023", "Xu2021", "Cùng TQ, cùng GAD-7"),
    ("Chen2023", "Wen2020", "Cùng TQ, so sánh vùng"),
    ("Qiu2022", "Wen2020", "Cùng TQ, cùng LPA"),
    ("Xu2021", "Wen2020", "Cùng TQ, nam > nữ"),
    ("Saikia2023", "Bhardwaj2020", "Cùng Ấn Độ, DASS-21"),
    ("Jenkins2023", "NSCH2020", "Cùng Hoa Kỳ"),
    # 10 bài mới
    ("Nepal2023", "Lovibond1995", "Sử dụng DASS-21"),
    ("Nepal2023", "Nakie2022", "Cùng DASS-21, cùng chủ đề"),
    ("SSAfrica2021", "Nakie2022", "Nakie so sánh với review này"),
    ("Uganda2025", "Nakie2022", "Cùng châu Phi, cùng BMC Psychiatry"),
    ("GBD2021Lancet", "NSCH2020", "Xác nhận xu hướng tăng"),
    ("GBD2021Lancet", "GBD2020", "Cập nhật GBD"),
    ("LMIC2025", "Saikia2023", "Bao gồm bối cảnh Ấn Độ"),
    ("LMIC2025", "Nakie2022", "Bao gồm bối cảnh Ethiopia"),
    ("SouthAfrica2022", "Spitzer2006", "Xác thực GAD-7 châu Phi"),
    ("SouthAfrica2022", "Kroenke2001", "Xác thực PHQ-9 châu Phi"),
    ("TangMeta2019", "Chen2023", "Chen trích dẫn Tang meta"),
    ("Chen2023", "ZhouCOVID2020", "Chen trích dẫn Zhou"),
    ("ZhouCOVID2020", "Xu2021", "Cùng TQ COVID-19"),
    ("KumarImphal2017", "Saikia2023", "Saikia trích dẫn Kumar"),
    ("KumarImphal2017", "Lovibond1995", "Sử dụng DASS-21"),
    ("Saikia2023", "Daya2018", "Saikia trích dẫn Daya"),
    ("Daya2018", "Lovibond1995", "Sử dụng DASS-21"),
    # Bài Việt Nam
    ("VNAMHS2023", "Lovibond1995", "Sử dụng DASS-21"),
    ("VN_COVID2023", "Lovibond1995", "Sử dụng DASS-21"),
    ("VNAMHS2023", "Racine2021", "So sánh tỷ lệ COVID-19"),
    ("VN_COVID2023", "Xu2021", "Cùng chủ đề COVID-19 thanh thiếu niên"),
    ("VN_HCM2022", "Qiu2022", "Cùng dùng CES-D"),
    ("VN_Stress2022", "Qiu2022", "Cùng chủ đề resilience"),
    ("VN_Ethnic2024", "VNAMHS2023", "So sánh thiểu số vs quốc gia VN"),
    ("VN_Ethnic2024", "Saikia2023", "Cùng bối cảnh thiểu số/vùng sâu"),
    ("VN_COVID2023", "VNAMHS2023", "Cùng DASS-21 tại VN"),
    ("Nepal2023", "VNAMHS2023", "Cùng khu vực ĐNA, DASS-21"),
]

# Build network
net = Network(height="850px", width="100%", bgcolor="#0d1117", font_color="white",
              directed=True, notebook=False, cdn_resources="remote")

net.barnes_hut(gravity=-4000, central_gravity=0.3, spring_length=180, damping=0.09)

for pid, d in CORE.items():
    net.add_node(pid, label=d["label"], title=d["title"], color=d["color"],
                 size=d["size"], borderWidth=3, shape="dot",
                 font={"size": 10, "color": "white", "face": "Arial"})

for pid, d in EXTERNAL.items():
    net.add_node(pid, label=d["label"], title=d["title"], color=d["color"],
                 size=d["size"], borderWidth=1, shape="dot",
                 font={"size": 9, "color": "#aaa", "face": "Arial"})

for s, t, lbl in EDGES:
    net.add_edge(s, t, title=lbl, color="#444", width=1.5, arrows="to")

net.set_options('{"interaction":{"hover":true,"tooltipDelay":50,"navigationButtons":true},"physics":{"barnesHut":{"gravitationalConstant":-4000,"centralGravity":0.3}}}')

output = os.path.join(BASE, "DocFiles", "citation_network.html")
net.save_graph(output)

# Inject legend + guide
with open(output, 'r', encoding='utf-8') as f:
    html = f.read()

inject = """
<div style="position:fixed;top:10px;left:10px;background:rgba(13,17,23,0.95);color:#e6edf3;padding:18px;border-radius:12px;font-family:Arial,sans-serif;font-size:12px;z-index:1000;max-width:320px;border:1px solid #30363d;">
<h2 style="margin:0 0 12px 0;color:#58a6ff;font-size:16px;">MẠNG LƯỚI TRÍCH DẪN</h2>
<p style="margin:0 0 8px 0;color:#8b949e;font-size:11px;">11 nghiên cứu về lo âu và trầm cảm ở học sinh</p>
<hr style="border-color:#30363d;margin:8px 0;">
<p style="margin:3px 0;"><span style="color:#E91E63;">⬤</span> Hoa Kỳ (USA)</p>
<p style="margin:3px 0;"><span style="color:#FF9800;">⬤</span> Ấn Độ (India)</p>
<p style="margin:3px 0;"><span style="color:#4CAF50;">⬤</span> Ả Rập Saudi</p>
<p style="margin:3px 0;"><span style="color:#9C27B0;">⬤</span> Ethiopia</p>
<p style="margin:3px 0;"><span style="color:#FFC107;">⬤</span> Trung Quốc (China)</p>
<p style="margin:3px 0;"><span style="color:#64B5F6;">⬤</span> Bài tham chiếu uy tín</p>
<p style="margin:3px 0;"><span style="color:#B0BEC5;">⬤</span> Công cụ đo lường gốc</p>
<p style="margin:3px 0;"><span style="color:#A5D6A7;">⬤</span> Bài trích dẫn 11 bài gốc</p>
<hr style="border-color:#30363d;margin:8px 0;">
<h3 style="margin:0 0 6px 0;color:#58a6ff;font-size:13px;">HƯỚNG DẪN SỬ DỤNG</h3>
<p style="margin:3px 0;">🖱 <b>Hover chuột</b> vào node → xem chi tiết bài báo</p>
<p style="margin:3px 0;">🔍 <b>Cuộn chuột</b> → phóng to/thu nhỏ</p>
<p style="margin:3px 0;">✋ <b>Kéo thả node</b> → di chuyển vị trí</p>
<p style="margin:3px 0;">👆 <b>Click node</b> → highlight kết nối</p>
<p style="margin:3px 0;">🔄 <b>Click nền trắng</b> → reset highlight</p>
<p style="margin:3px 0;">⬤ <b>Kích thước node</b> = cỡ mẫu tương đối</p>
<p style="margin:3px 0;">➡ <b>Mũi tên</b> = hướng trích dẫn</p>
<hr style="border-color:#30363d;margin:8px 0;">
<p style="margin:3px 0;"><span style="color:#00BCD4;">⬤</span> Đông Nam Á / Nam Á</p>
<p style="margin:3px 0;"><span style="color:#FF0000;">⬤</span> <b>VIỆT NAM 🇻🇳</b></p>
<p style="margin:0;color:#8b949e;font-size:10px;">Tổng: 35 nodes | 53 cạnh<br>Bao gồm 5 bài Việt Nam<br>Ngày tạo: 24/03/2026</p>
</div>
"""

html = html.replace("<body>", f"<body>\n{inject}")
html = html.replace("<title>", "<title>Mạng lưới trích dẫn — 11 nghiên cứu lo âu/trầm cảm | ")

with open(output, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"DONE: {output}")
