# -*- coding: utf-8 -*-
"""
Sửa các mục RAG bài QT bị sai (phát hiện qua audit 8 agent vs PDF gốc, 16/05/2026).
Áp dụng cho cả 3 bản rag_main_index.json (master + 2 nhánh web chatbot).
Nguyên tắc: chỉ sửa lỗi bịa đặt / sai dữ liệu rõ ràng đã được agent xác minh với PDF.
"""
import json, os, io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

RAGS = [
    "06_Scripts/questions_kg_data/rag_main_index.json",
    "tro-ly-nghien-cuu-tam-ly/web/data/rag_main_index.json",
    "tro-ly-nghien-cuu-tam-ly-light/web/data/rag_main_index.json",
]

# {id: [(old, new), ...]}
FIX = {
 "QT002": [
  ("thanh thiếu niên bị lưu ban có tỷ lệ lo âu 51,7%, gấp đôi nhóm có kết quả bình thường.",
   "thanh thiếu niên bị lưu ban có tỷ lệ lo âu 65,5%, cao hơn rõ rệt so với nhóm có kết quả bình thường."),
  ("Thực tế, sự chênh lệch giới trong lo âu ở thanh thiếu niên đã được xác nhận trong hàng chục nghiên cứu đa quốc gia (Daly, 2022).",
   "Thực tế, sự chênh lệch giới trong lo âu ở thanh thiếu niên đã được ghi nhận trong nhiều nghiên cứu quốc tế."),
 ],
 "QT007": [
  ("nữ chiếm 58% nhóm căng thẳng so với 48% nhóm không căng thẳng",
   "nữ chiếm 58% nhóm có triệu chứng trầm cảm hoặc lo âu so với 48% nhóm không có triệu chứng"),
  ("66% học sinh có căng thẳng tâm thần có rối loạn giấc ngủ, so với 17% nhóm không căng thẳng.",
   "66% học sinh có triệu chứng trầm cảm hoặc lo âu có rối loạn giấc ngủ, so với 17% nhóm không có triệu chứng."),
  ("8,3% học sinh căng thẳng có triệu chứng IGD so với 1,0% nhóm không căng thẳng.",
   "8,3% học sinh có triệu chứng trầm cảm hoặc lo âu có triệu chứng IGD so với 1,0% nhóm không có triệu chứng."),
 ],
 "QT013": [
  ("của Yin SZD, Low MK, Mishu MP (2025, University of York UK)",
   "của Yin SZD, Low MK, Mishu MP (2025, University College London, Anh)"),
 ],
 "QT014": [
  ("Anderson, T.L., Valiauga, R., Tallo, C., Hong, C.B., Manoranjithan, S., Domingo, C., Pilitsis, J.G. và cộng sự (2025)",
   "Anderson, T.L., Valiauga, R., Tallo, C., Hong, C.B., Manoranjithan, S., Domingo, C. và cộng sự (2025)"),
  ("Tác giả từ Mỹ (Rowan University School of Osteopathic Medicine).",
   "Tác giả từ Mỹ (Đại học Connecticut và Đại học Loyola Chicago)."),
  ("Là tổng quan tường thuật 61 bài về yếu tố góp phần gia tăng lo âu VTN.",
   "Là tổng quan tường thuật về các yếu tố góp phần gia tăng lo âu VTN."),
  ("• Tỷ lệ lo âu VTN 13–18 tuổi là 31,9 % (Daly 2022 — trích dẫn) — Thế hệ Z CAO NHẤT so với 3 thế hệ trước đó.",
   "• Lo âu ở Thế hệ Z (sinh 1997–2012) được mô tả là gia tăng so với các thế hệ trước."),
  ("Điểm mạnh: Tổng quan 61 bài bao quát nhiều yếu tố",
   "Điểm mạnh: Tổng quan tường thuật bao quát nhiều yếu tố"),
 ],
 "QT015": [
  ("trên 9.831 học sinh THCS và THPT tại Suzhou, miền Đông Trung Quốc. Sử dụng PHQ-9 (Patient Health Questionnaire-9 mục) đánh giá trầm cảm.",
   "trên 9.831 học sinh THCS và THPT (mẫu phân tích; tổng số thu được là 9.920) tại Nantong, miền Đông Trung Quốc. Sử dụng thang CES-D (Center for Epidemiologic Studies Depression Scale) đánh giá trầm cảm."),
  ('Khảo sát cắt ngang lặp lại 2019, 2021, 2023 (trước, trong và sau COVID) trên 9.831 HS tại Suzhou. Sử dụng PHQ-9 với 2 ngưỡng cắt: ≥5 "có thể có trầm cảm" (possible), ≥10 "chắc chắn có trầm cảm" (definite).',
   'Khảo sát cắt ngang lặp lại hằng năm trong giai đoạn 2019–2023 (trước, trong và sau COVID) trên học sinh tại Nantong. Sử dụng thang CES-D với 2 mức: "có thể có trầm cảm" (possible) và "chắc chắn có trầm cảm" (definite).'),
  ("Điểm mạnh: Cỡ mẫu rất lớn (n = 9.831), so sánh 3 thời điểm 2019–2023",
   "Điểm mạnh: Cỡ mẫu rất lớn (n = 9.831 phân tích), so sánh nhiều năm 2019–2023"),
  ("Hạn chế: (1) Chỉ tại Suzhou (miền Đông, đô thị phát triển)",
   "Hạn chế: (1) Chỉ tại Nantong (miền Đông, đô thị phát triển)"),
  ("(2) PHQ-9 là sàng lọc, không phải chẩn đoán lâm sàng;",
   "(2) CES-D là thang sàng lọc, không phải chẩn đoán lâm sàng;"),
  ("Gia đình đơn thân | 1,434 | — | Nguy cơ",
   "Gia đình đơn thân | 1,359 | — | Nguy cơ"),
 ],
 "QT016": [
  ("DOI 10.1016/S2772-3682(25)00003-4.",
   "DOI 10.1016/j.lansea.2025.100532."),
 ],
 "QT017": [
  ("(YAFS — Young Adult Fertility and Sexuality Survey)",
   "(YAFS — Young Adult Fertility and Sexuality Study)"),
  ("Nhóm nghèo nhất | 12,3 % | 25,1 % | × 2,0",
   "Nhóm nghèo nhất | 10,6 % | 25,1 % | × 2,4"),
  ("Không đi học | 8,5 % | 26,5 % | × 3,1",
   "Học vấn tiểu học trở xuống | 10,8 % | 26,5 % | × 2,5"),
 ],
 "QT022": [
  ("Cắt ngang T1 | Lo âu CAS-8 | 0,72/giờ | <0,001 | +3,60 điểm — ĐÁNG KỂ",
   "Cắt ngang T1 | Lo âu CAS-8 | 0,79/giờ | <0,001 | +3,95 điểm — ĐÁNG KỂ"),
 ],
 "QT023": [
  ("ADHD | 29,8 | 25,3 | 0,90 (0,77–1,06) | — Ổn định",
   "ADHD | 29,8 | 25,3 | 0,97 (0,77–1,22) | — Ổn định"),
 ],
 "QT024": [
  ("từ nhiều nguồn trong khu vực WHO châu Âu (53 quốc gia).",
   "từ nhiều nguồn trong khu vực WHO châu Âu."),
  ("Tổng quan tài liệu cho thấy 9 triệu VTN châu Âu (10–19 tuổi) sống với rối loạn SKTT. Lo âu + trầm cảm chiếm >50% tổng ca.",
   "Tổng quan tài liệu cho thấy gánh nặng rối loạn SKTT ở trẻ em và thanh thiếu niên trong khu vực là rất lớn; trên phạm vi toàn cầu, ước tính khoảng 300 triệu trong số 2,5 tỉ người 5–24 tuổi có một rối loạn sức khỏe tâm thần."),
  ("Con số 9 triệu VTN châu Âu cho thấy quy mô vấn đề. So sánh: ASEAN 80,4 triệu ca rối loạn (GBD 2025).",
   "Quy mô gánh nặng SKTT ở trẻ em và thanh thiếu niên châu Âu là rất lớn. So sánh: ASEAN 80,4 triệu ca rối loạn (GBD 2025)."),
  ("*Lo âu + trầm cảm > 50% ca rối loạn* — xác nhận lo âu là vấn đề hàng đầu, phù hợp JAACAP 2024 và tất cả NC VN.",
   "*Lo âu và trầm cảm là nhóm rối loạn phổ biến* — xác nhận lo âu là vấn đề hàng đầu, phù hợp JAACAP 2024 và tất cả NC VN."),
  ("*Chỉ 5/53 nước có tổ chức đa ngành cho SKTT* — ngay cả châu Âu cũng thiếu.",
   "*Hợp tác đa ngành cho SKTT còn hạn chế* — ngay cả châu Âu cũng thiếu."),
  ("Dữ liệu tổng hợp WHO châu Âu, cho thấy 9 triệu VTN đang sống với rối loạn SKTT và lo âu chiếm > 50% tổng ca, gợi ý rằng",
   "Dữ liệu tổng hợp WHO châu Âu cho thấy gánh nặng rối loạn SKTT ở trẻ em và thanh thiếu niên rất lớn, trong đó lo âu và trầm cảm là nhóm phổ biến, gợi ý rằng"),
  ("Lancet Regional Health Europe Q1 IF = 15. Phạm vi 53 quốc gia.",
   "Lancet Regional Health Europe Q1 IF = 15. Phạm vi khu vực WHO châu Âu."),
  ("Đánh giá: ⭐⭐⭐⭐⭐ Rất cao. Lancet Q1 IF = 15, 53 quốc gia, hướng chính sách.",
   "Đánh giá: ⭐⭐⭐⭐⭐ Rất cao. Lancet Q1 IF = 15, tổng quan chính sách khu vực châu Âu."),
  ("Tỷ lệ | 9 triệu VTN có RLSKTT | Ưu tiên quốc gia | VN: chưa có dữ liệu tổng",
   "Tỷ lệ | Gánh nặng RLSKTT lớn ở trẻ em/VTN | Ưu tiên quốc gia | VN: chưa có dữ liệu tổng"),
  ("Lo âu + trầm cảm | >50% tổng ca rối loạn | Can thiệp tập trung | Phù hợp — lo âu hàng đầu",
   "Lo âu + trầm cảm | Nhóm rối loạn phổ biến | Can thiệp tập trung | Phù hợp — lo âu hàng đầu"),
  ("Dịch vụ | 16/27 nước thiếu | Phát triển cộng đồng | VN: 8,4% tiếp cận",
   "Dịch vụ | Nhiều nước còn thiếu | Phát triển cộng đồng | VN: 8,4% tiếp cận"),
 ],
 "QT026": [
  ("do Carl Baker, House of Commons Library, UK Parliament (2025). 46 trang. Nguồn chính thức Quốc hội Anh, dựa trên NHS Mental Health Survey 2025 và nhiều khảo sát quốc gia.",
   "do Carl Baker và Esme Kirk-Wade, House of Commons Library, UK Parliament (2024). 46 trang. Nguồn chính thức Quốc hội Anh, tổng hợp dữ liệu NHS Digital và nhiều khảo sát quốc gia."),
  ("Chỉ số | 2014 | 2025 | Thay đổi",
   "Chỉ số | Số liệu (báo cáo House of Commons Library, 2024)"),
  ("Rối loạn TT phổ biến (16–24 tuổi) | 18,9% | 25,8% | +36% (11 năm)",
   "Trẻ 8–16 tuổi có rối loạn TT 'có khả năng' | 12% (2017) lên 20% (2023)"),
  ("Nữ 16–24 | 36,1% | Gấp 2,2 lần nam",
   "Thanh niên 17–19 tuổi có rối loạn TT 'có khả năng' | 10% (2017) lên 23% (2023)"),
  ("Nam 16–24 | 16,3%",
   "Người trưởng thành có rối loạn tâm thần phổ biến | khoảng 1/6 (Khảo sát Bệnh tật Tâm thần Người trưởng thành 2014)"),
  ("\nVTN 17–19 đủ tiêu chuẩn rối loạn TT | ~25% | 1/4 VTN\nTự hại suốt đời (mọi tuổi) | 10,3% | Tăng gấp 4 từ 2000\nNữ 16–24 tự hại | 31,7% | Rất cao",
   "\nNữ thanh thiếu niên có tỷ lệ rối loạn TT cao hơn nam"),
  ("*Nữ — nhóm đặc biệt dễ tổn thương.* 36,1% nữ 16–24 mắc rối loạn TT (gấp 2,2 lần nam) — cần can thiệp nhạy cảm giới.",
   "*Nữ — nhóm đặc biệt dễ tổn thương.* Báo cáo ghi nhận nữ thanh thiếu niên có tỷ lệ rối loạn TT cao hơn nam — cần can thiệp nhạy cảm giới."),
  ("*Tự hại tăng gấp 4 từ 2000.* Xu hướng đáng báo động, phù hợp với dữ liệu tự hại ở VN (Danh Lâm: 10%).",
   "*Tự hại là vấn đề đáng lưu tâm.* Xu hướng đáng báo động, phù hợp với dữ liệu tự hại ở VN (Danh Lâm: 10%)."),
  ("Dữ liệu NHS Anh cho thấy 25,8% thanh niên 16–24 mắc rối loạn tâm thần phổ biến — tăng 36% trong 11 năm — với nữ bị ảnh hưởng gấp đôi nam và tự hại tăng gấp 4.",
   "Dữ liệu tổng hợp của House of Commons Library cho thấy tỷ lệ rối loạn tâm thần 'có khả năng' ở trẻ 8–16 tuổi tăng từ 12% (2017) lên 20% (2023) và ở thanh niên 17–19 tuổi tăng từ 10% lên 23%, với nữ bị ảnh hưởng nặng hơn."),
  ("Đối chiếu liên bài : Xu hướng tăng SKTT ở VTN Anh (+36% trong 11 năm) nhất quán với Mỹ, Na Uy, Ireland. Nữ bị ảnh hưởng nặng hơn ở TẤT CẢ các nước. Đáng chú ý, tự hại ở nữ 16–24 Anh (31,7%) rất cao — so sánh Danh Lâm 2022 VN (tự làm đau 10%, cố tự tử 1,4%).",
   "Đối chiếu liên bài : Xu hướng tăng SKTT ở VTN Anh nhất quán với Mỹ, Na Uy, Ireland. Nữ bị ảnh hưởng nặng hơn ở TẤT CẢ các nước — so sánh Danh Lâm 2022 VN (tự làm đau 10%, cố tự tử 1,4%)."),
  ("Anh (NHS) | 2014–2025 | +36% rối loạn TT 16–24 | Bài này",
   "Anh (HoC Library) | 2017–2023 | Trẻ 8–16: 12% lên 20% | Bài này"),
 ],
 "QT028": [
  ("CBT được xác nhận hiệu quả nhất (47–66% phục hồi) bởi cả AJP 2024, BMC network meta 2025 (30 RCTs), và Zhameden 2025.",
   "CBT được xác nhận là phương pháp đã được xác lập bởi cả AJP 2024, BMC network meta 2025 (30 RCTs), và Yin và cộng sự 2025."),
  ("Đặc biệt, bằng chứng từ LMIC (nước thu nhập thấp-trung bình) rất hạn chế — 0 RCT từ Việt Nam (Zhameden 2025).",
   "Đặc biệt, bằng chứng từ LMIC (nước thu nhập thấp-trung bình) rất hạn chế — 0 RCT từ Việt Nam (Yin và cộng sự 2025)."),
  ("cho thấy CBT là phương pháp hiệu quả nhất (47–66% phục hồi) nhưng 30–40% không đáp ứng",
   "cho thấy CBT là phương pháp điều trị đã được xác lập nhưng khoảng một nửa đến hai phần ba bệnh nhân không đạt thuyên giảm hoàn toàn"),
  ("CBT cá nhân | Phục hồi 47–66%\nĐáp ứng 57–60% | RCTs nhiều | Bằng chứng mạnh nhất",
   "CBT cá nhân | Hiệu quả; nhiều bệnh nhân chưa đạt thuyên giảm hoàn toàn | RCTs nhiều | Bằng chứng mạnh nhất"),
  ("AJP 2024 (bài này) | Tổng quan | CBT 47–66% phục hồi | Chủ yếu phương Tây",
   "AJP 2024 (bài này) | Tổng quan | CBT là phương pháp đã xác lập | Chủ yếu phương Tây"),
  ("Zhameden 2025 | 6 RCTs LMIC | CBT 3/4 trầm cảm, 1/4 lo âu | GRADE rất thấp, 0 VN",
   "Yin và cộng sự 2025 | 6 RCTs LMIC | CBT 3/4 trầm cảm, 1/4 lo âu | GRADE rất thấp, 0 VN"),
 ],
 "QT029": [
  ("Đối chiếu liên bài : Ba tổng quan can thiệp (AJP 2024, BMC 2025, Zhameden 2025) đều xác nhận CBT hiệu quả nhất.",
   "Đối chiếu liên bài : Trong phân tích mạng này, Liệu pháp Chấp nhận và Cam kết (ACT) được xếp hạng cao nhất, CBT xếp thứ hai."),
  ("Tuy nhiên, Zhameden cảnh báo CBT chỉ hiệu quả 1/4 cho lo âu tại LMIC",
   "Tuy nhiên, Yin và cộng sự (2025) cảnh báo CBT chỉ hiệu quả 1/4 cho lo âu tại LMIC"),
  ("*CBT cá nhân xếp hạng 1.* Network meta Bayesian xác nhận CBT cá nhân hiệu quả nhất trong 7+ loại can thiệp — phù hợp AJP 2024.",
   "*ACT xếp hạng 1, CBT xếp hạng 2.* Phân tích mạng Bayesian so sánh 4 loại can thiệp phi dược lý (ACT, CBT, vận động thể chất, liệu pháp phơi nhiễm thực tế ảo); ACT đạt hiệu quả cao nhất (MD = −3,83; SUCRA 0,69), CBT đứng thứ hai (MD = −3,64; SUCRA 0,66)."),
  ("*CBT nhóm — lựa chọn thực tế cho VN.* Hiệu quả tương đương CBT cá nhân nhưng chi phí thấp hơn, phù hợp bối cảnh trường học VN (có thể triển khai cho lớp 20–30 HS).",
   "*CBT — lựa chọn thực tế cho trường học VN.* CBT có hiệu quả tốt, đứng thứ hai sau ACT, phù hợp triển khai trong bối cảnh trường học VN."),
  ("Dữ liệu từ phân tích tổng hợp mạng Bayesian trên 30 RCTs (1.711 trẻ), cho thấy CBT cá nhân hiệu quả nhất, CBT nhóm xếp thứ 2–3 (tương đương nhưng chi phí thấp hơn), gợi ý rằng CBT nhóm tại trường học là lựa chọn khả thi nhất cho RCT đầu tiên tại Việt Nam.",
   "Dữ liệu từ phân tích tổng hợp mạng Bayesian trên 30 RCTs (1.711 trẻ), cho thấy ACT đạt hiệu quả cao nhất và CBT đứng thứ hai trong bốn loại can thiệp phi dược lý, gợi ý rằng CBT (và ACT) tại trường học là lựa chọn khả thi cho RCT đầu tiên tại Việt Nam."),
  ("1 | CBT cá nhân | Hiệu quả nhất | Bằng chứng mạnh\n2–3 | CBT nhóm | Tương đương cá nhân | Chi phí thấp hơn\n3–4 | SSRI (thuốc) | Hiệu quả | Tác dụng phụ ở trẻ\n4–5 | CBT + SSRI | Tốt nhưng ít RCT | Cần thêm NC\n5+ | Tâm lý giáo dục | Hiệu quả thấp hơn | An toàn\nThư giãn | Hiệu quả hạn chế\nChờ đợi (waitlist) | Tham chiếu | Không can thiệp",
   "1 | ACT (Chấp nhận và Cam kết) | MD −3,83; SUCRA 0,69 | Hiệu quả cao nhất\n2 | CBT | MD −3,64; SUCRA 0,66 | Đứng thứ hai\n3 | Vận động thể chất (PE) | — | Có hiệu quả\n4 | Phơi nhiễm thực tế ảo (VRET) | — | Có hiệu quả"),
  ("BMC 2025 (bài này) | 30 RCTs, 1.711 trẻ | CBT cá nhân tốt nhất | NMA Bayesian — nâng cao nhất\nAJP 2024 (QT28) | Tổng quan tường thuật | CBT 47–66% phục hồi | Không xếp hạng\nZhameden 2025 | 6 RCTs LMIC | CBT 3/4 dep, 1/4 anx | GRADE rất thấp, 0 VN",
   "BMC 2025 (bài này) | 30 RCTs, 1.711 trẻ | ACT tốt nhất, CBT thứ hai | NMA Bayesian — nâng cao nhất\nAJP 2024 (QT28) | Tổng quan tường thuật | CBT là phương pháp đã xác lập | Không xếp hạng\nYin và cộng sự 2025 | 6 RCTs LMIC | CBT 3/4 dep, 1/4 anx | GRADE rất thấp, 0 VN"),
 ],
 "QT034": [
  ("Buồn bã | Giảm | TĂNG (28,2%) | Khoảng cách MỞ RỘNG",
   "Buồn bã | Giảm | TĂNG (nhóm thu nhập cao 28,2%; nhóm thu nhập thấp 46,8%) | Khoảng cách MỞ RỘNG"),
  ("Ý tưởng tự tử | Giảm | TĂNG (13,9%) | Thu nhập thấp tệ hơn",
   "Ý tưởng tự tử | Giảm | TĂNG (nhóm thu nhập cao 13,9%; nhóm thu nhập thấp 31,7%) | Thu nhập thấp tệ hơn"),
  ("16 năm", "17 năm"),
 ],
 "QT040": [
  ("đăng trên Journal of Medical Internet Research (JMIR — Q1, IF ≈ 7,4). DOI 10.2196/preprints.67067.",
   "ở dạng bản thảo tiền in (preprint) nộp cho Journal of Medical Internet Research (JMIR). DOI 10.2196/preprints.67067."),
 ],
 "QT044": [
  ("Can thiệp resilience tại trường CÓ HIỆU QUẢ tăng resilience và giảm triệu chứng SKTT ở trẻ em/VTN.",
   "Can thiệp resilience tại trường CÓ HIỆU QUẢ tăng resilience ở trẻ em/VTN; phân tích tổng hợp chỉ gộp kết quả đầu ra về resilience (hiệu lực nhỏ, SMD ≈ 0,17)."),
 ],
 "QT045": [
  ("Can thiệp 8 module, theo dõi đến tháng 3 sau khởi động. Đo lường: SIAS (lo âu xã hội chính), PHQ-9 (trầm cảm).",
   "Can thiệp 10 module, theo dõi đến tháng 3 sau khởi động. Đo lường: LSAS (thang lo âu xã hội Liebowitz — kết quả chính), SPIN (lo âu xã hội — kết quả phụ), PHQ-9 (trầm cảm)."),
  ("sàng lọc HS subthreshold SAD bằng SIAS-17",
   "sàng lọc HS subthreshold SAD bằng SPIN"),
 ],
 "QT057": [
  ("Tác giả | (xem PDF)", "Tác giả | Wenxuan Gong"),
 ],
 "QT058": [
  ("Nơi công bố | Translational Psychiatry (2022)",
   "Nơi công bố | Translational Psychiatry (2023)"),
 ],
 "QT061": [
  ("• Mẫu: Co-design phase: tham gia ý kiến từ chuyên gia + VTN. Sau đó RCT đang chuẩn bị tuyển 489 VTN 12-17 tuổi.",
   "• Mẫu: Giai đoạn đồng thiết kế (co-design) với 36 vị thành niên, 15 phụ huynh và 32 chuyên gia sức khỏe tâm thần. Một RCT đầy đủ đã được lên kế hoạch (chưa nêu cỡ mẫu)."),
  ("• Công cụ: ClearlyMe app (37 brief lessons CBT) + thang đo trầm cảm + lo âu (chi tiết trong RCT đang lên kế hoạch)",
   "• Công cụ: ứng dụng ClearlyMe (app CBT tự hướng dẫn) + thang đo trầm cảm + lo âu (chi tiết trong RCT đang lên kế hoạch)"),
  ("• ClearlyMe = app self-guided CBT 37 BÀI HỌC NGẮN",
   "• ClearlyMe = app CBT tự hướng dẫn; nội dung gồm 9 bộ sưu tập (Collections), mỗi bộ 3–5 hoạt động"),
  ("• Chia 5 modules chính + 32 mini-lessons",
   "• Quá trình đồng thiết kế rút ra 5 chủ đề chính từ phỏng vấn nhóm"),
  ("• RCT đang lên kế hoạch (target 489 VTN)",
   "• RCT đầy đủ đã được lên kế hoạch (chưa công bố cỡ mẫu)"),
  ('• Cấu trúc 37 bài ngắn — phù hợp tâm lý "scrolling" của VTN',
   '• Cấu trúc gồm nhiều hoạt động ngắn — phù hợp tâm lý "scrolling" của VTN'),
  ("• Theo dõi kết quả RCT chính của ClearlyMe (n=489) khi công bố.",
   "• Theo dõi kết quả RCT chính của ClearlyMe khi công bố."),
  ("5 module chính + 32 mini-lessons là cấu trúc đáng tham khảo.",
   "Năm chủ đề chính rút ra từ quá trình đồng thiết kế là điểm đáng tham khảo."),
  ("Key fact cho RAG: ClearlyMe (Úc) co-design 37 bài ngắn, app CBT cho VTN 12-17 trầm cảm + lo âu. Black Dog Institute UNSW.",
   "Key fact cho RAG: ClearlyMe (Úc) — app CBT cho VTN 12-17 trầm cảm + lo âu, phát triển theo quy trình đồng thiết kế. Black Dog Institute UNSW."),
  ("Nơi công bố | The Cognitive Behaviour Therapist · Cambridge University Press (2024)",
   "Nơi công bố | The Cognitive Behaviour Therapist · Cambridge University Press (2022)"),
  ("Năm | 2024", "Năm | 2022"),
  ("DOI | 10.1017/S1754470X24000345", "DOI | 10.1017/S1754470X22000095"),
 ],
 "QT063": [
  ("• Công cụ: App-based CBT-I (cognitive behavioral therapy for insomnia) + ISI (Insomnia Severity Index) + PHQ-9 / BDI / SCID cho MDD",
   "• Công cụ: App-based CBT-I (cognitive behavioral therapy for insomnia) + ISI (Insomnia Severity Index) + PHQ-9 + MINI (Mini-International Neuropsychiatric Interview) cho MDD"),
  ("• Cấu trúc CBT-I 5 module: psychoeducation → sleep restriction → stimulus control → cognitive restructuring → relaxation + sleep hygiene",
   "• Cấu trúc CBT-I 6 module: tổng quan về giấc ngủ → hạn chế giấc ngủ → kiểm soát kích thích → tái cấu trúc nhận thức → thời gian lo lắng có cấu trúc → phòng ngừa tái phát"),
  ("• Mẫu: Cần verify từ PDF — RCT lớn, đối tượng youth có insomnia + subclinical depression",
   "• Mẫu: 708 thanh thiếu niên/thanh niên 15–25 tuổi (mainland China + Hong Kong) có mất ngủ + triệu chứng trầm cảm dưới ngưỡng; phân nhóm 1:1 CBT-I (354) vs giáo dục sức khỏe (354)"),
  ("Adapt cấu trúc 5 module CBT-I sang CBT for anxiety",
   "Adapt cấu trúc 6 module CBT-I sang CBT for anxiety"),
  ("Key fact cho RAG: Chen 2025 PLOS Med Hong Kong: app CBT-I 5 module ngăn MDD",
   "Key fact cho RAG: Chen 2025 PLOS Med Hong Kong: app CBT-I 6 module ngăn MDD"),
 ],
 "QT064": [
  ("~200 vị thành niên 10-22 tuổi", "~200 vị thành niên 10-19 tuổi"),
 ],
 "QT065": [
  ("• Cấu trúc EACP: 25 buổi học sinh + 16 buổi cha mẹ + thành phần giáo viên (giảm từ bản gốc 34 buổi)",
   "• Cấu trúc EACP: 25 buổi học sinh + 12 buổi cha mẹ + thành phần giáo viên (giảm từ bản gốc 34 buổi)"),
  ("Key fact cho RAG: EACP RCT 40 trường, n=709 HS lớp 7 Mỹ. 25 buổi HS + 16 buổi cha mẹ.",
   "Key fact cho RAG: EACP RCT 40 trường, n=709 HS lớp 7 Mỹ. 25 buổi HS + 12 buổi cha mẹ."),
 ],
 "QT066": [
  ("• Bias bối cảnh: chủ yếu Ireland (Jigsaw) — không đại diện toàn cầu",
   "• Bias bối cảnh: 15 nghiên cứu chủ yếu từ nước thu nhập cao (5 Úc, 4 Mỹ, 3 Canada, 2 Anh, 1 Hà Lan) — không đại diện toàn cầu"),
  ("Lead author email | rachel.murphy@ucd.ie (Rachel Murphy, University College Dublin)",
   "Lead author email | rachel.murphy12@ucdconnect.ie (Rachel Murphy, University College Dublin)"),
 ],
 "QT067": [
  ("Stress làm giảm chất lượng giấc ngủ; thanh thiếu niên ngủ < 7h/đêm có nguy cơ trầm cảm cao gấp 2-3 lần.",
   "Stress làm giảm chất lượng giấc ngủ ở thanh thiếu niên (ví dụ Noland và cộng sự ghi nhận 42% học sinh báo cáo stress cản trở giấc ngủ)."),
  ("Liên kết hai chiều: stress → sử dụng chất, và sử dụng chất → stress tăng.",
   "Stress làm tăng xu hướng sử dụng chất (bài gốc nêu chiều stress → sử dụng chất)."),
  ("Mối quan hệ hình chữ U ngược: stress thấp → động lực thấp; stress vừa → tối ưu; stress cao → giảm hiệu suất.",
   "Học sinh có mức stress cảm nhận cao hơn có xu hướng kết quả học tập thấp hơn (quan hệ một chiều theo bài gốc)."),
  ("Stress quá mức làm giảm working memory và khả năng giải quyết vấn đề.",
   "Stress quá mức gây khó tập trung và giảm mức độ gắn kết học tập."),
  ("Open access (CC BY 4.0) — tiếp cận miễn phí, được trích dẫn rộng rãi (>1500 lần tính đến 2025).",
   "Open access (CC BY 4.0) — tiếp cận miễn phí, được trích dẫn rộng rãi."),
  ("Dữ liệu trích dẫn từ ~80 nghiên cứu trước đó, phần lớn từ USA, UK, Úc.",
   "Dữ liệu trích dẫn từ các nghiên cứu trước đó, phần lớn từ USA, UK, Úc."),
  ("Sinh viên đại học báo cáo mức stress học đường cao tương đương stress công việc của người trưởng thành.",
   "Sinh viên đại học báo cáo mức stress học đường ở mức cao."),
  ("Mối quan hệ U-ngược giữa stress và performance → có thể giải thích vì sao một số HS THPT có stress cao nhưng vẫn học giỏi (cho tới ngưỡng tới hạn).",
   "Quan hệ giữa stress và kết quả học tập → cần thận trọng khi diễn giải trường hợp HS có stress cao."),
 ],
}

total_ok = total_miss = 0
for ragpath in RAGS:
    if not os.path.exists(ragpath):
        print(f"(bỏ qua, không có) {ragpath}")
        continue
    rag = json.load(open(ragpath, encoding="utf-8"))
    ent = {(it.get("meta", {}).get("id") or it.get("id", "")): it for it in rag["entries"]}
    print(f"=== {ragpath} ===")
    for pid, repls in FIX.items():
        it = ent.get(pid)
        if it is None:
            print(f"  {pid}: KHÔNG có trong index")
            continue
        t = it.get("text", "")
        ok = miss = 0
        for old, new in repls:
            if old in t:
                t = t.replace(old, new)
                ok += 1
            else:
                miss += 1
                print(f"  !! {pid}: KHÔNG khớp -> {old[:60]}...")
        it["text"] = t
        total_ok += ok
        total_miss += miss
        print(f"  {pid}: {ok}/{len(repls)} sửa")
    json.dump(rag, open(ragpath, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"  -> đã lưu\n")
print(f"[DONE] tổng: {total_ok} sửa OK, {total_miss} không khớp")
