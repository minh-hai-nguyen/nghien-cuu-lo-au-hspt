# -*- coding: utf-8 -*-
"""
Sửa các mục RAG của bài VN bị sai (phát hiện qua audit 4 agent vs PDF gốc 16/05/2026).
Áp dụng cho cả 3 bản rag_main_index.json (master + 2 nhánh web chatbot).
"""
import json, os

RAGS = [
    "06_Scripts/questions_kg_data/rag_main_index.json",
    "tro-ly-nghien-cuu-tam-ly/web/data/rag_main_index.json",
    "tro-ly-nghien-cuu-tam-ly-light/web/data/rag_main_index.json",
]

# {id: [(old, new), ...]}
FIX = {
 "VN003": [
  ("tại các cơ sở hỗ trợ xã hội tại Huế, Việt Nam, do Pham và cộng sự (2024), khảo sát trên 273+273 thanh thiếu niên (nhóm can thiệp và nhóm đối chứng) tại các cơ sở hỗ trợ xã hội.",
   "tại các cơ sở hỗ trợ xã hội ở chín tỉnh, thành phố của Việt Nam, do Phạm Sỹ Tiến và cộng sự (2024), khảo sát trên 273 thanh thiếu niên đang sống tại cơ sở hỗ trợ xã hội và 273 thanh thiếu niên sống cùng gia đình (nhóm so sánh)."),
  ("Phát hiện quan trọng nhất : Chăm sóc CẢM XÚC (beta = –0,40, P < 0,001) giảm lo âu mạnh, trong khi chăm sóc THỂ CHẤT không có tác động có ý nghĩa. Điều này gợi ý can thiệp nên tập trung vào hỗ trợ cảm xúc (lắng nghe, chia sẻ, đồng cảm) hơn là chỉ đáp ứng nhu cầu vật chất.",
   "Phát hiện chính: chăm sóc CẢM XÚC (beta = –0,40; P < 0,001) có liên hệ nghịch và đáng kể với lo âu (chăm sóc cảm xúc tốt hơn đi kèm lo âu thấp hơn). Chăm sóc THỂ CHẤT lại có liên hệ THUẬN và có ý nghĩa với lo âu (beta = 0,28; P < 0,01). Can thiệp nên ưu tiên nâng đỡ cảm xúc (lắng nghe, chia sẻ, đồng cảm)."),
  ("• Phát hiện chăm sóc thể chất không có ý nghĩa là bất ngờ — thách thức quan niệm \"cung cấp vật chất đầy đủ = sức khỏe tâm thần tốt\".",
   "• Chăm sóc thể chất có liên hệ thuận với lo âu (beta = 0,28) — phát hiện cần được diễn giải thận trọng và kiểm chứng thêm."),
  ("Chăm sóc cảm xúc (beta = –0,40) là yếu tố bảo vệ mạnh cho SKTT thanh thiếu niên, trong khi chăm sóc thể chất không có tác động.",
   "Chăm sóc cảm xúc (beta = –0,40) là yếu tố bảo vệ mạnh cho SKTT thanh thiếu niên; chăm sóc thể chất có liên hệ thuận với lo âu (beta = 0,28)."),
  ("cỡ mẫu nhỏ (273+273), chỉ tại Huế, đối tượng đặc thù",
   "cỡ mẫu vừa (273+273), đối tượng đặc thù"),
  ("Chăm sóc thể chất | Không có ý nghĩa | >0,05 | KHÔNG tác động",
   "Chăm sóc thể chất | +0,28 | <0,01 | Liên hệ THUẬN với lo âu"),
 ],
 "VN006": [
  ("Tên công trình (VN) | Rối loạn lo âu ở học sinh trung học phổ thông chuyên",
   "Tên công trình (VN) | Rối loạn lo âu ở học sinh trung học phổ thông"),
  ("Công cụ đo: DASS-42 (40 mục về lo âu — cutoff không nêu rõ trong bản tóm tắt)",
   "Công cụ đo: DASS-42 (thang đầy đủ 42 mục — cutoff không nêu rõ trong bản tóm tắt)"),
 ],
 "VN014": [
  ("Đợt 2 — tháng 9/2023 (sau COVID, n = 4.337)",
   "Đợt 2 — tháng 1/2023 (sau COVID, n = 4.337)"),
 ],
 "VN019": [
  ("khảo sát 685 HS THPT (233 nam, 452 nữ, tuổi 15–18)",
   "khảo sát 685 HS THPT (233 nam, 452 nữ, tuổi 15–19)"),
  ("nữ 67,4% (thiên lệch giới)", "nữ 66,0% (thiên lệch giới)"),
 ],
 "VN020": [
  ("khảo sát 976 HS THPT tại 2 trường TPHCM, năm 2023.",
   "khảo sát 976 HS THPT tại 2 trường TPHCM, dữ liệu thu thập tháng 2–4/2024."),
  ("Hồi quy logistic đa biến. Phân tích yếu tố liên quan.",
   "Hồi quy Poisson đa biến (báo cáo tỷ số tỷ lệ hiện mắc PR). Phân tích yếu tố liên quan."),
 ],
 "VN022": [
  ("47% HS học thêm >3h/tuần, 15% >9h/tuần.",
   "50% HS học thêm >3h/tuần, 15% >9h/tuần."),
  ("đã xác thực VN (Truc et al., 2012)", "đã xác thực VN (Truc et al., 2015)"),
  ("5 tỉnh", "4 tỉnh"),
 ],
 "VN023": [
  ("Các yếu tố nguy cơ chính được xác định là: ở ký túc xá (PR = 1,71), năm học cuối (gần thi tốt nghiệp), tiếp xúc trực tiếp với F0/F1, kinh tế gia đình giảm sút trong dịch, và thiếu hỗ trợ tâm lý từ trường.",
   "Các yếu tố liên quan có ý nghĩa trong mô hình hồi quy hiệu chỉnh gồm: giới tính, loại hình người ở cùng (ở ký túc xá có tỷ số tỷ lệ hiện mắc hiệu chỉnh PR = 1,11), từng nhiễm hoặc tiếp xúc COVID-19, tình trạng tiêm chủng, tình trạng sức khỏe và chất lượng quan hệ xã hội."),
  ("Ở ký túc xá vs nhà | PR = 1,71 | < 0,001\nNăm học cuối | ↑ nguy cơ | < 0,01\nTiếp xúc F0/F1 | ↑ nguy cơ | < 0,001\nKinh tế gia đình giảm | ↑ nguy cơ | < 0,01",
   "Ở ký túc xá vs ở nhà | PR hiệu chỉnh = 1,11 | có ý nghĩa\nTừng nhiễm COVID-19 | PR hiệu chỉnh ≈ 1,18 | có ý nghĩa\nTiếp xúc gần ca F0 | PR hiệu chỉnh ≈ 1,23 | có ý nghĩa\nTiêm chủng / sức khỏe / quan hệ xã hội | yếu tố liên quan | có ý nghĩa"),
 ],
 "VN024": [
  ("Tỷ lệ trầm cảm (CES-D ≥ 16) là 12,2 %",
   "Tỷ lệ trầm cảm (CES-D, ngưỡng > 25/60) là 12,2 %"),
  ("Các yếu tố nguy cơ liên quan có ý nghĩa: nghiện rượu, áp lực học tập cao theo thang ESSA, gia đình mâu thuẫn, sống xa cha mẹ. Nữ có tỷ lệ cao hơn nam.",
   "Các yếu tố liên quan có ý nghĩa trong mô hình hồi quy đa biến: sống cùng người có rối loạn trầm cảm hoặc rối loạn tâm thần, bị lăng mạ – chế giễu – làm nhục, tranh cãi gay gắt với giáo viên, bị giáo viên la mắng – đe dọa."),
  ("Tỷ lệ 12,2 % thấp có thể do cut-off CES-D ≥ 16 (nghiêm ngặt) hoặc do mẫu Vĩnh Long ít chịu áp lực thi đại học hơn các tỉnh đô thị.",
   "Tỷ lệ 12,2 % có thể chịu ảnh hưởng của ngưỡng cắt CES-D > 25/60."),
 ],
 "VN026": [
  ("đăng trên Tạp chí Y học Việt Nam (VJOL) năm 2025. Khảo sát học sinh THPT tại Long An.",
   "của Trần Đức Sĩ và Mai Phương Dung (Trường Đại học Y khoa Phạm Ngọc Thạch), đăng trên Tạp chí Y Dược học Phạm Ngọc Thạch, 2025, 4(4), 31–39. Khảo sát 360 học sinh trung học phổ thông tại thành phố Tân An, tỉnh Long An."),
  ("Cắt ngang trên học sinh THPT các trường tại Long An.",
   "Cắt ngang trên 360 học sinh trung học phổ thông tại thành phố Tân An, tỉnh Long An (3 trong 6 trường)."),
  ("Tỷ lệ lo âu là 57,2 % — CAO NHẤT trong số các nghiên cứu lo âu HS THPT tại Việt Nam được hệ thống tổng hợp đến thời điểm hiện tại. Tỷ lệ này cao hơn rõ rệt so với Hà Nội (Hoa 2024: 40,6 %), Hải Phòng (39,3 %), Huế (Trần Thảo Vi 2024: ~ 30–40 %). Yếu tố nguy cơ: nữ giới, áp lực thi đại học, mâu thuẫn gia đình, lạm dụng MXH.",
   "Tỷ lệ lo âu là 57,2 % (tỷ lệ hiệu chỉnh theo cụm; tỷ lệ thô 61,9 %). Đây là mức cao nhưng KHÔNG phải cao nhất trong các nghiên cứu HS THPT Việt Nam — chính bài báo dẫn một số nghiên cứu trong nước có tỷ lệ lo âu cao hơn. Yếu tố liên quan có ý nghĩa: loại hình trường (chuyên/tư thục) và khối lớp 12; khác biệt theo giới không đạt ý nghĩa thống kê."),
  ("Tỷ lệ 57,2 % là CAO BẤT THƯỜNG so với các vùng khác — cần giải thích.",
   "Tỷ lệ 57,2 % ở mức cao nhưng nằm trong khoảng dao động của các nghiên cứu trong nước."),
 ],
 "VN027": [
  ("của Dinh và cộng sự (2021), đăng trên ResearchGate năm 2021.",
   "của Đinh Thị Hồng Vân, Đỗ Thị Lê Hằng và Phan Thị Mai Hương (2021), đăng trên tạp chí Psychology and Education, 58(1), 883–894."),
  ("Cắt ngang trên HS THCS tại Việt Nam. Sử dụng bảng hỏi về yếu tố trường (áp lực học tập, mối quan hệ với GV và bạn bè, môi trường lớp học, ...) và thang đo lo âu. Phân tích hồi quy đa biến.",
   "Cắt ngang trên 749 học sinh trung học cơ sở tại Hà Nội và Thành phố Hồ Chí Minh (4 trường). Sử dụng bộ câu hỏi 6 nhóm yếu tố trường học do nhóm tác giả tự xây dựng. Phân tích hồi quy logistic nhị phân."),
  ("Các yếu tố trường học có liên quan có ý nghĩa với lo âu HS THCS Việt Nam: áp lực học tập, kỳ vọng từ GV, mối quan hệ tiêu cực với bạn bè (bạo lực học đường, bắt nạt), khối lượng bài tập về nhà nhiều, ít thời gian nghỉ ngơi.",
   "Yếu tố được nhiều học sinh ghi nhận gây lo âu nhất là điểm số và thi cử (91,1 %) — cao nhất trong sáu nhóm yếu tố trường học khảo sát; các nhóm khác gồm yếu tố môn học, an toàn trường học, giao tiếp với giáo viên, quan hệ bạn bè và cơ sở vật chất."),
  ("Bài đăng trên ResearchGate (preprint hoặc tạp chí trong nước), không phải tạp chí Q1.",
   "Bài đăng trên tạp chí quốc tế Psychology and Education."),
  ("không có thông tin chi tiết về cỡ mẫu và công cụ trong bài tóm lược.",
   "cỡ mẫu n = 749 và công cụ đã được xác nhận từ bản gốc."),
 ],
 "VN029": [
  ("của Truc Thanh Thai, Minh Cuong Duong và cộng sự (2025)",
   "của Truc Thanh Thai và cộng sự (2026; Minh Cuong Duong là tác giả liên hệ)"),
  ("YBRS (Youth Behavior Risk Scale, HRB)",
   "YRBS (Youth Risk Behavior Surveillance, đo HRB)"),
 ],
}

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
        ok = 0
        for old, new in repls:
            if old in t:
                t = t.replace(old, new)
                ok += 1
            else:
                print(f"  !! {pid}: KHÔNG khớp -> {old[:55]}...")
        it["text"] = t
        print(f"  {pid}: {ok}/{len(repls)} sửa")
    json.dump(rag, open(ragpath, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"  -> đã lưu\n")
print("[DONE]")
