"""Build 00_Thuc trang va khung tap huan-v2.docx
- COPY file goc
- INSERT binh luan XANH sau cac bang so lieu thuc trang key
- Bang 3.17, 3.18, 3.19, 3.20, 3.21, 3.22, 3.24, 3.26, 3.28, 3.30, 3.32,
  3.34, 3.36, 3.38, 3.42, 3.43, 3.45

Format: chu mau XANH cho phan binh luan moi.
"""
import sys, io
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import shutil

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SRC = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/01_Bao-cao/bang-so-lieu-binh-luan/00_Thực trạng và khung tập huấn.docx')
OUT = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/01_Bao-cao/bang-so-lieu-binh-luan/00_Thực trạng và khung tập huấn-v2.docx')


def make_blue_paragraph(text):
    """Create blue paragraph XML element."""
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    jc = OxmlElement('w:jc'); jc.set(qn('w:val'), 'both'); pPr.append(jc)
    ind = OxmlElement('w:ind'); ind.set(qn('w:firstLine'), '567'); pPr.append(ind)
    p.append(pPr)

    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    rPr.append(rFonts)
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '26'); rPr.append(sz)
    szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), '26'); rPr.append(szCs)
    color = OxmlElement('w:color'); color.set(qn('w:val'), '0070C0'); rPr.append(color)
    r.append(rPr)

    t = OxmlElement('w:t')
    t.text = text
    t.set(qn('xml:space'), 'preserve')
    r.append(t)
    p.append(r)
    return p


def make_blue_label():
    """Create label '[Bình luận bổ sung — chữ xanh]'."""
    return make_blue_paragraph('[BÌNH LUẬN BỔ SUNG]')


# ============================================================
# Bình luận cho từng bảng — text dài nên đoạn đôi paragraph
# ============================================================
COMMENTS = {
    # Table index → list of paragraphs (each is a string)
    # Tables theo verified mapping
    16: [  # Bảng 3.17 RLLATQ biểu hiện
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.17 — chữ xanh]',
        'Bảng 3.17 cho thấy 7 mục RCADS đo lo âu lan tỏa có ĐTB từ 45,86 (RCADS35: '
        'lo lắng tương lai) đến 64,28 (RCADS4: nghĩ rằng đã không hoàn thành) trên '
        'thang 0–100. Mục có ĐTB CAO NHẤT (RCADS4 = 64,28) phản ánh đặc thù lứa '
        'tuổi vị thành niên — tự đánh giá nghiêm khắc và xu hướng tự trách. Mục '
        'thứ hai (RCADS13 = 59,62) "lo lắng điều gì đó tồi tệ sẽ xảy ra" gợi ý '
        'lo âu mang tính dự kiến (anticipatory anxiety). Phù hợp với khung phát '
        'triển Steinberg (2014) về sự nhạy cảm tăng với đánh giá ở tuổi dậy thì.',
    ],
    17: [  # Bảng 3.18 RLLACL biểu hiện
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.18 — chữ xanh]',
        'Bảng 3.18 cho thấy 4 mục RCADS đo lo âu chia ly có ĐTB chỉ từ 21,52 đến '
        '27,88 — ĐÁNG KỂ THẤP HƠN so với lo âu lan tỏa (45,86–64,28) và lo âu xã '
        'hội (42,09–56,98). Mục cao nhất (RCADS46 = 27,88) "sợ phải ở xa nhà qua '
        'đêm" và mục thấp nhất (RCADS45 = 21,52) "lo lắng khi đi ngủ ban đêm". '
        'Pattern này phù hợp với khung phát triển: lo âu chia ly là rối loạn KHỞI '
        'PHÁT SỚM (DSM-5), thường giảm khi trẻ trưởng thành và tự lập — phù hợp '
        'với xu hướng giảm theo khối lớp ở Bảng 3.20 (32,13 → 19,46 từ lớp 6 đến '
        'lớp 9).',
    ],
    18: [  # Bảng 3.19 RLLAXH biểu hiện
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.19 — chữ xanh]',
        'Bảng 3.19 cho thấy 4 mục RCADS đo lo âu xã hội có ĐTB từ 42,09 đến '
        '56,98. Mục CAO NHẤT (RCADS32 = 56,98) "lo lắng về việc người khác nghĩ '
        'gì về mình" — biểu hiện CỐT LÕI của social anxiety theo DSM-5. Mục thứ '
        'hai (RCADS43 = 49,26) "sợ làm trò cười trước mặt người khác" và mục '
        'thấp nhất (RCADS20 = 42,09) "lo trông ngốc nghếch". Đây là pattern phù '
        'hợp với mô hình Clark & Wells (1995) về lo âu xã hội: tâm điểm là MỐI '
        'BẬN TÂM với đánh giá tiêu cực từ người khác.',
    ],
    19: [  # Bảng 3.20 giới × khối
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.20 — chữ xanh]',
        'Bảng 3.20 cho thấy ba phát hiện then chốt về khác biệt giới và khối lớp.',
        'Thứ nhất, khác biệt giới tính RỌ RỆT cho RLLATQ (F = 44,484; p < 0,001) '
        'và RLLAXH (F = 45,984; p < 0,001) nhưng KHÔNG ý nghĩa cho RLLACL '
        '(F = 0,246; p = 0,620). Tính Cohen d (chuẩn hóa theo SD) cho hai dạng '
        'có ý nghĩa: d RLLATQ = 0,365 và d RLLAXH = 0,370 — GẦN NHƯ BẰNG NHAU '
        '(chênh chỉ 1,5%), đều thuộc nhóm small-to-medium theo Cohen (1988). '
        'Lưu ý: F-test phụ thuộc cỡ mẫu và variance, nên F RLLAXH > F RLLATQ '
        'KHÔNG nghĩa là effect size lớn hơn.',
        'Thứ hai, RLLAC giảm đơn điệu theo khối lớp: 32,13 → 27,14 → 20,88 → '
        '19,46 từ lớp 6 đến lớp 9 — TỔNG GIẢM 12,67 điểm trong 3 năm. Phù hợp '
        'khung DSM-5: lo âu chia ly là rối loạn khởi phát sớm và giảm khi trẻ '
        'trưởng thành tự lập.',
        'Thứ ba, áp dụng ngưỡng 65/100 (T-score borderline) với giả định phân '
        'phối normal, ước lượng tỷ lệ học sinh có triệu chứng vượt ngưỡng: '
        'RLLATQ ≈ 33,85% (nam 26,88%, nữ 40,11%); RLLAXH ≈ 25,98% (nam 19,25%, '
        'nữ 32,06%); RLLAC ≈ 5,01%. Tỷ lệ RLLAXH (25,98%) GẦN với V-NAMHS '
        '(2022) chẩn đoán DSM-5 (18,45%) và Hoàng Trung Học (2025) sàng lọc '
        'DASS-21 sau COVID (25,4%) — gợi ý quy mô vấn đề ở mức KHẨN CẤP. '
        'CẢNH BÁO: ước lượng giả định normal — dữ liệu RLLA thường lệch phải, '
        'tỷ lệ thực có thể thấp hơn.',
    ],
    20: [  # Bảng 3.21 YTNC
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.21 — chữ xanh]',
        'Bảng 3.21 cho thấy ba nhóm yếu tố nguy cơ có thứ tự cường độ: ÁP LỰC '
        'HỌC TẬP (ĐTB = 51,13) > NGHIỆN ĐIỆN THOẠI (28,38) > BẮT NẠT THỂ CHẤT '
        '(13,52). Áp lực học tập là yếu tố CAO NHẤT trong nhóm nguy cơ — vượt '
        'giữa thang đo (50/100). Trong áp lực học tập, biểu hiện cao nhất là '
        'ESSA.4 "việc học và đòi hỏi sự nghiệp tương lai" (ĐTB = 58,56) — phản '
        'ánh đặc thù bối cảnh giáo dục Việt Nam: học sinh THCS đã chịu áp lực '
        'TƯƠNG LAI sớm hơn so với học sinh phương Tây.',
        'Lưu ý phương pháp: chương 3 dùng ESSA RÚT GỌN 4 mục (ESSA.3, .4, .5, '
        '.6) thay vì ESSA full 16 mục (Sun và cộng sự, 2011). Phiên bản rút '
        'gọn thiếu chiều "Despondency" (chán nản học tập) — một chiều cốt lõi '
        'theo tác giả gốc. Khuyến nghị nghiên cứu tiếp theo dùng ESSA full để '
        'so sánh trực tiếp với y văn quốc tế và Trần Thảo Vi và cộng sự (2024) '
        'tại Huế.',
    ],
    21: [  # Bảng 3.22 YTBV
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.22 — chữ xanh]',
        'Bảng 3.22 cho thấy thứ tự cường độ yếu tố BẢO VỆ: HỖ TRỢ TỪ CHA MẸ '
        '(ĐTB = 57,65) > HỖ TRỢ TỪ BẠN BÈ (54,27) > GẮN BÓ TRƯỜNG HỌC (52,60). '
        'Cha mẹ là nguồn hỗ trợ CAO NHẤT — phù hợp đặc thù văn hóa Á tập trung '
        'vào gia đình.',
        'Đáng chú ý: trong nhóm hỗ trợ cha mẹ, mục cao nhất MSPSS.3 "gia đình '
        'thực sự cố gắng giúp đỡ" (65,48) phản ánh hỗ trợ THỰC TẾ, trong khi '
        'mục thấp nhất MSPSS.8 "có thể nói chuyện về vấn đề" (47,54) gợi ý '
        'cha mẹ Việt Nam giỏi HỖ TRỢ THỰC TẾ nhưng yếu hơn về GIAO TIẾP CẢM '
        'XÚC. Phù hợp với phát hiện Phạm và cộng sự (2024 — VN003): chăm sóc '
        'cảm xúc β = −0,40 mạnh hơn chăm sóc thể chất (không có ý nghĩa thống '
        'kê).',
        'Trong gắn bó trường học, mục THẤP NHẤT là PSSM.5 "phần lớn giáo viên '
        'quan tâm đến tôi" (47,36) — gợi ý điểm yếu về QUAN HỆ GV-HS, cần '
        'can thiệp đào tạo giáo viên về kỹ năng quan tâm và lắng nghe.',
    ],
    23: [  # Bảng 3.24 ALHT path
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.24 — chữ xanh]',
        'Bảng 3.24 cho thấy áp lực học tập có cường độ tác động MẠNH NHẤT lên '
        'ba dạng RLLA. β chuẩn hóa: ALHT → RLLATQ = 0,510 (large theo Cohen '
        '1988); ALHT → RLLAXH = 0,490 (large); ALHT → RLLACL = 0,253 (small-'
        'to-medium); ALHT → RLLA tổng (3 factors) = 0,533, R² = 0,284. Tất cả '
        'p < 0,001.',
        'Pattern: ALHT tác động MẠNH NHẤT lên LO ÂU LAN TỎA và LO ÂU XÃ HỘI '
        '(hai dạng liên quan đánh giá), YẾU NHẤT lên LO ÂU CHIA LY (rối loạn '
        'khởi phát sớm ít phụ thuộc bối cảnh học đường). Phù hợp với khung 6 '
        'trục của Pascoe, Hetrick và Parker (2020): áp lực thi cử tác động '
        'qua trục HPA, thay thế giấc ngủ, và mất kết nối trường học.',
        'Phát hiện này phù hợp với β = 0,508 (Walder và cộng sự, 2025 meta '
        'DMHI 21 RCT) và bằng chứng cấp dân số của Brunborg và cộng sự (2025) '
        'trên 979.043 thanh thiếu niên Na Uy: bất mãn trường học giải thích '
        'TOÀN BỘ xu hướng tăng căng thẳng tâm thần.',
    ],
    26: [  # Bảng 3.26 NĐT path
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.26 — chữ xanh]',
        'Bảng 3.26 cho thấy nghiện điện thoại tác động MẠNH NHẤT lên LO ÂU XÃ '
        'HỘI (β = 0,383, p < 0,001) — cao hơn cả lan tỏa (β = 0,336) và chia '
        'ly (β = 0,265). β tổng (3 factors) = 0,400; R² = 0,160 (medium).',
        'Phát hiện β NĐT → RLLAXH cao nhất phù hợp với cơ chế SO SÁNH XÃ HỘI '
        'và NỖI SỢ BỊ ĐÁNH GIÁ trên mạng xã hội. Phù hợp với Fassi và cộng '
        'sự (2025) trong Nature Human Behaviour trên 3.340 thanh thiếu niên '
        'UK với chẩn đoán lâm sàng: thanh thiếu niên có rối loạn nội hóa '
        'dùng mạng xã hội nhiều hơn và kém hài lòng hơn về quan hệ online.',
        'Lưu ý: Walder và cộng sự (2025) phân tích tổng hợp 21 RCT DMHI cho '
        'lo âu xã hội xác lập can thiệp số ĐẶC THÙ cho SAD đạt Hedges g = '
        '0,878 — gợi ý hướng can thiệp tiềm năng qua iCBT app cho học sinh '
        'có nguy cơ cao.',
    ],
    29: [  # Bảng 3.28 BNHĐ path
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.28 — chữ xanh]',
        'PHÁT HIỆN ĐẶC BIỆT: Bảng 3.28 cho thấy bắt nạt học đường có pattern '
        'KHÁC BIỆT với áp lực học tập và tự trọng — β BNHĐ → RLLACL = 0,376 '
        'là MẠNH NHẤT trong ba dạng RLLA, cao hơn cả β → RLLATQ (0,215) và '
        'β → RLLAXH (0,253). Tất cả p < 0,001.',
        'Ba cơ chế giải thích: (1) Cơ chế gắn bó-an toàn: học sinh bị bắt '
        'nạt mất cảm giác AN TOÀN ở trường, không muốn tách khỏi cha mẹ — '
        'biểu hiện chính của RLLACL; (2) School refusal: bắt nạt → từ chối '
        'đến trường → biểu hiện như lo âu chia ly; (3) Đặc thù lứa tuổi '
        'THCS chuyển tiếp — bắt nạt có thể KÍCH HOẠT lại pattern lo âu '
        'chia ly chưa hoàn toàn biến mất.',
        'Đây là phát hiện ĐẶC THÙ Việt Nam — chưa được ghi nhận rõ rệt '
        'trong y văn quốc tế. Gợi ý: theo dõi học sinh có biểu hiện "school '
        'refusal" và kiểm tra liệu nguyên nhân là bắt nạt — không chỉ là '
        '"lo âu chia ly đơn thuần". Lưu ý: mô hình BNHĐ–RLLA (3 factors) '
        'có RMSEA = 0,129 (KÉM); nên ưu tiên báo cáo mô hình 2 factors '
        '(RMSEA = 0,030, CFI = 0,999 — XUẤT SẮC).',
    ],
    32: [  # Bảng 3.30 GBTH
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.30 — chữ xanh]',
        'Bảng 3.30 cho thấy gắn bó trường học có tác động bảo vệ KHIÊM TỐN: '
        'β GBTH → RLLATQ = −0,108 (p = 0,002, small); β → RLLAXH = −0,187 '
        '(p < 0,001, mạnh nhất); β → RLLACL = 0,014 (p = 0,696, KHÔNG ý '
        'nghĩa). R² tổng = 0,024 — chỉ 2,4% phương sai.',
        'Pattern: GBTH bảo vệ MẠNH NHẤT cho LO ÂU XÃ HỘI (logic — qua bạn '
        'bè và giáo viên) nhưng KHÔNG có tác động cho LO ÂU CHIA LY (logic '
        '— chia ly liên quan cha mẹ chứ không phải trường học). Phù hợp '
        'với khung Goodenow (1993) về Psychological Sense of School '
        'Membership.',
    ],
    35: [  # Bảng 3.32 TTr
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.32 — chữ xanh]',
        'Bảng 3.32 cho thấy lòng tự trọng là yếu tố BẢO VỆ MẠNH NHẤT: β TTr '
        '→ RLLATQ = −0,455 (p < 0,001); β → RLLAXH = −0,415 (p < 0,001); '
        'β → RLLACL = −0,087 (p = 0,020, yếu); β → RLLA tổng = −0,457, '
        'R² = 0,209.',
        'Đáng chú ý: |β TTr| / |β ALHT| = 0,455/0,510 ≈ 0,89 cho RLLATQ và '
        '0,415/0,490 ≈ 0,85 cho RLLAXH — tự trọng có cường độ tác động '
        'NGANG BẰNG ~85–89% áp lực học tập. Đây là phát hiện QUAN TRỌNG '
        'phù hợp giả thuyết "yếu tố bảo vệ KHÔNG yếu hơn yếu tố nguy cơ '
        'như giả định thông thường".',
        'Phù hợp với Sowislo và Orth (2013) phân tích tổng hợp 95 nghiên '
        'cứu thuần tập trong Psychological Bulletin: lòng tự trọng và lo '
        'âu có quan hệ HAI CHIỀU với β = −0,10. Cường độ ở mẫu chương 3 '
        '(−0,455) lớn hơn meta-analysis (−0,10) do mô hình SEM tích hợp '
        'biến tiềm ẩn — kết quả đáng tin cậy.',
    ],
    38: [  # Bảng 3.34 HTCM
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.34 — chữ xanh]',
        'PHÁT HIỆN ĐẶC BIỆT: Bảng 3.34 cho thấy hỗ trợ từ cha mẹ có pattern '
        'KHÁC THƯỜNG — β HTCM → RLLACL = 0,000 (p = 0,997, TUYỆT ĐỐI KHÔNG '
        'có tác động!). Trong khi đó, β HTCM → RLLAXH = −0,273 là MẠNH '
        'NHẤT trong ba dạng; β → RLLATQ = −0,172 (p < 0,001).',
        'TRÁI với kỳ vọng thông thường: thường giả định "hỗ trợ cha mẹ tốt '
        '→ giảm lo âu chia ly". Nhưng dữ liệu cho thấy hỗ trợ cha mẹ '
        'KHÔNG có tác động cho lo âu chia ly. Khả năng giải thích: trong '
        'văn hóa Á (đặc biệt Việt Nam), hỗ trợ cha mẹ mạnh có thể CỦNG CỐ '
        'gắn bó với cha mẹ → không tách khỏi cha mẹ → có thể duy trì '
        'pattern lo âu chia ly thay vì giảm. Cần kiểm chứng thêm trong '
        'nghiên cứu tiếp theo.',
    ],
    41: [  # Bảng 3.36 HTBB
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.36 — chữ xanh]',
        'PHÁT HIỆN CRITICAL: Bảng 3.36 cho thấy hỗ trợ từ bạn bè GẦN NHƯ '
        'KHÔNG có tác động bảo vệ trước RLLA. β HTBB → RLLATQ = −0,015 '
        '(p = 0,646, KHÔNG ý nghĩa); β → RLLACL = −0,019 (p = 0,577, KHÔNG '
        'ý nghĩa); β → RLLAXH = −0,079 (p = 0,020, có ý nghĩa nhưng cường '
        'độ NHỎ |β| < 0,10). R² tổng = 0,002 (tối thiểu).',
        'TRÁI với y văn quốc tế (Murphy 2024 peer support có tiềm năng cải '
        'thiện outcome). Bốn khả năng giải thích: (1) Đặc thù lứa tuổi '
        'THCS: học sinh lớp 6–9 chưa phát triển đầy đủ kỹ năng dựa vào '
        'bạn bè — quan trọng hơn ở THPT; (2) Văn hóa Á tập trung gia đình: '
        'cha mẹ là nguồn hỗ trợ chính, bạn bè ở vị trí thứ yếu; (3) Chất '
        'lượng vs số lượng: học sinh có nhiều bạn nhưng quan hệ NÔNG; '
        '(4) Mô hình SEM 2 factor được tác giả luận án đánh dấu "mô hình '
        'lỗi" — có thể vấn đề specification.',
        'Hàm ý quan trọng: KHÔNG dựa vào "tăng hỗ trợ bạn bè" làm can thiệp '
        'chính cho RLLA học sinh THCS Việt Nam. Cần đo CHẤT LƯỢNG quan hệ '
        '+ lặp lại ở THPT.',
    ],
    43: [  # Bảng 3.38 YTNC + YTBV tích hợp
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.38 — chữ xanh]',
        'Bảng 3.38 cho thấy mô hình TÍCH HỢP yếu tố nguy cơ + bảo vệ: β '
        'YTNC → RLLA = 0,669 và β YTBV → RLLA = −0,220 (cả hai p < 0,001). '
        'R² CHUNG = 0,598 — vượt ngưỡng "effect size LỚN" R² > 0,26 theo '
        'Cohen (1988) hơn gấp đôi.',
        'So sánh cường độ: |β YTNC| / |β YTBV| = 0,669 / 0,220 ≈ 3,04 — '
        'yếu tố NGUY CƠ áp đảo yếu tố BẢO VỆ với tỷ số 3 lần. Hàm ý can '
        'thiệp: ưu tiên GIẢM yếu tố nguy cơ (áp lực học tập, nghiện điện '
        'thoại, bắt nạt) — vì cường độ tác động lớn hơn — đồng thời tăng '
        'yếu tố bảo vệ để có hiệu ứng cộng hưởng.',
    ],
    47: [  # Bảng 3.42 YTBV riêng
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.42 — chữ xanh]',
        'Bảng 3.42 cho thấy YTBV TỔNG riêng (không tích hợp với YTNC): β '
        '= −0,352, R² = 0,124. So với YTNC riêng (β = 0,747, R² = 0,558 '
        'từ Bảng 3.40), tỷ số R² = 0,124 / 0,558 ≈ 0,222 — yếu tố bảo vệ '
        'tổng chỉ giải thích ~22% so với yếu tố nguy cơ tổng.',
        'Khi tích hợp (Bảng 3.38), |β YTBV| giảm từ 0,352 xuống 0,220 — '
        'GIẢM 37% — gợi ý YTBV không hoàn toàn ĐỘC LẬP với YTNC. Có thể '
        'có tương quan: ví dụ học sinh có ALHT cao có tự trọng thấp; '
        'học sinh bị bắt nạt có gắn bó trường học thấp. Khuyến nghị '
        'nghiên cứu tiếp theo phân tích MEDIATION cụ thể giữa YTNC và '
        'YTBV.',
    ],
    48: [  # Bảng 3.43 BPĐP biểu hiện
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.43 — chữ xanh]',
        'Bảng 3.43 cho thấy ba nhóm biện pháp đối phó với thứ tự ĐTB: TÌM '
        'KIẾM SỰ HỖ TRỢ (55,00) > TẬP TRUNG GIẢI QUYẾT VẤN ĐỀ (53,18) > '
        'TRÁNH NÉ VẤN ĐỀ (41,34). Học sinh THCS Việt Nam có xu hướng dùng '
        'NHIỀU NHẤT là tìm kiếm hỗ trợ — phù hợp văn hóa Á tập trung gia '
        'đình.',
        'Tuy nhiên, ĐTB tránh né vẫn ở mức 41,34 (gần giữa thang) — gợi '
        'ý một bộ phận học sinh dùng chiến lược ÍT THÍCH ỨNG. Phù hợp '
        'khung Compas và cộng sự (2017) phân tích tổng hợp 212 nghiên '
        'cứu trên 80.850 trẻ em + vị thành niên: ứng phó NGẮT KẾT NỐI '
        '(né tránh, phủ nhận, ức chế cảm xúc) liên quan với MỨC CAO HƠN '
        'của triệu chứng nội hóa.',
    ],
    50: [  # Bảng 3.45 BPĐP path
        '[BÌNH LUẬN BỔ SUNG sau Bảng 3.45 — chữ xanh]',
        'PHÁT HIỆN CRITICAL: Bảng 3.45 cho thấy biện pháp đối phó có cường '
        'độ tác động CỰC LỚN nhưng dấu DƯƠNG: β BPĐP → RLLATQ = 0,749 '
        '(R² = 0,561 — lớn nhất chương 3); β → RLLAXH = 0,670 (R² = '
        '0,449); β → RLLACL = 0,132 (yếu, p = 0,004); β tổng = 0,735, '
        'R² = 0,540.',
        'NGHỊCH LÝ — dấu DƯƠNG trái với giả định trực giác "đối phó nhiều '
        'thì lo âu giảm". Phù hợp hiện tượng MALADAPTIVE COPING ESCALATION '
        '(Compas và cộng sự, 2017): học sinh càng lo âu càng dùng nhiều '
        'biện pháp đối phó, nhưng nếu không hiệu quả thì lo âu vẫn duy '
        'trì hoặc tăng. Tương đương với giải thích Bảng 3.43: khoảng 41% '
        'biểu hiện tránh né — chiến lược ÍT THÍCH ỨNG.',
        'CẢNH BÁO: mô hình BPĐP có FIT KÉM trên TẤT CẢ tổ hợp — RMSEA '
        '0,080–0,204 (vượt ngưỡng 0,06; thậm chí 0,204 = RẤT KÉM); CFI '
        '0,865–0,911 (dưới 0,95); χ²/df 9,6–57,3. KHÔNG nên báo cáo '
        'β = 0,749 làm phát hiện chính. Đề xuất: phân tách Brief-COPE 14 '
        'nhân tố theo Carver (1997) trong nghiên cứu tiếp theo.',
    ],
}


def main():
    # Copy file gốc
    shutil.copy(SRC, OUT)
    print(f'  Copied: {SRC.name} → {OUT.name}')

    # Open copy and insert comments
    d = Document(OUT)
    tables = d.tables
    print(f'  Total tables: {len(tables)}')

    # Insert in REVERSE order (so later inserts don't shift earlier indices)
    inserted_count = 0
    for tidx in sorted(COMMENTS.keys(), reverse=True):
        if tidx >= len(tables):
            print(f'  ⚠ Table {tidx} out of range')
            continue
        comments = COMMENTS[tidx]
        table = tables[tidx]
        # Insert paragraphs in REVERSE so each goes RIGHT AFTER table (last inserted = first appearing)
        for para_text in reversed(comments):
            new_p = make_blue_paragraph(para_text)
            table._element.addnext(new_p)
        inserted_count += len(comments)

    d.save(OUT)
    print(f'  ✓ Inserted {inserted_count} blue paragraphs into {len(COMMENTS)} tables')
    print(f'  ✓ {OUT.name} ({OUT.stat().st_size//1024} KB)')


if __name__ == '__main__':
    main()
