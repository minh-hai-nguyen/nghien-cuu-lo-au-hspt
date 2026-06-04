"""Insert blue paragraphs vào Tiểu kết chương 3 + Kết luận
trong file 00_Thuc trang va khung tap huan-v2.docx
"""
import sys, io
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

FILE = Path('c:/Users/HLC/OneDrive/read_books/Lo-au/01_Bao-cao/bang-so-lieu-binh-luan/00_Thực trạng và khung tập huấn-v2.docx')


def make_blue_paragraph(text):
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


# Map paragraph_substr → list of blue paragraphs to insert AFTER
INSERTIONS = [
    # 1. Tiểu kết đoạn 428 — bổ sung sau đoạn về thực trạng RLLA
    ('Về thực trạng, rối loạn lo âu của học sinh chủ yếu ở mức trung bình', [
        '[Bổ sung CTH — chữ xanh] Cụ thể, trên thang RCADS quy đổi 0–100, ĐTB '
        'lo âu lan tỏa = 55,82 (cao nhất); lo âu xã hội = 48,41; lo âu chia ly '
        '= 25,06 (thấp nhất). Áp dụng ngưỡng 65/100 (T-score borderline) với '
        'giả định phân phối normal, ước lượng tỷ lệ học sinh có triệu chứng '
        'vượt ngưỡng: RLLATQ ≈ 33,85% (nam 26,88%, nữ 40,11%); RLLAXH ≈ 25,98% '
        '(nam 19,25%, nữ 32,06%); RLLAC ≈ 5,01% — phù hợp khung phát triển '
        'DSM-5 (lo âu chia ly là rối loạn khởi phát sớm). Tỷ lệ ước lượng '
        'gần với V-NAMHS 2022 chẩn đoán DSM-5 (18,45%) và Hoàng Trung Học 2025 '
        'sàng lọc DASS-21 sau COVID (25,4%) — gợi ý quy mô vấn đề ở mức KHẨN '
        'CẤP cho lo âu lan tỏa và xã hội.',

        '[Bổ sung CTH — chữ xanh] Về khác biệt giới tính theo loại RLLA, '
        'F-test trên Bảng 3.20 cho thấy mô hình BA TẦNG: F (giới × RLLATQ) = '
        '44,484 (p < 0,001) và F (giới × RLLAXH) = 45,984 (p < 0,001) — đều '
        'rõ rệt với Cohen d = 0,365 và 0,370 (gần nhau, small-to-medium). '
        'Trong khi đó, F (giới × RLLAC) = 0,246 (p = 0,620) — KHÔNG có chênh '
        'lệch giới ở lo âu chia ly, phù hợp khung phát triển. Về khối lớp, '
        'RLLAC giảm đơn điệu từ lớp 6 (32,13) đến lớp 9 (19,46) — tổng giảm '
        '12,67 điểm trong 3 năm.',
    ]),

    # 2. Tiểu kết đoạn 429 — bổ sung sau đoạn về SEM yếu tố
    ('Đối với các yếu tố ảnh hưởng, mô hình SEM chỉ ra rằng các yếu tố nguy cơ', [
        '[Bổ sung CTH — chữ xanh] Cụ thể, hệ số chuẩn hóa β trong mô hình SEM '
        'cho từng yếu tố: ÁP LỰC HỌC TẬP β = 0,510 (RLLATQ) và 0,490 (RLLAXH) '
        '— cường độ lớn nhất; NGHIỆN ĐIỆN THOẠI β = 0,336 và 0,383 (mạnh nhất '
        'cho RLLAXH); BẮT NẠT HỌC ĐƯỜNG có pattern ĐẶC BIỆT — β cho RLLACL '
        '(0,376) MẠNH NHẤT trong ba dạng, vượt cả β cho RLLATQ (0,215) và '
        'RLLAXH (0,253) — phát hiện đặc thù Việt Nam (cơ chế gắn bó-an toàn '
        'tại trường + school refusal). Yếu tố nguy cơ tổng có β = 0,747 và '
        'giải thích 55,8% phương sai (R² = 0,558).',

        '[Bổ sung CTH — chữ xanh] Đối với yếu tố bảo vệ, TỰ TRỌNG có cường '
        'độ MẠNH NHẤT (β = −0,455 cho RLLATQ và −0,415 cho RLLAXH) — tỷ số '
        '|β TTr| / |β ALHT| = 0,85–0,89, gần ngang với áp lực học tập. HỖ '
        'TRỢ TỪ CHA MẸ có pattern đặc biệt — β = −0,273 cho RLLAXH (mạnh '
        'nhất) nhưng β = 0,000 cho RLLACL (KHÔNG có tác động — phù hợp đặc '
        'thù văn hóa Á gắn bó gia đình). Đáng chú ý nhất, HỖ TRỢ TỪ BẠN BÈ '
        'GẦN NHƯ KHÔNG có tác động bảo vệ (β = −0,015 đến −0,079) — TRÁI '
        'với y văn quốc tế, gợi ý đặc thù lứa tuổi THCS Việt Nam. Yếu tố '
        'bảo vệ tổng có β = −0,352 và R² = 0,124. Mô hình tích hợp YTNC + '
        'YTBV đạt R² = 0,598 (very large theo Cohen 1988).',

        '[Bổ sung CTH — chữ xanh] Cảnh báo phương pháp luận: mô hình SEM '
        'cho biện pháp đối phó có hệ số β CỰC LỚN (β = 0,749 cho RLLATQ; '
        'R² = 0,561), NHƯNG fit indices KÉM (RMSEA 0,080–0,204; CFI 0,865–'
        '0,911 — dưới ngưỡng Hu & Bentler 1999). Phù hợp hiện tượng '
        'maladaptive coping escalation (Compas và cộng sự, 2017) nhưng cần '
        'phân tách Brief-COPE 14 nhân tố theo Carver (1997) trong nghiên '
        'cứu tiếp theo — KHÔNG nên báo cáo β = 0,749 làm phát hiện chính.',
    ]),

    # 3. Kết luận đoạn 451 — bổ sung sau đoạn về khác biệt giới + khối
    ('Kết quả nghiên cứu chỉ ra sự khác biệt về mức độ rối loạn lo âu giữa các nhóm', [
        '[Bổ sung CTH — chữ xanh] Định lượng cụ thể: tỷ số Cohen d (chuẩn '
        'hóa effect size) cho thấy chênh lệch giới ở lo âu lan tỏa (d = '
        '0,365) và lo âu xã hội (d = 0,370) GẦN NHƯ BẰNG NHAU — đều thuộc '
        'nhóm small-to-medium. Trong khi F-test cho thấy RLLAXH (F = 45,984) '
        'có vẻ vượt RLLATQ (F = 44,484), khác biệt thực tế về effect size '
        'chỉ chênh 1,52% — không đáng kể. Đối với lo âu chia ly, F = 0,246 '
        '(p = 0,620) và Cohen d ≈ 0,03 — KHÔNG có chênh lệch giới có ý '
        'nghĩa thống kê. Pattern ba tầng này phù hợp khung Salk, Hyde và '
        'Abramson (2017): chênh lệch giới đỉnh ở tuổi 13–15 (OR = 3,02 '
        'cho rối loạn nội hóa).',
    ]),

    # 4. Kết luận đoạn 452 — bổ sung sau đoạn về yếu tố nguy cơ
    ('Đối với các yếu tố nguy cơ, kết quả nghiên cứu khẳng định rằng áp lực học tập', [
        '[Bổ sung CTH — chữ xanh] Đặc biệt, kết quả SEM cho thấy bắt nạt '
        'học đường có pattern KHÁC BIỆT với áp lực học tập và tự trọng — '
        'tác động MẠNH NHẤT lên lo âu CHIA LY (β = 0,376), vượt cả tác '
        'động lên lan tỏa (β = 0,215) và xã hội (β = 0,253). Đây là phát '
        'hiện ĐẶC THÙ chưa được ghi nhận rõ rệt trong y văn quốc tế, gợi '
        'ý cơ chế đặc biệt: bắt nạt → mất an toàn ở trường → tránh né '
        'trường → biểu hiện như lo âu chia ly (school refusal) ở lứa '
        'tuổi THCS chuyển tiếp.',
    ]),

    # 5. Kết luận đoạn 453 — bổ sung sau đoạn về yếu tố bảo vệ
    ('Các yếu tố bảo vệ như gắn bó với trường học, hỗ trợ xã hội và tự trọng', [
        '[Bổ sung CTH — chữ xanh] Trong các yếu tố bảo vệ, tự trọng có '
        'cường độ MẠNH NHẤT — tỷ số |β tự trọng| / |β áp lực học tập| đạt '
        '0,85–0,89, tức ngang bằng 85–89% áp lực học tập. Đáng chú ý, hai '
        'phát hiện đặc biệt cần lưu ý: (1) Hỗ trợ từ cha mẹ có β = 0,000 '
        'cho lo âu chia ly (KHÔNG có tác động) — có thể do văn hóa Á '
        'củng cố gắn bó gia đình mạnh; (2) Hỗ trợ từ bạn bè GẦN NHƯ '
        'KHÔNG có tác động bảo vệ trên cả ba dạng RLLA — TRÁI với y văn '
        'quốc tế, gợi ý đặc thù lứa tuổi THCS Việt Nam. Hai phát hiện '
        'này đặt ra câu hỏi nghiên cứu cho các công trình tiếp theo về '
        'cơ chế hỗ trợ xã hội ở thanh thiếu niên Việt Nam.',
    ]),

    # 6. Kết luận đoạn 454 — bổ sung sau đoạn về biện pháp đối phó
    ('Về các biện pháp đối phó, kết quả nghiên cứu cho thấy các chiến lược', [
        '[Bổ sung CTH — chữ xanh] Định lượng cụ thể: β BPĐP → RLLATQ = '
        '0,749 với R² = 0,561 — cường độ LỚN NHẤT trong toàn chương 3, '
        'vượt cả β áp lực học tập (0,510) và β tự trọng (−0,455). Tuy '
        'nhiên, fit indices của mô hình BPĐP KÉM trên tất cả tổ hợp '
        '(RMSEA 0,080–0,204; CFI 0,865–0,911) — vi phạm ngưỡng Hu & '
        'Bentler (1999). Phù hợp hiện tượng maladaptive coping '
        'escalation (Compas và cộng sự, 2017): học sinh càng lo âu càng '
        'dùng nhiều biện pháp đối phó, nhưng nếu không hiệu quả thì lo '
        'âu vẫn duy trì hoặc tăng. Khuyến nghị nghiên cứu tiếp theo '
        'phân tách Brief-COPE 14 nhân tố theo Carver (1997) để xác '
        'định chiến lược TÍCH CỰC vs TIÊU CỰC riêng biệt.',
    ]),
]


def main():
    d = Document(FILE)
    paragraphs = list(d.paragraphs)
    print(f'Loaded {len(paragraphs)} paragraphs')

    inserted = 0
    for substr, blue_texts in INSERTIONS:
        # Find paragraph containing substr
        target = None
        for p in paragraphs:
            if substr in p.text:
                target = p
                break
        if target is None:
            print(f'  ⚠ NOT FOUND: "{substr[:50]}..."')
            continue
        # Insert blue paragraphs in REVERSE so they appear in order after target
        for blue_text in reversed(blue_texts):
            new_p = make_blue_paragraph(blue_text)
            target._element.addnext(new_p)
        inserted += len(blue_texts)
        print(f'  ✓ Inserted {len(blue_texts)} after: "{substr[:50]}..."')

    d.save(FILE)
    print(f'\n✓ Total inserted: {inserted} blue paragraphs')
    print(f'✓ Saved: {FILE.name} ({FILE.stat().st_size//1024} KB)')


if __name__ == '__main__':
    main()
