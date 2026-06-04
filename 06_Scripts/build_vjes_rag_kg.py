"""
Build RAG + KG cho hướng dẫn VJES (Vietnam Journal of Educational Sciences).
Output:
  - bai-bao-khgdvn/vjes_rag.json   (chunks + metadata để tra cứu)
  - bai-bao-khgdvn/vjes_kg.json    (nodes + edges concepts)

Sử dụng để đối chiếu khi viết bài đăng VJES — tránh sai format.
"""
import json
from pathlib import Path
from datetime import datetime

ROOT = Path(r"c:/Users/HLC/OneDrive/read_books/Lo-au")
OUT_DIR = ROOT / "bai-bao-khgdvn"

# ============================================================
# RAG — 30 chunk có cấu trúc, mỗi chunk = 1 quy định độc lập
# ============================================================
chunks = [
    # --- THÔNG TIN CHUNG ---
    {"id": "VJES_01", "topic": "thong_tin_chung", "title": "Cơ quan chủ quản + tên tạp chí",
     "content": "Tạp chí Khoa học Giáo dục Việt Nam (Vietnam Journal of Educational Sciences - VJES) là cơ quan ngôn luận khoa học của Viện Khoa học Giáo dục Việt Nam (VNIES). ISSN 2615-8957; e-ISSN 3030-4490. Tổng biên tập gắn với viện trưởng VNIES (thầy Lê Anh Vinh). Bài đăng VJES được Hội đồng Giáo sư Nhà nước tính đến 1 điểm.",
     "source": "Hướng dẫn 2025 — Mục I.1"},

    {"id": "VJES_02", "topic": "thong_tin_chung", "title": "Mục đích, phạm vi xuất bản",
     "content": "VJES công bố các kết quả nghiên cứu lý luận + thực tiễn chất lượng cao trong giáo dục - đào tạo: lý thuyết, ứng dụng, đánh giá, triển khai, thực nghiệm, định hướng. Chủ đề: dạy-học; chiến lược, chính sách, quản lý giáo dục; điều kiện đảm bảo chất lượng các bậc học.",
     "source": "Hướng dẫn 2025 — Mục I.1.1"},

    {"id": "VJES_03", "topic": "the_loai_bai", "title": "6 thể loại bài chấp nhận",
     "content": "VJES chấp nhận 6 thể loại: (1) Nghiên cứu gốc (Original research) — phổ biến nhất, 4 mục Giới thiệu / PP NC / Kết quả / Kết luận; (2) Tổng quan tài liệu (Review paper) — phân tích toàn diện tài liệu đã có, 4 mục như Original research; (3) Bài báo thực nghiệm (Experimental article) — lý thuyết + thực hành; (4) Bài báo phương pháp luận (Methodology paper); (5) Bài đánh giá (Review article); (6) Nghiên cứu trường hợp (Case studies) — 6 mục có Mô tả trường hợp + Thảo luận.",
     "source": "Hướng dẫn 2025 — Mục 2.3.2"},

    {"id": "VJES_04", "topic": "cau_truc_review_paper", "title": "Cấu trúc Review paper (Tổng quan tài liệu)",
     "content": "Cấu trúc 4 mục BẮT BUỘC: (1) Giới thiệu / Đặt vấn đề; (2) Phương pháp nghiên cứu; (3) Kết quả nghiên cứu; (4) Kết luận. Có thể thêm mục Thảo luận (không bắt buộc). Review paper có 3 dạng: phân tích tổng hợp; đánh giá có hệ thống; đánh giá tài liệu.",
     "source": "Hướng dẫn 2025 — Bảng thể loại + Template"},

    {"id": "VJES_05", "topic": "dinh_ky", "title": "Định kỳ xuất bản",
     "content": "VJES xuất bản 2 phiên bản: số tiếng Việt (12 số/năm, mỗi cuối tháng) và số tiếng Anh (4 số/năm vào cuối tháng 3, 6, 9, 12). Mỗi số có bản in + bản trực tuyến. Có thể có số đặc biệt (số chuyên đề).",
     "source": "Hướng dẫn 2025 — Mục I.1.2"},

    {"id": "VJES_06", "topic": "phan_bien", "title": "Quy trình phản biện",
     "content": "Phản biện kín 2 chiều (double-blind peer review) bởi các nhà khoa học uy tín đúng lĩnh vực. Quy trình 6 bước: Tác giả gửi bài → Sơ loại → Phản biện kín 2 chiều → Sửa chữa → Quyết định của Tổng biên tập (Chấp nhận / Từ chối) → Chế bản và biên tập → Trực tuyến → Xuất bản.",
     "source": "Hướng dẫn 2025 — Mục I.1.3"},

    {"id": "VJES_07", "topic": "kinh_phi", "title": "Kinh phí xuất bản",
     "content": "Bài số tiếng Việt: 2.500.000 VNĐ/bài (đóng SAU KHI bài được chấp nhận đăng). Bài số tiếng Anh: KHÔNG phải hỗ trợ kinh phí (theo quy định hiện hành). Tạp chí truy cập mở (Creative Commons Attribution-Non-Commercial-NoDerivatives 4.0).",
     "source": "Hướng dẫn 2025 — Mục I.1.4"},

    # --- TURNITIN ---
    {"id": "VJES_08", "topic": "turnitin", "title": "Tỷ lệ trùng lặp Turnitin",
     "content": "VJES kiểm tra trùng lặp qua Turnitin. KHUYẾN NGHỊ: tỷ lệ trùng lặp ≤ 25% (KHÔNG tính TLTK). Đây là TIÊU CHÍ SƠ LOẠI — Ban Biên tập căn cứ tỷ lệ + nội dung + mức độ trùng lặp để quyết định cuối cùng. → CHIẾN LƯỢC: paraphrase 100%; không copy từ tác phẩm đã upload Turnitin (luận án CTH, báo cáo Can thiệp v5, CLEAN_v3_CTH_v6).",
     "source": "Hướng dẫn 2025 — Mục 2.1 điểm 9"},

    # --- ĐỘ DÀI + FORMAT ---
    {"id": "VJES_09", "topic": "do_dai", "title": "Độ dài bài + số trang",
     "content": "Độ dài: TỐI THIỂU 5.000 TỪ, KHÔNG QUÁ 7.000 TỪ (tương đương 8-12 trang A4). Áp dụng cho TẤT CẢ thể loại bài.",
     "source": "Hướng dẫn 2025 — Mục Định dạng văn bản"},

    {"id": "VJES_10", "topic": "format", "title": "Định dạng văn bản",
     "content": "Định dạng: Microsoft Word hoặc RTF (KHÔNG nhận PDF). Spacing: Before 3pt, After 3pt. Line spacing: Multiple 1.1. Font: Times New Roman. Cỡ chữ: Tên bài 12pt in đậm; Tóm tắt + Từ khóa 11pt thường (nhưng chữ 'Tóm tắt'/'Từ khóa' in đậm); Body 12pt. Đánh số trang tự động. Indent bằng tab, KHÔNG dùng phím cách. Bảng bằng chức năng table, không phải spreadsheet.",
     "source": "Hướng dẫn 2025 — Mục Định dạng văn bản"},

    {"id": "VJES_11", "topic": "heading", "title": "Mục và tiểu mục",
     "content": "Tối đa 3 cấp độ heading. Cấp 1 (mục chính): chữ thường, in đậm. VD: '1. Đặt vấn đề'. Cấp 2 (tiểu mục cấp 1): chữ nghiêng, in đậm. VD: '2.3. Một số kết quả'. Cấp 3 (tiểu mục cấp 2): chữ nghiêng, in thường. VD: '2.3.1. Kết quả khảo sát'. KHÔNG dùng chú thích cuối trang (footnote). KHÔNG viết tắt (trừ thông lệ quốc tế).",
     "source": "Hướng dẫn 2025 — Mục Mục và tiểu mục + Viết tắt"},

    # --- PHẦN TIÊU ĐỀ ---
    {"id": "VJES_12", "topic": "ten_bai", "title": "Tên bài báo",
     "content": "Tên bài phải NGẮN GỌN: DƯỚI 40 KÝ TỰ, không quá 12 từ. Chứa 1-2 từ khóa chính. KHÔNG viết tắt trong tên bài. Bằng cả tiếng Việt + tiếng Anh. Thân thiện với công cụ tìm kiếm (SEO).",
     "source": "Hướng dẫn 2025 — Mục 2.3.1 Tên bài báo"},

    {"id": "VJES_13", "topic": "tom_tat", "title": "Tóm tắt (Abstract)",
     "content": "Tóm tắt 150-250 TỪ, bằng cả VN + EN. KHÔNG chứa viết tắt không xác định hoặc TLTK. Phải trả lời 3 câu hỏi: (1) Mục đích NC + khoảng trống đang giải quyết; (2) Phương pháp đạt mục tiêu; (3) Kết quả quan trọng nhất. NÊN: dùng cụm từ trong tên bài ở 1-2 câu đầu (chỉ 2 câu đầu hiển thị trong kết quả công cụ tìm kiếm); lặp từ khóa 3-6 lần.",
     "source": "Hướng dẫn 2025 — Mục 2.3.1 Tóm tắt"},

    {"id": "VJES_14", "topic": "tu_khoa", "title": "Từ khóa",
     "content": "TỪ KHÓA: 4-6 từ khóa (hoặc cụm từ khóa ngắn), bằng VN + EN. Phản ánh sát nhất nội dung bản thảo. NÊN tra Google Trends/AdWords trước khi chọn — từ khóa quan trọng với SEO + AI.",
     "source": "Hướng dẫn 2025 — Mục 2.3.1 Từ khóa"},

    {"id": "VJES_15", "topic": "tac_gia", "title": "Thông tin tác giả",
     "content": "BẮT BUỘC ghi đầy đủ: (1) Họ + tên đầy đủ của các tác giả; (2) Đơn vị công tác, điện thoại, email (tương ứng từng tác giả); (3) Ghi rõ TÁC GIẢ LIÊN HỆ + email đang hoạt động; (4) ORCID (KHUYẾN NGHỊ — sẽ tự đính vào hồ sơ ORCID khi bài đăng). Đăng ký ORCID miễn phí tại https://orcid.org.",
     "source": "Hướng dẫn 2025 — Mục 2.3.1 Thông tin tác giả"},

    # --- LLM ---
    {"id": "VJES_16", "topic": "llm", "title": "Khai báo sử dụng LLM/AI",
     "content": "Nếu tác giả đã sử dụng LLM (ChatGPT, Claude, ...) trong quá trình phát triển bài báo, PHẢI KHAI BÁO trong mục Phương pháp nghiên cứu (hoặc vị trí phù hợp khác). Phải nêu rõ: cách thức sử dụng công cụ + cách thức sử dụng dữ liệu mà công cụ tạo ra. VJES KHÔNG chấp nhận LLM làm đồng tác giả. Khuyến nghị tham chiếu EU AI Act.",
     "source": "Hướng dẫn 2025 — Mục 2.1 điểm 12 + Mục Lưu ý"},

    # --- TLTK ---
    {"id": "VJES_17", "topic": "tltk", "title": "Quy định TLTK",
     "content": "TLTK tuân thủ APA Style 7th Edition. Chỉ liệt kê tài liệu ĐÃ được trích dẫn trong văn bản và ngược lại. TLTK trực tuyến PHẢI có DOI hoặc URL kèm ngày truy cập. KHÔNG dùng footnote.",
     "source": "Hướng dẫn 2025 — Mục 2.3.3 + Mục TLTK"},

    {"id": "VJES_18", "topic": "tltk_format_baibao", "title": "TLTK: Bài báo tạp chí",
     "content": "Bài báo (Anh): Họ, A. (năm). Tên bài viết in đứng. Tên tạp chí in nghiêng, Tập in nghiêng(Số in đứng), Trang đầu-Trang cuối. DOI. VD: Paivio, A. (1975). Perceptual comparisons through the mind's eye. Memory & Cognition, 3(1), 635-647. Bài báo (Việt): Họ và tên đầy đủ. (năm). Tên bài viết in đứng. Tên tạp chí in nghiêng, Tập(Số), trang. VD: Nguyễn Văn Hoàng. (2024). Một số giải pháp nâng cao chất lượng dạy học. Tạp chí Khoa học Giáo dục Việt Nam, 26(2), 12-20.",
     "source": "Hướng dẫn APA 7 — Mục 2.1.1"},

    {"id": "VJES_19", "topic": "tltk_format_baibao", "title": "TLTK: Bài báo nhiều tác giả",
     "content": "2-7 tác giả: Liệt kê đầy đủ tất cả, tác giả cuối có '&'. VD: Grady, J. S., Her, M., Moreno, G., Perez, C., & Yelinek, J. (2019). Trên 20 tác giả: 19 tác giả đầu tiên + dấu cách-chấm-cách-chấm-cách-chấm-cách + tác giả cuối cùng. Tác giả VN: VIẾT ĐẦY ĐỦ HỌ VÀ TÊN (không viết tắt). VD: Nguyễn Văn Hoàng, Lê Văn Ngọc & Trần Anh. (2024).",
     "source": "Hướng dẫn APA 7 — Mục 2.1.2 + 2.1.3"},

    {"id": "VJES_20", "topic": "tltk_format_sach", "title": "TLTK: Sách + Chương sách",
     "content": "Sách (Anh): Họ, A. (năm). Tên sách in nghiêng (Số lần xuất bản, Tập). Nhà xuất bản. DOI. VD: Strunk, W., Jr., Becker, E., & White, E. B. (1979). The guide to everything and then some more stuff (3rd ed.). Macmillan. Chương sách: Họ, A. (năm). Tên chương đứng. In B. Họ (Ed.), Tên sách in nghiêng (trang đầu-cuối). NXB.",
     "source": "Hướng dẫn APA 7 — Mục 2.2 + 2.3"},

    {"id": "VJES_21", "topic": "tltk_format_khac", "title": "TLTK: Báo cáo, Luận án, Hội thảo, Web",
     "content": "Báo cáo: Các tác giả. (năm). Tiêu đề viết nghiêng (Số báo cáo). NXB. URL. Luận án/luận văn: Họ, A. (năm). Tên LA in nghiêng [Doctoral thesis, Tên CSĐT]. Tên cơ sở lưu trữ. URL. Hội thảo: Họ, A. (năm, ngày-ngày tháng). Tiêu đề bài in nghiêng [Paper presentation]. Tên Hội nghị, Thành phố, Quốc gia. Web: Họ, A. (năm). Tiêu đề bài viết in nghiêng. Tên trang in đứng. URL.",
     "source": "Hướng dẫn APA 7 — Mục 2.4, 2.7, 2.8, 2.6"},

    {"id": "VJES_22", "topic": "sap_xep_tltk", "title": "Sắp xếp danh sách TLTK",
     "content": "Sắp xếp ALPHABET theo Họ tác giả đầu tiên (TLTK tiếng nước ngoài). TLTK tiếng Việt: sắp xếp theo TÊN (không phải họ). Cùng tác giả: theo thời gian từ xa đến gần. Cùng tác giả + cùng năm: thêm hậu tố a/b/c... (Berndt, T. J. (1981a), (1981b)). Cùng họ-cùng chữ đầu tên: thêm tên đầy đủ trong ngoặc vuông (Green, L. [Laura]).",
     "source": "Hướng dẫn APA 7 — Mục 3"},

    # --- TRÍCH DẪN ---
    {"id": "VJES_23", "topic": "trich_dan", "title": "Trích dẫn trong văn bản (in-text)",
     "content": "Phương pháp tác giả-năm (Jones, 1998). KHÔNG trực tiếp: chỉ cite tác giả + năm. Trực tiếp/mượn nguyên văn: thêm số trang (Jones, 1998, tr. 199) hoặc (Jones, 1998, tr. 199-201). Trích dẫn dài ≥ 40 từ: tách thành block riêng, thụt lề 1/2 inch, BỎ ngoặc kép. 2 tác giả: Walker và Allen (2004). 3+ tác giả: 'Wasserstein và cộng sự (2005)' hoặc '(Wasserstein & cs, 2005)'.",
     "source": "Hướng dẫn APA 7 — Mục 4.1-4.4"},

    {"id": "VJES_24", "topic": "trich_dan_dac_biet", "title": "Trích dẫn đặc biệt",
     "content": "Nhiều tác phẩm trong 1 ngoặc: sắp xếp alphabet, ngăn cách bằng dấu chấm phẩy. VD: (Bộ Giáo dục, 1996; Durie, 2003; McShane & Travaglione, 2007). Cùng tác giả nhiều năm: liệt kê năm ngăn cách bằng dấu phẩy. VD: (Bộ Giáo dục, 1996, 1999). Trích nguồn thứ cấp (HẠN CHẾ): Wright (được trích dẫn bởi Bragdon, 2013, tr. 223). Chỉ đưa nguồn thứ cấp vào danh sách TLTK; tác phẩm gốc CHỈ trích trong văn bản.",
     "source": "Hướng dẫn APA 7 — Mục 4.9"},

    # --- KIỂM TRA HỒ SƠ ---
    {"id": "VJES_25", "topic": "kiem_tra_ho_so", "title": "Checklist trước khi nộp bài",
     "content": "12 điểm kiểm tra: (1) Phù hợp tôn chỉ VJES; (2) Bài CHƯA xuất bản/đang bình duyệt ở tạp chí khác; (3) Xác định đúng loại bài; (4) Hồ sơ có bản thảo (Word/RTF) + tệp nguồn bảng/hình riêng; (5) Bảng/hình ở định dạng CHỈNH SỬA ĐƯỢC; (6) TLTK online có DOI/URL; (8) Đề xuất người phản biện (tùy chọn); (9) Turnitin ≤25%; (10) Đồng thuận tất cả đồng tác giả; (11) Tuân thủ CC BY-NC-ND 4.0; (12) Khai báo LLM nếu sử dụng.",
     "source": "Hướng dẫn 2025 — Mục 2.1"},

    {"id": "VJES_26", "topic": "nop_bai", "title": "Cách nộp hồ sơ",
     "content": "Email: vjes@vnies.edu.vn (hệ thống online đang bảo trì). Hồ sơ gồm: (i) Bản thảo Word/RTF/Open Office; (ii) Tệp nguồn bảng + hình (nếu có) RIÊNG BIỆT. KHÔNG chấp nhận PDF.",
     "source": "Hướng dẫn 2025 — Mục 2.2"},

    # --- TEMPLATE STRUCTURE ---
    {"id": "VJES_27", "topic": "template", "title": "Template VJES 06.2025 — Cấu trúc",
     "content": "Template có cả 2 ngôn ngữ: (Phần 1 EN) Title EN; Authors EN; Abstract EN + Keywords EN; Author info EN (đơn vị, địa chỉ, chức danh, email, ORCID); (Phần 2 VN) Tên bài VN; Tác giả VN; Tóm tắt VN + Từ khóa VN; Thông tin tác giả VN; (Phần 3) 4 mục chính (1. Đặt vấn đề; 2. PP NC; 3. Kết quả; 4. Kết luận); (Tùy chọn) Thảo luận trước Kết luận; (Cuối) Lời cảm ơn + TLTK.",
     "source": "VJES Article Template V06.2025.docx"},

    {"id": "VJES_28", "topic": "lo_chinh", "title": "Lời cảm ơn",
     "content": "Trình bày lời cảm ơn các đơn vị/tổ chức/cá nhân liên quan kết quả NC trình bày (nếu có). VD: 'Các tác giả trân trọng cảm ơn... vì đã...'; 'Bài báo được tài trợ bởi...'; 'Bài báo là sản phẩm của đề tài KH&CN cấp Bộ của Bộ GD&ĐT, mã số B2024.xxx-yyy'.",
     "source": "Hướng dẫn 2025 — Mục Lời cảm ơn"},

    # --- BẢNG & HÌNH ---
    {"id": "VJES_29", "topic": "bang_hinh", "title": "Bảng + Đồ thị + Hình",
     "content": "Bảng/đồ thị/hình minh họa vẽ bằng gói vẽ phù hợp. Có thể dùng màu. Đặt tại nơi xuất hiện trong văn bản. Đảm bảo vừa lề + có thể thay đổi kích thước không biến dạng. ĐỊNH DẠNG CHỈNH SỬA ĐƯỢC tại mỗi lần nộp/sửa đổi (tránh ảnh PNG flatten).",
     "source": "Hướng dẫn 2025 — Mục Bảng/hình"},

    {"id": "VJES_30", "topic": "thuat_ngu", "title": "Thuật ngữ + Viết tắt",
     "content": "VIẾT TẮT: ngoài cụm từ viết tắt theo thông lệ quốc tế (DSM-5, WHO, GBD, RCT, CBT, ...), VJES quy định KHÔNG VIẾT TẮT trong nội dung bài. Lần đầu xuất hiện: viết đầy đủ + viết tắt trong ngoặc. Sử dụng chữ in nghiêng để nhấn mạnh (KHÔNG dùng bold body text).",
     "source": "Hướng dẫn 2025 — Mục Viết tắt"},
]

# ============================================================
# KG — concepts + edges
# ============================================================
nodes = [
    # === Concepts về tạp chí ===
    {"id": "VJES", "type": "Journal", "label": "Tạp chí KHGD VN", "description": "Vietnam Journal of Educational Sciences, cơ quan ngôn luận VNIES"},
    {"id": "VNIES", "type": "Organization", "label": "Viện KHGD VN", "description": "Vietnam National Institute of Educational Sciences"},
    {"id": "LE_ANH_VINH", "type": "Person", "label": "Lê Anh Vinh", "description": "Viện trưởng VNIES, tổng biên tập VJES"},

    # === Thể loại bài ===
    {"id": "ORIGINAL_RESEARCH", "type": "ArticleType", "label": "Nghiên cứu gốc"},
    {"id": "REVIEW_PAPER", "type": "ArticleType", "label": "Tổng quan tài liệu"},
    {"id": "EXPERIMENTAL", "type": "ArticleType", "label": "Bài báo thực nghiệm"},
    {"id": "METHODOLOGY_PAPER", "type": "ArticleType", "label": "Bài báo phương pháp luận"},
    {"id": "REVIEW_ARTICLE", "type": "ArticleType", "label": "Bài đánh giá"},
    {"id": "CASE_STUDIES", "type": "ArticleType", "label": "Nghiên cứu trường hợp"},

    # === Cấu trúc 4 mục ===
    {"id": "SECTION_1_INTRO", "type": "Section", "label": "1. Giới thiệu / Đặt vấn đề", "description": "BẮT BUỘC; nêu câu hỏi NC, bối cảnh, ý nghĩa"},
    {"id": "SECTION_2_METHOD", "type": "Section", "label": "2. Phương pháp nghiên cứu", "description": "BẮT BUỘC; mô tả PP, công cụ, KHAI BÁO LLM nếu có"},
    {"id": "SECTION_3_RESULT", "type": "Section", "label": "3. Kết quả", "description": "BẮT BUỘC; trình bày kết quả NC"},
    {"id": "SECTION_3_5_DISCUSSION", "type": "Section", "label": "Thảo luận", "description": "KHÔNG BẮT BUỘC; diễn giải + so sánh"},
    {"id": "SECTION_4_CONCLUSION", "type": "Section", "label": "4. Kết luận", "description": "BẮT BUỘC; tóm tắt + đề xuất hướng tương lai"},

    # === Yêu cầu kỹ thuật ===
    {"id": "WORD_LIMIT", "type": "Rule", "label": "5.000-7.000 từ", "description": "8-12 trang A4"},
    {"id": "TURNITIN_LIMIT", "type": "Rule", "label": "Turnitin ≤25%", "description": "Không tính TLTK; tiêu chí sơ loại"},
    {"id": "FONT_FORMAT", "type": "Rule", "label": "Times New Roman 12, line 1.1", "description": "Spacing before/after 3pt"},
    {"id": "HEADING_3_LEVELS", "type": "Rule", "label": "Heading ≤ 3 cấp", "description": "Cấp 1 in đậm thường; cấp 2 nghiêng đậm; cấp 3 nghiêng thường"},
    {"id": "ABSTRACT_LEN", "type": "Rule", "label": "Tóm tắt 150-250 từ", "description": "VN + EN, trả lời 3 câu hỏi"},
    {"id": "KEYWORDS_4_6", "type": "Rule", "label": "Từ khóa 4-6", "description": "VN + EN"},
    {"id": "TITLE_LEN", "type": "Rule", "label": "Tên bài <40 ký tự, ≤12 từ", "description": "Không viết tắt, có 1-2 keyword chính"},
    {"id": "LLM_DECLARATION", "type": "Rule", "label": "Khai báo LLM bắt buộc nếu sử dụng", "description": "Trong mục PP NC; LLM không là đồng tác giả"},
    {"id": "APA_7", "type": "Standard", "label": "APA 7th Edition", "description": "Cho TLTK + trích dẫn"},
    {"id": "ORCID", "type": "Rule", "label": "ORCID khuyến nghị", "description": "https://orcid.org đăng ký miễn phí"},
    {"id": "PUBLICATION_FEE", "type": "Rule", "label": "Phí 2.500.000 VNĐ", "description": "Số tiếng Việt; tiếng Anh miễn phí"},

    # === Quy trình ===
    {"id": "DOUBLE_BLIND_PEER_REVIEW", "type": "Process", "label": "Phản biện kín 2 chiều", "description": "Double-blind"},
    {"id": "SUBMISSION_EMAIL", "type": "Channel", "label": "vjes@vnies.edu.vn", "description": "Hệ thống online đang bảo trì"},

    # === TLTK ===
    {"id": "TLTK_BAIBAO_VN", "type": "TLTK_Format", "label": "TLTK Bài báo VN", "description": "Họ Tên đầy đủ. (năm). Tên bài. Tên tạp chí, Tập(Số), trang."},
    {"id": "TLTK_BAIBAO_EN", "type": "TLTK_Format", "label": "TLTK Bài báo EN", "description": "Họ, A. (năm). Title. Journal, Vol(Iss), pages. DOI"},
    {"id": "TLTK_SACH", "type": "TLTK_Format", "label": "TLTK Sách"},
    {"id": "TLTK_CHUONG_SACH", "type": "TLTK_Format", "label": "TLTK Chương sách"},
    {"id": "TLTK_LUAN_AN", "type": "TLTK_Format", "label": "TLTK Luận án/luận văn"},
    {"id": "TLTK_HOI_THAO", "type": "TLTK_Format", "label": "TLTK Bài hội thảo"},
    {"id": "TLTK_WEB", "type": "TLTK_Format", "label": "TLTK Bài web"},
    {"id": "TLTK_VAN_BAN_PHAP_LY", "type": "TLTK_Format", "label": "TLTK Văn bản pháp lý"},
    {"id": "TLTK_BAO_CAO", "type": "TLTK_Format", "label": "TLTK Báo cáo"},

    # === Chiến lược anti-Turnitin ===
    {"id": "PARAPHRASE_100", "type": "Strategy", "label": "Paraphrase 100%", "description": "Không copy nguyên văn từ tác phẩm đã upload Turnitin"},
    {"id": "WRITE_FROM_SCRATCH", "type": "Strategy", "label": "Viết mới hoàn toàn", "description": "Không tái dùng đoạn từ luận án CTH, báo cáo Can thiệp v5, CLEAN_v3"},
    {"id": "USE_ORIGINAL_NUMBERS", "type": "Strategy", "label": "Số liệu từ nguồn gốc", "description": "Lấy trực tiếp từ paper QT/VN, không dùng số đã hậu xử lý chương 3 luận án"},
    {"id": "FULL_APA_CITATION", "type": "Strategy", "label": "Cite APA đầy đủ", "description": "Turnitin coi cite & quote không tính trùng lặp"},
    {"id": "NATURAL_VN_STYLE", "type": "Strategy", "label": "Câu chữ tự nhiên VN", "description": "Tránh style ChatGPT generic (Furthermore, It is important to note, ...)"},
]

edges = [
    # VJES belongs to VNIES, headed by LAV
    {"source": "VJES", "target": "VNIES", "type": "PUBLISHED_BY"},
    {"source": "LE_ANH_VINH", "target": "VNIES", "type": "HEADS"},
    {"source": "LE_ANH_VINH", "target": "VJES", "type": "EDITOR_IN_CHIEF"},

    # VJES accepts 6 article types
    {"source": "VJES", "target": "ORIGINAL_RESEARCH", "type": "ACCEPTS"},
    {"source": "VJES", "target": "REVIEW_PAPER", "type": "ACCEPTS"},
    {"source": "VJES", "target": "EXPERIMENTAL", "type": "ACCEPTS"},
    {"source": "VJES", "target": "METHODOLOGY_PAPER", "type": "ACCEPTS"},
    {"source": "VJES", "target": "REVIEW_ARTICLE", "type": "ACCEPTS"},
    {"source": "VJES", "target": "CASE_STUDIES", "type": "ACCEPTS"},

    # Review paper structure = 4 sections
    {"source": "REVIEW_PAPER", "target": "SECTION_1_INTRO", "type": "REQUIRES_SECTION"},
    {"source": "REVIEW_PAPER", "target": "SECTION_2_METHOD", "type": "REQUIRES_SECTION"},
    {"source": "REVIEW_PAPER", "target": "SECTION_3_RESULT", "type": "REQUIRES_SECTION"},
    {"source": "REVIEW_PAPER", "target": "SECTION_4_CONCLUSION", "type": "REQUIRES_SECTION"},
    {"source": "REVIEW_PAPER", "target": "SECTION_3_5_DISCUSSION", "type": "OPTIONAL_SECTION"},

    # VJES rules
    {"source": "VJES", "target": "WORD_LIMIT", "type": "ENFORCES"},
    {"source": "VJES", "target": "TURNITIN_LIMIT", "type": "ENFORCES"},
    {"source": "VJES", "target": "FONT_FORMAT", "type": "ENFORCES"},
    {"source": "VJES", "target": "HEADING_3_LEVELS", "type": "ENFORCES"},
    {"source": "VJES", "target": "ABSTRACT_LEN", "type": "ENFORCES"},
    {"source": "VJES", "target": "KEYWORDS_4_6", "type": "ENFORCES"},
    {"source": "VJES", "target": "TITLE_LEN", "type": "ENFORCES"},
    {"source": "VJES", "target": "LLM_DECLARATION", "type": "ENFORCES"},
    {"source": "VJES", "target": "APA_7", "type": "REQUIRES"},
    {"source": "VJES", "target": "ORCID", "type": "RECOMMENDS"},
    {"source": "VJES", "target": "PUBLICATION_FEE", "type": "CHARGES"},

    # Process
    {"source": "VJES", "target": "DOUBLE_BLIND_PEER_REVIEW", "type": "USES_PROCESS"},
    {"source": "VJES", "target": "SUBMISSION_EMAIL", "type": "RECEIVES_VIA"},

    # APA 7 covers all TLTK formats
    {"source": "APA_7", "target": "TLTK_BAIBAO_VN", "type": "DEFINES"},
    {"source": "APA_7", "target": "TLTK_BAIBAO_EN", "type": "DEFINES"},
    {"source": "APA_7", "target": "TLTK_SACH", "type": "DEFINES"},
    {"source": "APA_7", "target": "TLTK_CHUONG_SACH", "type": "DEFINES"},
    {"source": "APA_7", "target": "TLTK_LUAN_AN", "type": "DEFINES"},
    {"source": "APA_7", "target": "TLTK_HOI_THAO", "type": "DEFINES"},
    {"source": "APA_7", "target": "TLTK_WEB", "type": "DEFINES"},
    {"source": "APA_7", "target": "TLTK_VAN_BAN_PHAP_LY", "type": "DEFINES"},
    {"source": "APA_7", "target": "TLTK_BAO_CAO", "type": "DEFINES"},

    # Strategies to satisfy TURNITIN_LIMIT
    {"source": "PARAPHRASE_100", "target": "TURNITIN_LIMIT", "type": "ENABLES_COMPLIANCE"},
    {"source": "WRITE_FROM_SCRATCH", "target": "TURNITIN_LIMIT", "type": "ENABLES_COMPLIANCE"},
    {"source": "USE_ORIGINAL_NUMBERS", "target": "TURNITIN_LIMIT", "type": "ENABLES_COMPLIANCE"},
    {"source": "FULL_APA_CITATION", "target": "TURNITIN_LIMIT", "type": "ENABLES_COMPLIANCE"},
    {"source": "NATURAL_VN_STYLE", "target": "TURNITIN_LIMIT", "type": "ENABLES_COMPLIANCE"},
]


# ============================================================
# Build RAG
# ============================================================
rag = {
    "meta": {
        "name": "VJES Guidelines RAG",
        "version": "1.0",
        "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "source": [
            "bai-bao-khgdvn/huong-dan-tac-gia-tap-chi-khgdvn/1.Tóm tắt Hướng dẫn dành cho tác giả 2025.pdf",
            "bai-bao-khgdvn/huong-dan-tac-gia-tap-chi-khgdvn/3. VJES. Hướng dẫn sắp xếp tài liệu tham khảo.pdf",
            "bai-bao-khgdvn/huong-dan-tac-gia-tap-chi-khgdvn/VJES Article Template V06.2025.docx",
        ],
        "n_chunks": len(chunks),
        "topics": sorted(set(c["topic"] for c in chunks)),
    },
    "chunks": chunks,
}

# Build inverted index for keyword search
inv_index = {}
for c in chunks:
    text = (c["title"] + " " + c["content"] + " " + c["topic"]).lower()
    for word in set(text.split()):
        word = word.strip(".,;:()[]{}—-/")
        if len(word) >= 3:
            inv_index.setdefault(word, []).append(c["id"])
rag["inverted_index"] = {k: sorted(set(v)) for k, v in inv_index.items()}


# ============================================================
# Build KG
# ============================================================
kg = {
    "meta": {
        "name": "VJES Guidelines KG",
        "version": "1.0",
        "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "n_nodes": len(nodes),
        "n_edges": len(edges),
        "node_types": sorted(set(n["type"] for n in nodes)),
        "edge_types": sorted(set(e["type"] for e in edges)),
    },
    "nodes": nodes,
    "edges": edges,
}


# ============================================================
# Write outputs
# ============================================================
OUT_DIR.mkdir(parents=True, exist_ok=True)

(OUT_DIR / "vjes_rag.json").write_text(
    json.dumps(rag, ensure_ascii=False, indent=2), encoding="utf-8"
)
(OUT_DIR / "vjes_kg.json").write_text(
    json.dumps(kg, ensure_ascii=False, indent=2), encoding="utf-8"
)

print(f"[DONE] Wrote {OUT_DIR / 'vjes_rag.json'} ({len(chunks)} chunks)")
print(f"[DONE] Wrote {OUT_DIR / 'vjes_kg.json'} ({len(nodes)} nodes / {len(edges)} edges)")
print()
print("=== TOPICS ===")
for t in rag["meta"]["topics"]:
    print(f"  {t}")
print()
print("=== NODE TYPES ===")
for t in kg["meta"]["node_types"]:
    count = sum(1 for n in nodes if n["type"] == t)
    print(f"  {t}: {count}")
