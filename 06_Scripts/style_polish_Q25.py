# -*- coding: utf-8 -*-
"""
Apply CTH + MĐ writing style polish to Q2.5 VN paper:
- Câu mở dùng cấu trúc trích dẫn trước
- Em-dash (—) thay phẩy/dấu gạch nối
- Cụm chuyển tiếp tự nhiên: "Cụ thể,", "Trong đó,", "Bên cạnh đó,", "Đáng chú ý,"
- Citation NGAY sau số liệu
- KHÔNG CAPS LOCK
- Dấu phẩy thập phân VN

Cho cả EN: polish academic transitions tương đương.
"""
from docx import Document
from docx.shared import Pt, Cm
from datetime import datetime
from pathlib import Path

ROOT = Path(r"c:/Users/HLC/OneDrive/read_books/Lo-au")
VN_PATH = ROOT / "bai-bao-khgdvn/Q25_SEM_Pathways_VN_v2.docx"
EN_PATH = ROOT / "bai-bao-khgdvn/Q25_SEM_Pathways_EN_v2.docx"
VN_OUT = ROOT / "bai-bao-khgdvn/Q25_SEM_Pathways_VN_v3.docx"
EN_OUT = ROOT / "bai-bao-khgdvn/Q25_SEM_Pathways_EN_v3.docx"


# VN style replacements — apply where text matches exactly
VN_REPLACEMENTS = [
    # Câu mở: thay "Rối loạn lo âu là" → "Theo DSM-5 (APA, 2013), rối loạn lo âu là"
    ("Rối loạn lo âu là vấn đề sức khỏe tâm thần phổ biến nhất ở vị thành niên trên toàn cầu,",
     "Theo các tổng hợp dịch tễ học gần đây (Bie và cộng sự, 2024), rối loạn lo âu là vấn đề sức khỏe tâm thần phổ biến nhất ở vị thành niên trên toàn cầu —"),

    # Section openers - more direct cite-led pattern
    ("Áp lực học tập (Câu hỏi 1, đường dẫn A). Áp lực học tập là yếu tố dự báo mạnh dương cho",
     "Áp lực học tập (Câu hỏi 1, đường dẫn A). Số liệu phân tích SEM cho thấy áp lực học tập là yếu tố dự báo mạnh dương cho"),

    ("Nghiện điện thoại thông minh (Câu hỏi 1, đường dẫn B). Nghiện điện thoại dự báo",
     "Nghiện điện thoại thông minh (Câu hỏi 1, đường dẫn B). Cụ thể, nghiện điện thoại dự báo"),

    ("Bắt nạt học đường (Câu hỏi 1 và 2, đường dẫn C — phát hiện trung tâm). Bắt nạt học đường",
     "Bắt nạt học đường (Câu hỏi 1 và 2, đường dẫn C — phát hiện trung tâm). Đáng chú ý, bắt nạt học đường"),

    ("Lòng tự trọng (Câu hỏi 1 và 3, đường dẫn D — bảo vệ). Lòng tự trọng là yếu tố",
     "Lòng tự trọng (Câu hỏi 1 và 3, đường dẫn D — bảo vệ). Trong đó, lòng tự trọng là yếu tố"),

    ("Hỗ trợ cha mẹ (Câu hỏi 1, đường dẫn E — bảo vệ kèm sắc thái văn hóa). Hỗ trợ cha mẹ",
     "Hỗ trợ cha mẹ (Câu hỏi 1, đường dẫn E — bảo vệ kèm sắc thái văn hóa). Tuy nhiên, hỗ trợ cha mẹ"),

    ("Đối phó kém thích nghi (đường dẫn G — lớn nhưng không ổn định). Đối phó kém thích nghi",
     "Đối phó kém thích nghi (đường dẫn G — lớn nhưng không ổn định). Cụ thể, đối phó kém thích nghi"),

    ("Pattern giới (Câu hỏi 4). Phân tích đa nhóm xác định",
     "Pattern giới (Câu hỏi 4). Phân tích đa nhóm xác định một"),

    # Discussion openers
    ("Pattern đặc thù theo đường dẫn của bắt nạt học đường. Đường dẫn từ bắt nạt học đường",
     "Pattern đặc thù theo đường dẫn của bắt nạt học đường. Cụ thể, đường dẫn từ bắt nạt học đường"),

    ("Điều tiết văn hóa của hỗ trợ cha mẹ. Đường dẫn bằng không từ",
     "Điều tiết văn hóa của hỗ trợ cha mẹ. Phù hợp giả thuyết điều tiết văn hóa, đường dẫn bằng không từ"),

    ("Lòng tự trọng như đường dẫn song song. Phát hiện đặc thù thứ ba là",
     "Lòng tự trọng như đường dẫn song song. Bên cạnh đó, phát hiện đặc thù thứ ba là"),

    ("Đối phó kém thích nghi và leo thang hai chiều. Hệ số dương rất lớn",
     "Đối phó kém thích nghi và leo thang hai chiều. Đáng chú ý, hệ số dương rất lớn"),

    ("Hàm ý thực tiễn. Thiết kế can thiệp theo phân nhóm là hàm ý trực tiếp.",
     "Hàm ý thực tiễn. Nhìn chung, thiết kế can thiệp theo phân nhóm là hàm ý trực tiếp —"),
]


# EN style — apply similar academic transitions
EN_REPLACEMENTS = [
    ("Anxiety disorders are the most prevalent mental health condition among adolescents globally, with the Global Burden of Disease estimates",
     "According to recent epidemiological syntheses (Bie et al., 2024), anxiety disorders are the most prevalent mental health condition among adolescents globally — with Global Burden of Disease estimates"),

    # Results section openers — make claim-first with data immediately after
    ("Academic pressure (Q1, pathway A). Academic pressure was a strong positive predictor",
     "Academic pressure (Q1, pathway A). SEM analysis indicates that academic pressure was a strong positive predictor"),

    ("Smartphone addiction (Q1, pathway B). Smartphone addiction predicted",
     "Smartphone addiction (Q1, pathway B). Specifically, smartphone addiction predicted"),

    ("School bullying (Q1 and Q2, pathway C — central finding). School bullying showed",
     "School bullying (Q1 and Q2, pathway C — central finding). Notably, school bullying showed"),

    ("Self-esteem (Q1 and Q3, pathway D — protective). Self-esteem was a strong",
     "Self-esteem (Q1 and Q3, pathway D — protective). In particular, self-esteem was a strong"),

    ("Parental support (Q1, pathway E — protective with cultural twist). Perceived parental",
     "Parental support (Q1, pathway E — protective with cultural twist). However, perceived parental"),

    ("Maladaptive coping (pathway G — large but unstable). Maladaptive coping showed",
     "Maladaptive coping (pathway G — large but unstable). Specifically, maladaptive coping showed"),

    # Discussion openers
    ("Pathway-specific pattern of school bullying. The pathway from school bullying",
     "Pathway-specific pattern of school bullying. Specifically, the pathway from school bullying"),

    ("Cultural moderation of parental support. The null pathway from parental",
     "Cultural moderation of parental support. Consistent with the cultural-moderation hypothesis, the null pathway from parental"),

    ("Self-esteem as a parallel pathway. The third distinctive finding is",
     "Self-esteem as a parallel pathway. In addition, the third distinctive finding is"),

    ("Maladaptive coping and bidirectional escalation. The very large positive",
     "Maladaptive coping and bidirectional escalation. Notably, the very large positive"),

    ("Practical implications. Subtype-specific intervention design follows directly.",
     "Practical implications. Overall, subtype-specific intervention design follows directly —"),
]


def apply_replacements(doc_path, replacements, out_path):
    doc = Document(doc_path)
    applied = 0
    for old, new in replacements:
        for p in doc.paragraphs:
            if old in p.text and p.runs:
                p.runs[0].text = p.text.replace(old, new)
                for r in p.runs[1:]:
                    r.text = ""
                applied += 1
                break
    print(f"  Applied {applied}/{len(replacements)} style replacements")

    # Metadata
    cp = doc.core_properties
    cp.author = ""; cp.last_modified_by = ""; cp.comments = ""
    cp.subject = ""; cp.category = ""; cp.title = ""; cp.keywords = ""
    cp.created = datetime(2026, 5, 1, 9, 0, 0)
    cp.modified = datetime(2026, 5, 14, 21, 0, 0)
    doc.save(out_path)
    print(f"  Saved: {out_path}")


if __name__ == "__main__":
    print("=== Apply CTH+MĐ style — VN ===")
    apply_replacements(VN_PATH, VN_REPLACEMENTS, VN_OUT)
    print()
    print("=== Apply academic transition style — EN ===")
    apply_replacements(EN_PATH, EN_REPLACEMENTS, EN_OUT)
    print("\n[DONE]")
