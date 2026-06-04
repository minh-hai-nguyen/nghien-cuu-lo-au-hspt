# -*- coding: utf-8 -*-
"""Fix v3 -> v3.1:
1. Replace v2 note ở para 178 (Nhiệm vụ 6.2) — align với chiến lược thầy mới
2. Fix 'né tránh' -> 'tránh né' (2 chỗ trong v3 edits)
Date: 26/05/2026."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.shared import Pt, RGBColor

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
SRC = os.path.join(ROOT, 'Luận án TS', '01_LuanAn_v3_CapNhatCoping_26052026.docx')
OUT = os.path.join(ROOT, 'Luận án TS', '01_LuanAn_v3_1_FixCoping_26052026.docx')

doc = Document(SRC)
ap = doc.paragraphs

# ============================================================
# FIX 1: Replace v2 note Para 178 (Nhiệm vụ 6.2)
# ============================================================
# Tìm paragraph chứa "NCS làm rõ: \"biện pháp đối phó\" được nghiên cứu như BIẾN TRUNG GIAN"
OLD_NOTE = "[NCS làm rõ:"
OLD_NOTE_CONTENT_HINT = "BIẾN TRUNG GIAN"
NEW_NOTE_TEXT = '[NCS làm rõ theo định hướng thầy hướng dẫn: Trong cơ sở lý luận (Chương 1), biện pháp đối phó được trình bày như một YẾU TỐ BẢO VỆ — cùng nhóm với lòng tự trọng, gắn bó trường học và hỗ trợ xã hội. Trong phương pháp (Chương 2) và kết quả (Chương 3), biện pháp đối phó được phân tách thành hai chiều ADAPTIVE và MALADAPTIVE để kiểm định riêng biệt: chiều adaptive (giải quyết vấn đề + tìm kiếm hỗ trợ) kỳ vọng có hệ số β âm với điểm lo âu (tác dụng bảo vệ); chiều maladaptive (tránh né) kỳ vọng có hệ số β dương (tác dụng làm gia tăng triệu chứng). Cách trình bày này vừa giữ khung lý thuyết "ứng phó là yếu tố bảo vệ" như thầy gợi ý, vừa trả lời được câu hỏi của phản biện về tính nhất quán biến.]'

fix1_done = False
for i, p in enumerate(ap):
    if OLD_NOTE_CONTENT_HINT in p.text:
        # Tìm run có chứa text này
        for r in p.runs:
            if OLD_NOTE_CONTENT_HINT in r.text and r.font.color and r.font.color.rgb and str(r.font.color.rgb) == 'C00000':
                # Replace text in this run
                # The run text starts with " [NCS làm rõ: ..." and ends with "...yếu tố bảo vệ/nguy cơ.]"
                old_text = r.text
                # Find start and end positions
                start_idx = old_text.find('[NCS làm rõ:')
                end_idx = old_text.find('.]', start_idx)
                if end_idx != -1:
                    new_run_text = old_text[:start_idx] + NEW_NOTE_TEXT + old_text[end_idx + 2:]
                    r.text = new_run_text
                    # Keep formatting (red, bold)
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(13)
                    r.font.color.rgb = RGBColor(192, 0, 0)
                    r.bold = True
                    r.italic = True
                    fix1_done = True
                    print(f"FIX 1: Replaced v2 note at para {i}")
                    break
        if fix1_done:
            break

if not fix1_done:
    print("FIX 1: Could not find/replace v2 note")

# ============================================================
# FIX 2: 'né tránh' -> 'tránh né' trong RED RUNS v3
# ============================================================
fix2_count = 0
for i, p in enumerate(ap):
    for r in p.runs:
        # Chỉ sửa trong RED RUNS (edits)
        if r.font.color and r.font.color.rgb and str(r.font.color.rgb) == 'C00000':
            if 'né tránh' in r.text:
                old = r.text
                new = r.text.replace('né tránh', 'tránh né')
                r.text = new
                fix2_count += 1
                print(f"FIX 2 (para {i}): 'né tránh' -> 'tránh né'")

print(f"\nFIX 2 total: {fix2_count} replacements")

# ============================================================
# UPDATE HEADER v3 NOTE
# ============================================================
for i, p in enumerate(ap):
    if 'v3 NCS thêm' in p.text:
        for r in p.runs:
            if 'v3 NCS thêm' in r.text:
                r.text = ' [v3.1 NCS sửa lại: 12 đoạn bổ sung CHỮ ĐỎ về biện pháp đối phó adaptive/maladaptive ở Chương 2, Chương 3, Kết luận. Đã cập nhật ghi chú ở Nhiệm vụ NC 6.2 (Mở đầu) align với chiến lược thầy HD: ứng phó vào YTBV ở Chương 1, phân tách 2 chiều ở PP+KQ. Đã sửa thuật ngữ "tránh né" cho nhất quán. Ngày 26/05/2026.]'
                break
        break

# ============================================================
# SAVE
# ============================================================
doc.save(OUT)
print(f"\nDa luu: {OUT}")
print(f"Size: {os.path.getsize(OUT) // 1024}KB")
