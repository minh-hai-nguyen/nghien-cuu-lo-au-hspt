# -*- coding: utf-8 -*-
"""
(1) Bài 1 v6: sửa lỗi abstract tiếng Anh — V-NAMHS (2,3% DSM-5 lo âu, không phải ~20%).
(2) Bài 03 v4: xóa dấu hiệu AI/đánh dấu —
    - alt-text hình "AI-generated content may be incorrect" -> rỗng
    - màu chữ EE0000 (đỏ) + 0000FF (xanh) -> đen.
File gốc giữ nguyên (chỉ Bài 03 v4 = bản đang hoàn thiện).
"""
import os, zipfile
from docx import Document

ROOT = r"c:/Users/HLC/OneDrive/read_books/Lo-au/bai-bao-khgdvn"
B1 = os.path.join(ROOT, "Bai1_YTNC_HSTHCS_v6_16052026.docx")
B03 = os.path.join(ROOT, "bài-03", "Công Thị Hằng_v4_đã sửa.docx")


# ---- (1) Bài 1 EN abstract ----
OLD = ("the National Adolescent Mental Health Survey (V-NAMHS) reported that "
       "nearly one in five children and adolescents met DSM-5 diagnostic criteria "
       "for at least one anxiety disorder, yet")
NEW = ("the National Adolescent Mental Health Survey (V-NAMHS) reported that "
       "2.3% of adolescents met full DSM-5 diagnostic criteria for an anxiety "
       "disorder while approximately 21.7% experienced at least one mental "
       "health problem, yet")
d = Document(B1)
ok = False
for p in d.paragraphs:
    if OLD in p.text:
        if p.runs:
            p.runs[0].text = p.text.replace(OLD, NEW)
            for r in p.runs[1:]:
                r.text = ""
        ok = True
        break
print(f"(1) Bài 1 EN abstract V-NAMHS: {'OK' if ok else 'NOT FOUND'}")
if ok:
    d.save(B1)


# ---- (2) Bài 03: dọn document.xml ----
zin = zipfile.ZipFile(B03, "r")
items = zin.infolist()
docxml = zin.read("word/document.xml").decode("utf-8")
parts = {it.filename: zin.read(it.filename) for it in items}
zin.close()

before = docxml
docxml = docxml.replace(
    'descr="A diagram of a flowchart&#10;&#10;AI-generated content may be incorrect."',
    'descr=""')
n_ai = before.count("AI-generated content may be incorrect") - \
       docxml.count("AI-generated content may be incorrect")
n_red = docxml.count('w:val="EE0000"')
n_blue = docxml.count('w:val="0000FF"')
docxml = docxml.replace('w:val="EE0000"', 'w:val="000000"')
docxml = docxml.replace('w:val="0000FF"', 'w:val="000000"')
print(f"(2) Bài 03: alt-text AI xóa {n_ai} | màu đỏ EE0000 {n_red} | màu xanh 0000FF {n_blue}")

zout = zipfile.ZipFile(B03, "w", zipfile.ZIP_DEFLATED)
for it in items:
    data = docxml.encode("utf-8") if it.filename == "word/document.xml" else parts[it.filename]
    zout.writestr(it, data)
zout.close()

# verify
z = zipfile.ZipFile(B03)
v = z.read("word/document.xml").decode("utf-8")
print(f"    Verify Bài 03 — còn 'AI-generated': {v.count('AI-generated content may be incorrect')}"
      f" | EE0000: {v.count('EE0000')} | 0000FF: {v.count('0000FF')}")
print("[DONE]")
