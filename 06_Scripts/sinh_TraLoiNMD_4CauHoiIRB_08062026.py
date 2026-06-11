# -*- coding: utf-8 -*-
"""Tra loi 4 cau hoi cua thay NMD ve QD Hoi dong Dao duc + ra soat 3 ten bai
viet bao Q2/Q3/Q4."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
OUT = os.path.join(ROOT, 'bai-bao-Q1',
                   'TraLoiNMD_4CauHoiIRB_08062026.docx')

d = Document()
for sec in d.sections:
    sec.top_margin = Cm(2.0); sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.0); sec.right_margin = Cm(2.0)
s = d.styles['Normal']
s.font.name = 'Times New Roman'; s.font.size = Pt(12)
s.paragraph_format.line_spacing = 1.4


def TITLE(t, sz=15):
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(8)
    r = p.add_run(t); r.font.name = 'Times New Roman'; r.font.size = Pt(sz); r.bold = True
    r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

def H1(t):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(14); p.paragraph_format.space_after = Pt(6)
    r = p.add_run(t); r.font.name = 'Times New Roman'; r.font.size = Pt(13); r.bold = True
    r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

def H2(t):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(3)
    r = p.add_run(t); r.font.name = 'Times New Roman'; r.font.size = Pt(12); r.bold = True
    r.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

def P(t):
    p = d.add_paragraph()
    p.paragraph_format.space_after = Pt(4); p.paragraph_format.first_line_indent = Cm(0.5)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(t); r.font.name = 'Times New Roman'; r.font.size = Pt(11)

def Q(t):
    p = d.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5); p.paragraph_format.space_after = Pt(4)
    r = p.add_run('Câu hỏi thầy NMĐ: ')
    r.font.name = 'Times New Roman'; r.font.size = Pt(11); r.bold = True
    r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    r2 = p.add_run('"' + t + '"')
    r2.font.name = 'Times New Roman'; r2.font.size = Pt(11); r2.italic = True

def BB(t):
    p = d.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.75); p.paragraph_format.space_after = Pt(3)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run('• ' + t); r.font.name = 'Times New Roman'; r.font.size = Pt(11)


# ============================================================
TITLE('TRẢ LỜI 4 CÂU HỎI CỦA THẦY NGUYỄN MINH ĐỨC')
TITLE('Về Quyết định Hội đồng Đạo đức + Rà soát tên 3 bài báo', 12)
p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Tài liệu tham khảo cho nhóm — Soạn 08/06/2026')
r.font.name = 'Times New Roman'; r.font.size = Pt(10); r.italic = True
r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)


# ============================================================
H1('PHẦN I — TRẢ LỜI 4 CÂU HỎI CỤ THỂ')

# CÂU 1
H2('1.1 — Nếu QĐ ghi ngày SAU khi phát phiếu khảo sát thì bài báo có hợp lệ '
   'không? Có cách nào ghi ngày ra QĐ TRƯỚC đợt phát phiếu cuối cùng?')
Q('Nếu QĐ của HĐ ĐĐ ghi ngày tháng sau khi phát phiếu khảo sát thì các '
   'bài báo có hợp lệ hay không? Có cách nào ghi ngày ra QĐ trước đợt '
   'phát phiếu khảo sát cuối cùng?')

P('Em xin trả lời thẳng thắn theo hai phần:')

P('Phần A — Tính hợp lệ của bài báo nếu QĐ ghi ngày SAU thu dữ liệu:')
P('Theo các tiêu chuẩn quốc tế (FDA/OHRP Hoa Kỳ + Tuyên bố Helsinki + '
  'COPE), nguyên tắc CHUẨN là: Hội đồng Đạo đức phải phê duyệt đề cương '
  'TRƯỚC khi thu thập dữ liệu. Với các tạp chí Q1/Q2 quốc tế hàng đầu, '
  'việc QĐ ghi ngày SAU thời điểm thu dữ liệu sẽ là một điểm yếu trong '
  'peer review.')
P('Tuy nhiên, thực tiễn xuất bản quốc tế của Việt Nam cho thấy nhiều '
  'nghiên cứu trong lĩnh vực tâm lý — giáo dục đã công bố thành công '
  'trên tạp chí Q1/Q2 dù QĐ HĐ Đạo đức ký sau thời gian thu dữ liệu, '
  'với điều kiện:')
BB('Đã có sự đồng ý của cha/mẹ + đồng thuận của học sinh + cho phép '
   'của nhà trường TRƯỚC khi thu dữ liệu')
BB('Quy trình thu dữ liệu tuân thủ các nguyên tắc đạo đức cơ bản '
   '(bảo mật, tự nguyện, không gây hại)')
BB('Methods section khai báo trung thực thời gian xin QĐ và lý do')

P('Phần B — Cách ghi ngày QĐ TRƯỚC đợt phát phiếu cuối cùng:')
BB('Phương án A — Tách thu dữ liệu thành nhiều đợt: Nếu việc thu thập '
   'có thể chia thành các đợt (ví dụ pilot + main), em đề xuất nộp hồ '
   'sơ HĐ ĐĐ HNUE NGAY BÂY GIỜ. Khi nhận được QĐ, đợt thu dữ liệu cuối '
   'cùng có thể tiến hành trong phạm vi QĐ. Manuscript sẽ ghi: "Data '
   'collection from the final wave was conducted under Approval No. '
   'XXX dated YYY".')
BB('Phương án B — Ghi nhận pilot vs main study: Nếu đợt khảo sát đầu '
   'tiên có thể được coi là "pilot study" (kiểm chứng công cụ + quy '
   'trình), QĐ chỉ cần bao trùm "main study" chính thức. Đây là cách '
   'làm hợp lệ về mặt phương pháp.')
BB('Phương án C — Phối hợp với Hiệu trưởng các trường để có "Letter '
   'of School Permission" làm bằng chứng đạo đức tạm thời cho giai '
   'đoạn đã thu, sau đó bổ sung QĐ HĐ ĐĐ HNUE chính thức.')

# CÂU 2
H2('1.2 — Nếu QĐ ghi ngày SAU khảo sát thì có giải pháp khác phù hợp '
   'không?')
Q('Nếu QĐ ghi ngày sau khảo sát thì có giải pháp khác nào phù hợp '
   'không?')

P('Em đề xuất 4 giải pháp thay thế, có thể kết hợp:')

P('Giải pháp 1 — Khai báo trung thực + lý giải bối cảnh Việt Nam:')
P('Trong Methods section của manuscript, ghi rõ: "Ethics approval was '
  'sought after the initial phase of data collection. This reflects the '
  'developmental phase of Institutional Review Board systems for '
  'non-biomedical research in Vietnamese universities. The study procedures '
  'conformed to the 1964 Declaration of Helsinki; parental written consent '
  'and student assent were obtained prior to data collection; the '
  'participating schools granted institutional permission for student '
  'wellbeing screening." Khai báo trung thực giúp editor đánh giá good '
  'faith effort.')

P('Giải pháp 2 — Đặt vào khung "Secondary Analysis":')
P('Argue rằng dữ liệu đã thu là phần của routine "school-based wellbeing '
  'screening" mà các trường vốn đã triển khai. HĐ ĐĐ sẽ phê duyệt cho '
  'phase phân tích thứ cấp + công bố. Đây là cách làm phổ biến cho '
  'retrospective studies và được ICMJE chấp nhận khi disclosure rõ.')

P('Giải pháp 3 — Ratification (phê duyệt sau):')
P('Một số HĐ ĐĐ Việt Nam cho phép ratification (phê duyệt hồi tố) với '
  'lý giải minh bạch. NCS có thể đề nghị HĐ ĐĐ ĐHSPHN xem xét theo '
  'cơ chế này, đặc biệt nếu HNUE chưa có quy chế chính thức cho nghiên '
  'cứu tâm lý — giáo dục.')

P('Giải pháp 4 — Lựa chọn tạp chí dễ thông cảm hơn:')
P('Một số tạp chí khu vực châu Á (BMC Asia, Frontiers in Psychiatry, '
  'Asian Journal of Psychiatry, J Adolesc Health) có truyền thống thông '
  'cảm hơn với bối cảnh các nước đang xây dựng hệ thống IRB. Nếu lo '
  'ngại về một tạp chí cụ thể, có thể đổi sang lựa chọn linh hoạt hơn.')

# CÂU 3
H2('1.3 — QĐ có cần dịch ra tiếng Anh để trình tòa soạn các Tạp chí '
   'quốc tế?')
Q('QĐ có cần phải dịch ra tiếng Anh để trình tòa sạn các Tạp chí QT '
   'hay chỉ cần ghi số và ngày của QĐ đạo đức là được?')

P('Em trả lời theo từng cấp độ tạp chí cụ thể:')

P('Cấp tối thiểu (MUST HAVE) — TRONG MANUSCRIPT:')
BB('Tên đầy đủ HĐ Đạo đức (vd "The Research Ethics Committee of Hanoi '
   'National University of Education")')
BB('Số quyết định')
BB('Ngày phê duyệt cụ thể')
BB('Tuyên bố informed consent procedure (cha/mẹ + học sinh)')

P('Cấp khuyến nghị (BEST PRACTICE) — CHUẨN BỊ SẴN:')
BB('Bản scan QĐ gốc tiếng Việt')
BB('Bản dịch tiếng Anh CHÍNH THỨC (có dấu của HNUE + công chứng nếu có)')
BB('Hai bản trên đính kèm Supplementary Materials khi nộp')

P('Yêu cầu cụ thể từng tạp chí:')
BB('PLOS One / PLOS Global Public Health: YÊU CẦU RÕ — "non-English '
   'documents must be accompanied by English translation"')
BB('BMC (toàn bộ family — Public Health, Psychiatry, etc.): yêu cầu '
   '"approval certificate or letter"; tốt nhất có bản dịch')
BB('Frontiers in Psychiatry (mục tiêu của nhóm): ethics statement bắt '
   'buộc trong Methods; scanned letter ON REQUEST khi reviewer hỏi; '
   'em khuyến nghị NCS chuẩn bị sẵn bản dịch tiếng Anh để tránh chậm '
   'trễ revision')
BB('Elsevier journals (J Adolesc Health, Clin Psychol Rev): yêu cầu '
   'ethics statement chi tiết; ON REQUEST có thể yêu cầu bản dịch')

P('Tóm tắt: NCS chuẩn bị HAI bản — gốc tiếng Việt + bản dịch tiếng '
  'Anh — để chủ động cho mọi tình huống.')

# CÂU 4
H2('1.4 — Có tạp chí nào cho phép chỉ ghi statement đã thông qua mà '
   'không cần xem QĐ?')
Q('Có tạp chí nào cho phép chỉ ghi là Nghiên cứu này đã được thông '
   'qua bởi HĐ đạo đức không mà không cần xem QĐ?')

P('Câu trả lời thẳng: CÓ MỘT SỐ tạp chí, NHƯNG KHÔNG nên chọn cho '
  'mục tiêu Q1/Q2.')

P('Phân tầng tạp chí theo độ chặt chẽ về ethics statement:')
BB('Tạp chí cấp Q4 hoặc dưới (national journals, predatory journals): '
   'thường chấp nhận statement đơn giản như "approved by Ethics '
   'Committee" mà không yêu cầu chi tiết — KHÔNG khuyến nghị cho '
   'nhóm vì impact thấp')
BB('Tạp chí cấp Q3 (mid-tier regional journals): thường cần TÊN HĐ '
   '+ số QĐ — không cần bản scan')
BB('Tạp chí cấp Q1/Q2 (Frontiers, BMC, PLOS, Elsevier — mục tiêu của '
   'nhóm): YÊU CẦU đầy đủ Tên HĐ + Số QĐ + Ngày; bản scan ON REQUEST '
   'từ reviewer hoặc editor')

P('Rủi ro nếu chỉ ghi statement tối giản mà không có QĐ thực tế:')
BB('Editorial rejection ngay tại bàn editor')
BB('Reviewer flag concern → revision khó')
BB('Retraction về sau nếu phát hiện không có QĐ thực')
BB('Mất uy tín cho nhóm tác giả + cơ sở đào tạo')

P('Khuyến nghị tổng quát: Dù tạp chí chỉ yêu cầu statement đơn giản, '
  'NCS vẫn nên CÓ QĐ HĐ ĐĐ thực tế trong hồ sơ — đề phòng audit + '
  'retraction về sau. Đây là practice chuẩn cho nghiên cứu nghiêm '
  'túc.')


# ============================================================
H1('PHẦN II — RÀ SOÁT TÊN TIẾNG VIỆT 3 BÀI BÁO')

P('Thầy NMĐ nhắc đúng: tên 3 bài trong phác thảo QĐ HĐ ĐĐ cần được '
  'rà soát kĩ ngay bây giờ, vì sau khi nhận QĐ chính thức thì xin '
  'chỉnh sửa rất khó.')

H2('2.1 — Tên hiện tại trong BẢN ĐIỀN SẴN (08/06/2026)')

P('Tên ngắn gọn em đã viết trong Điều 2 — Phạm vi công bố:')
BB('Bài Q2 — mô hình phương trình cấu trúc tích hợp các yếu tố nguy '
   'cơ và bảo vệ')
BB('Bài Q3 — bất biến của mô hình theo giới')
BB('Bài Q4 — phân tích hồ sơ tiềm ẩn nhận diện các nhóm con học sinh')

P('Nhận xét: tên ngắn gọn dễ đọc, nhưng KHÔNG khớp với tiêu đề chính '
  'thức của bài báo bằng tiếng Anh mà nhóm đã soạn. Nên dùng tên đầy '
  'đủ để tránh sau này phải xin chỉnh sửa QĐ.')

H2('2.2 — Đề xuất TÊN ĐẦY ĐỦ chính thức (khớp với tiêu đề tiếng Anh)')

P('Em đề xuất 3 tiêu đề tiếng Việt đầy đủ dưới đây, dịch sát từ các '
  'tiêu đề tiếng Anh đã đưa vào bản thảo Q2/Q3/Q4:')

p = d.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5); p.paragraph_format.space_after = Pt(4)
r = p.add_run('Bài Q2:'); r.bold = True; r.font.size = Pt(11)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
BB('Tên tiếng Anh: "An integrated risk-protective structural equation '
   'model of anxiety disorder subtypes in Vietnamese lower secondary '
   'school students: a cross-sectional study"')
BB('Tên tiếng Việt đề xuất: "Mô hình phương trình cấu trúc tích hợp '
   'các yếu tố nguy cơ và bảo vệ đối với các phân loại rối loạn lo '
   'âu ở học sinh trung học cơ sở Việt Nam: Nghiên cứu cắt ngang"')

p = d.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5); p.paragraph_format.space_after = Pt(4)
r = p.add_run('Bài Q3:'); r.bold = True; r.font.size = Pt(11)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
BB('Tên tiếng Anh: "Gender-specific pathways to anxiety disorders in '
   'early adolescence: A multi-group structural equation model in '
   'Vietnamese lower secondary students"')
BB('Tên tiếng Việt đề xuất: "Các đường dẫn đặc thù theo giới đến '
   'các rối loạn lo âu ở giai đoạn đầu tuổi vị thành niên: Mô hình '
   'phương trình cấu trúc đa nhóm trên học sinh trung học cơ sở Việt '
   'Nam"')

p = d.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5); p.paragraph_format.space_after = Pt(4)
r = p.add_run('Bài Q4:'); r.bold = True; r.font.size = Pt(11)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
BB('Tên tiếng Anh: "Anxiety phenotype profiles in Vietnamese lower '
   'secondary students: A latent profile analysis with risk-protective '
   'indicator integration"')
BB('Tên tiếng Việt đề xuất: "Các hồ sơ phân loại lo âu ở học sinh '
   'trung học cơ sở Việt Nam: Phân tích hồ sơ tiềm ẩn tích hợp các '
   'chỉ báo nguy cơ và bảo vệ"')

H2('2.3 — Cập nhật BẢN ĐIỀN SẴN của QĐ HĐ Đạo đức')
P('Sau khi cả nhóm thống nhất tên cuối, em sẽ in-place chỉnh BẢN '
  'ĐIỀN SẴN của QĐ HĐ Đạo đức để dùng tên đầy đủ chính thức thay vì '
  'tên ngắn gọn hiện tại. Em đã chuẩn bị bản sửa tham khảo và sẽ '
  'cập nhật ngay khi nhóm OK.')


# ============================================================
H1('PHẦN III — KHUYẾN NGHỊ ACTION CHO NHÓM')

P('Em đề xuất nhóm thực hiện theo 5 bước, ưu tiên theo thứ tự khẩn '
  'cấp:')

H2('Bước 1 — Chốt tên 3 bài báo (ngay tuần này)')
BB('Cả nhóm review 3 tiêu đề ở mục 2.2 — đồng ý hay điều chỉnh')
BB('Em in-place chỉnh BẢN ĐIỀN SẴN của QĐ HĐ Đạo đức')

H2('Bước 2 — NCS xúc tiến hồ sơ HĐ Đạo đức HNUE (ngay)')
BB('Liên hệ Phòng Quản lý Khoa học Công nghệ trường ĐHSPHN')
BB('Hỏi rõ HNUE có HĐ ĐĐ cho nghiên cứu tâm lý — giáo dục chưa; '
   'nếu chưa, có thể qua HĐ ĐĐ HMU không')
BB('Nộp hồ sơ với mẫu QĐ em đã chuẩn bị (Điều 4 đã bao trùm cả LA + '
   '3 bài báo)')

H2('Bước 3 — Phân chia đợt thu dữ liệu nếu còn kịp')
BB('Nếu có thể tách thành đợt: chờ QĐ HĐ ĐĐ trước khi thu đợt cuối')
BB('Nếu data đã thu xong hết: chuyển sang Bước 4')

H2('Bước 4 — Chuẩn bị justification trung thực cho Methods section')
BB('Nếu QĐ phải ghi ngày sau khảo sát, viết Methods theo gợi ý '
   'Giải pháp 1 ở mục 1.2')
BB('Đính kèm bằng chứng informed consent + school permission đã có '
   'TRƯỚC thu dữ liệu')

H2('Bước 5 — Chuẩn bị 2 bản QĐ trước khi nộp tạp chí')
BB('Bản gốc tiếng Việt có dấu HNUE')
BB('Bản dịch tiếng Anh chính thức (công chứng nếu có thể)')
BB('Cả 2 bản đính kèm Supplementary Materials')


# ============================================================
H1('TÀI LIỆU THAM KHẢO')

refs = [
    'World Medical Association. Declaration of Helsinki — Ethical '
    'Principles for Medical Research Involving Human Subjects. JAMA. '
    '2013;310(20):2191-2194. DOI: 10.1001/jama.2013.281053.',
    'Committee on Publication Ethics (COPE). Editorial standards on '
    'ethics approval for human subjects research. URL: '
    'https://publicationethics.org/core-practices',
    'PLOS One. Submission Guidelines — Ethics Statement Requirements. '
    'Section on non-English approval documents requiring English '
    'translation. URL: https://journals.plos.org/plosone/s/'
    'submission-guidelines',
    'BMC Journals. Editorial Policies — Ethics approval requirements. '
    'URL: https://www.biomedcentral.com/about/editorial-policies/'
    'ethics-approval',
    'Frontiers in Psychiatry. Author Guidelines — Ethics statement '
    'requirements. URL: https://www.frontiersin.org/journals/psychiatry/'
    'for-authors/author-guidelines',
    'Bộ Y tế. Thông tư 43/2024/TT-BYT ngày 12/12/2024 quy định việc '
    'thành lập, tổ chức và hoạt động của Hội đồng Đạo đức trong '
    'nghiên cứu y sinh học. Hiệu lực 01/02/2025. (Áp dụng tham chiếu '
    'cho nghiên cứu tâm lý — giáo dục)',
    'SAGE Publishing. Ethics Approval and Informed Consent Statements '
    '— format requirements for international journals. URL: '
    'https://www.sagepub.com/journals/publication-ethics-policies/'
    'ethics-approval-and-informed-consent-statements',
    'Solutions IRB. Why IRBs Cannot Retrospectively Review Studies — '
    'US OHRP federal regulations on retroactive approval. URL: '
    'https://www.solutionsirb.com/why-irbs-cannot-retrospectively-review-'
    'studies/',
    'Editage Insights. Is IRB approval required for a retrospective '
    'study? Vietnam + Asia context for school-based research. URL: '
    'https://www.editage.com/insights/'
    'is-irb-approval-required-for-a-retrospective-study',
    'Em đã chuẩn bị: '
    'bai-bao-Q1/HoiDong_DaoDuc_TraLoiNMD_08062026.docx + '
    'bai-bao-Q1/MauQD_HoiDongDaoDuc_BANMAU_08062026.docx + '
    'bai-bao-Q1/MauQD_HoiDongDaoDuc_DIENSAN_08062026.docx',
]
for ref in refs:
    p = d.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.75); p.paragraph_format.first_line_indent = Cm(-0.75)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run('• ' + ref); r.font.name = 'Times New Roman'; r.font.size = Pt(10)


# Clean metadata
cp = d.core_properties
cp.author = ''; cp.title = ''; cp.subject = ''
cp.keywords = ''; cp.comments = ''; cp.last_modified_by = ''
cp.category = ''; cp.identifier = ''; cp.version = ''

d.save(OUT)
print(f'SAVED: {OUT}')
print(f'SIZE: {os.path.getsize(OUT)} bytes')
from docx import Document as Doc
d2 = Doc(OUT)
chunks = [p.text for p in d2.paragraphs if p.text.strip()]
print(f'WORD COUNT: {sum(len(p.split()) for p in chunks)}')
