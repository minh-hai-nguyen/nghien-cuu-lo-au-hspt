# -*- coding: utf-8 -*-
"""Verify every quantitative claim in the critique against the source PDF."""
import os, sys, io, re
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

PDF_TXT = 'C:/Users/HLC/AppData/Local/Temp/vn022_pdf_full.txt'
with open(PDF_TXT, encoding='utf-8') as f:
    pdf = f.read()

# Normalise whitespace
pdf_norm = re.sub(r'\s+', ' ', pdf)
pdf_low = pdf_norm.lower()

print(f'PDF text length: {len(pdf):,}\n')

# Each claim with multiple variants to search for (PDF could have "26.1%", "26.1 %", "26 per cent", etc.)
claims = [
    ('26,1 % nguy cơ trung bình/cao SDQ',
     ['26.1', '26%', '26 per cent', 'moderate or high', 'Total Difficulties']),
    ('30,9 % vấn đề cảm xúc',
     ['30.9', '30 per cent', 'emotional problems', 'emotional symptoms']),
    ('32 % vấn đề bạn bè (peer)',
     ['32%', '32 per cent', 'peer problems', 'peer relationships']),
    ('14,4 % tăng động (hyperactivity)',
     ['14.4', 'hyperactivity', '14 per cent']),
    ('11 % vấn đề hành vi (conduct)',
     ['11%', 'conduct problems', '11 per cent']),
    ('Nữ M=4,92 vs Nam M=4,24 cảm xúc p<0,01',
     ['4.92', '4.24', '4,92', '4,24']),
    ('Nữ M=48,11 vs Nam M=44,27 áp lực học tập',
     ['48.11', '44.27', '48,11', '44,27', '48.1', '44.3']),
    ('50% học thêm >3h/tuần',
     ['50%', 'more than 3 hours', '3 hours per week', 'half']),
    ('15% HS học >9h/tuần',
     ['15%', 'more than 9 hours', '9 hours', 'nine hours']),
    ('28% HS học >3h/đêm',
     ['28%', '3 hours per night', 'three hours', 'study at night']),
    ('52,2% cyberbullying victim',
     ['52.2', '52%', 'cyberbully']),
    ('91% GV lo ngại stress HS',
     ['91%', '91 per cent', 'stress', 'concerned']),
    ('95% GV lo ngại SKTT HS',
     ['95%', 'mental health', '95 per cent']),
    ('86,4% GV chưa nhận đào tạo MHPSS',
     ['86.4', '86%', 'training', 'MHPSS', 'no training']),
    ('70% trường nông thôn thiếu phòng tư vấn',
     ['70%', 'counselling room', 'private', 'rural']),
    ('668 HS',
     ['668', 'n=668', '668 students', '668 adolescents']),
    ('66 GV surveys',
     ['66 teacher', 'n=66', 'teacher surveys', '66 teachers']),
    ('34 KII',
     ['34 KII', '34 key informant', '34 in-depth']),
    ('39 HS FGD',
     ['39 student', '39 FGD']),
    ('21 phụ huynh FGD',
     ['21 parent', '21 FGD', 'parents']),
    ('35 GV FGD',
     ['35 teacher', '35 FGD']),
    ('4 tỉnh Ha Noi, Dien Bien, Gia Lai, Dong Thap',
     ['Ha Noi', 'Dien Bien', 'Gia Lai', 'Dong Thap', 'HCMC', 'Ho Chi Minh City']),
]

print('CLAIM VERIFICATION AGAINST PDF:')
print('='*70)
for claim, variants in claims:
    hits = []
    for v in variants:
        if v.lower() in pdf_low:
            # Find first occurrence with context
            idx = pdf_low.find(v.lower())
            ctx = pdf_norm[max(0, idx-80):idx+80]
            hits.append((v, ctx))
    if hits:
        print(f'\n[OK] {claim}')
        for v, ctx in hits[:2]:
            print(f'    "{v}" -> ...{ctx}...')
    else:
        print(f'\n[MISSING] {claim} — no variants found in PDF')
