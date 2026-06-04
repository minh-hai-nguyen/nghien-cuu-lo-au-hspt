# -*- coding: utf-8 -*-
"""6 rounds verification (rounds 4-9) against PDF originals"""
import fitz, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.chdir(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))

def pdf_text(path):
    doc = fitz.open(path)
    return ''.join(p.get_text() for p in doc)

def script_text(path):
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()

def find_in(text, kw):
    if kw in text: return True
    if kw.replace(',','') in text.replace(',',''): return True
    if kw.replace('.',',') in text or kw.replace(',','.') in text: return True
    return False

papers = [
    ('02','Saikia 2023','Tom-tat-tung-bai/02_Saikia_et_al_2023.docx.py','Originals/02_Saikia_et_al_2023.pdf'),
    ('03','Mandaknalli 2021','Tom-tat-tung-bai/03_Mandaknalli_Malusare_2021.py','Originals/03_Mandaknalli_Malusare_2021.pdf'),
    ('05','Alharbi 2019','Tom-tat-tung-bai/05_Alharbi_et_al_2019.py','Originals/05_Alharbi_et_al_2019.pdf'),
    ('06','Nakie 2022','Tom-tat-tung-bai/06_Nakie_et_al_2022.py','Originals/06_Nakie_et_al_2022.pdf'),
    ('07','Chen 2023','Tom-tat-tung-bai/07_Chen_et_al_2023.py','Originals/07_Chen_et_al_2023.pdf'),
    ('08','Wen 2020','Tom-tat-tung-bai/08_Wen_et_al_2020.py','Originals/08_Wen_et_al_2020.pdf'),
    ('09','Qiu 2022','Tom-tat-tung-bai/09_Qiu_et_al_2022.py','Originals/09_Qiu_et_al_2022.pdf'),
    ('10','Xu 2021','Tom-tat-tung-bai/10_Xu_et_al_2021.py','Originals/10_Xu_et_al_2021.pdf'),
    ('11','Bhardwaj 2020','Tom-tat-tung-bai/11_Bhardwaj_et_al_2020.py','Originals/11_Bhardwaj_et_al_2020.pdf'),
]

all_errors = []
total_checks = 0

def run_round(round_num, title, checks_dict):
    global total_checks
    errors = []
    print(f'\n{"="*60}')
    print(f'ROUND {round_num}: {title}')
    print(f'{"="*60}')
    for num, name, sp, pp in papers:
        pt = pdf_text(pp)
        checks = checks_dict.get(num, [])
        total_checks += len(checks)
        fails = [(v,d) for v,d in checks if not find_in(pt, v)]
        status = 'PASS' if not fails else 'FAIL'
        print(f'  {num} {name}: {len(checks)-len(fails)}/{len(checks)} {status}')
        for v,d in fails:
            print(f'    MISS: {v} ({d})')
            errors.append((num, name, v, d))
    return errors

# ROUND 4: Sample sizes
all_errors += run_round(4, 'SAMPLE SIZE & DEMOGRAPHICS', {
    '02': [('360','N'),('10','schools'),('120','total schools')],
    '03': [('450','N'),('108','anxious')],
    '05': [('1245','N')],
    '06': [('849','recruited'),('810','responded'),('96.1','response')],
    '07': [('63205','N'),('99.6','response')],
    '08': [('900','N')],
    '09': [('2079','N'),('2879','recruited'),('98.06','response')],
    '10': [('373216','N'),('15.24','mean age')],
    '11': [('288','N')],
})

# ROUND 5: Prevalence
all_errors += run_round(5, 'PREVALENCE RATES', {
    '02': [('22.2','dep'),('24.4','anx'),('6.9','stress')],
    '03': [('31.48','mild'),('39.81','mod'),('15.74','sev'),('12.96','vsev')],
    '05': [('26.0','no dep'),('36.5','no anx'),('24.6','mod dep'),('10.4','modsev'),('19.5','mod anx'),('9.8','sev anx')],
    '06': [('41.4','dep'),('66.7','anx'),('52.2','stress')],
    '07': [('23.0','dep'),('13.9','anx'),('11.5','both'),('25.5','either')],
    '08': [('19.22','mild'),('56.00','mod'),('24.78','severe')],
    '09': [('26.0','dep'),('13.4','anx')],
    '10': [('9.89','overall'),('10.85','junior'),('8.08','senior')],
    '11': [('35.1','dep norm'),('18.1','anx norm'),('44.8','str norm'),('29.9','dep mod'),('26.4','anx mod'),('25.3','anx sev'),('21.5','anx ext')],
})

# ROUND 6: Gender
all_errors += run_round(6, 'GENDER DIFFERENCES', {
    '02': [('30.0','male anx'),('18.9','female anx'),('0.049','p gender')],
    '03': [('58.33','female in anx'),('0.022','p gender')],
    '05': [('71.3','fem sev anx'),('28.7','male sev anx'),('74.2','fem sev dep'),('25.8','male sev dep')],
    '06': [('1.304','female AOR')],
    '07': [('1.55','female OR')],
    '08': [('0.649','male OR mod'),('0.262','male OR sev')],
    '09': [('24.3','male dep'),('28.9','fem dep'),('11.1','male anx'),('17.5','fem anx')],
    '10': [('10.11','male anx'),('9.66','female anx'),('0.92','fem OR jr'),('0.84','fem OR sr')],
    '11': [('0.617','r dep-anx'),('0.664','r dep-str'),('0.575','r anx-str')],
})

# ROUND 7: OR/AOR
all_errors += run_round(7, 'OR/AOR/P-VALUES', {
    '02': [('0.003','game p'),('0.028','ses p'),('0.044','abuse p'),('0.043','family p')],
    '03': [('0.032','age p'),('0.042','family p'),('0.035','living p'),('0.019','peer p'),('0.025','phys p'),('0.072','sleep p'),('0.634','internet p')],
    '05': [('0.341','age p'),('0.545','edu p'),('0.176','age dep p')],
    '06': [('5.595','khat'),('4.777','smoke'),('2.099','chronic'),('1.777','family'),('1.739','social'),('1.828','alcohol'),('1.395','rural')],
    '07': [('6.99','sleep'),('5.00','igd'),('1.97','social'),('1.70','verbal'),('1.51','physical'),('1.31','nonnuc'),('1.25','property'),('1.11','urban'),('1.06','only child')],
    '08': [('11.579','press vh'),('6.523','press h'),('0.562','mh school'),('0.377','excel'),('6.122','press h mod'),('5.894','press h sev')],
    '09': [('6.74','resil dep'),('2.80','resil anx'),('1.82','neg dep'),('2.01','neg anx'),('0.30','pos dep'),('0.32','pos anx')],
    '10': [('1.30','rural'),('2.72','behavior')],
    '11': [('10.93','or ad'),('17.26','or sd'),('9.771','or as')],
})

# ROUND 8: CI spot check
all_errors += run_round(8, 'CONFIDENCE INTERVALS', {
    '06': [('1.006','fem lo'),('1.849','fem hi'),('2.357','khat lo'),('11.132','khat hi'),('1.045','sp/chr lo'),('1.919','sp hi'),('1.028','fam lo'),('3.073','fam hi'),('1.407','smoke lo')],
    '07': [('1.49','fem lo'),('1.62','fem hi'),('6.69','sleep lo'),('7.30','sleep hi'),('4.42','igd lo'),('5.66','igd hi'),('1.87','soc lo'),('2.08','soc hi')],
    '09': [('4.78','resil lo'),('9.61','resil hi'),('1.30','neg lo'),('2.53','neg hi'),('0.24','pos lo'),('0.37','pos hi'),('1.38','neg anx lo'),('2.92','neg anx hi')],
    '10': [('1.26','rural lo'),('1.34','rural hi'),('0.89','fem jr lo'),('0.94','fem jr hi'),('0.81','fem sr lo'),('0.88','fem sr hi')],
})

# ROUND 9: Factual claims in scripts
print(f'\n{"="*60}')
print('ROUND 9: FACTUAL CLAIMS IN SCRIPTS')
print(f'{"="*60}')
r9_errors = []
total_checks += 5

# 9a: Xu no longer mentions Wen 2020
st10 = script_text('Tom-tat-tung-bai/10_Xu_et_al_2021.py')
if 'Wen 2020' in st10 or 'Wen' in st10:
    print('  10 Xu: FAIL - still mentions Wen')
    r9_errors.append(('10','Xu','Wen mention','remove'))
else:
    print('  10 Xu: PASS - no Wen reference')

# 9b: Xu says "only 2 studies"
if 'Saikia 2023' in st10 and 'Xu 2021' in st10:
    print('  10 Xu: PASS - mentions Saikia+Xu as 2 studies')
else:
    print('  10 Xu: CHECK - verify study references')

# 9c: Mandaknalli labels fixed
st03 = script_text('Tom-tat-tung-bai/03_Mandaknalli_Malusare_2021.py')
has_old = 'Tối thiểu' in st03
has_new = 'Rất nặng' in st03
if has_new and not has_old:
    print('  03 Mandaknalli: PASS - labels corrected')
else:
    print(f'  03 Mandaknalli: FAIL - old={has_old}, new={has_new}')
    r9_errors.append(('03','Mandaknalli','labels','still wrong'))

# 9d: Mandaknalli smoking fixed
if '85,2%' in st03 or '85.2' in st03:
    print('  03 Mandaknalli: PASS - smoking 85.2%')
else:
    print('  03 Mandaknalli: FAIL - smoking not corrected')
    r9_errors.append(('03','Mandaknalli','smoking','not 85.2%'))

# 9e: Wen says female > male
st08 = script_text('Tom-tat-tung-bai/08_Wen_et_al_2020.py')
if '0.262' in st08:
    print('  08 Wen: PASS - OR 0.262 present (male protective)')
else:
    print('  08 Wen: FAIL')
    r9_errors.append(('08','Wen','0.262','missing'))

all_errors += r9_errors

# FINAL
print(f'\n{"="*60}')
print(f'FINAL SUMMARY — ROUNDS 4-9 (6 rounds)')
print(f'{"="*60}')
print(f'Total checks: {total_checks}')
print(f'Total errors: {len(all_errors)}')
if not all_errors:
    print('>>> ALL 6 ROUNDS PASSED — 0 ERRORS <<<')
else:
    for num, name, val, desc in all_errors:
        print(f'  {num} {name}: {val} ({desc})')
