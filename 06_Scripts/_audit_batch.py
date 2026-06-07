"""Batch extract: writes ONE file per QT ID — summary + first 4 PDF pages."""
import os, sys, json, glob, io
import docx
import fitz
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

ROOT = r'c:/Users/OS/OneDrive/read_books/Lo-au'
TOM_TAT = os.path.join(ROOT, 'Tom-tat-tung-bai')
PAPERS = os.path.join(ROOT, '02_Papers-goc')
OUT = os.path.join(ROOT, '_audit_tmp')
os.makedirs(OUT, exist_ok=True)

with open(os.path.join(PAPERS, 'canonical_index.json'), 'r', encoding='utf-8') as f:
    index = json.load(f)

def read_docx(path):
    try:
        d = docx.Document(path)
        return '\n'.join(p.text for p in d.paragraphs if p.text.strip())
    except Exception as e:
        return f'[ERR docx {e}]'

def find_summary(qid):
    pats = [
        os.path.join(TOM_TAT, f'{qid}_*_FIXED_27052026.docx'),
        os.path.join(TOM_TAT, f'{qid}_*.docx'),
    ]
    for pat in pats:
        files = [f for f in glob.glob(pat)
                 if 'backup' not in f.lower() and 'before_fix' not in f.lower()]
        if files:
            return files[0]
    return None

def find_pdf(qid, info):
    pdf_name = info.get('pdf')
    if not pdf_name:
        return None
    folder = info.get('pdf_folder')
    if folder:
        c = os.path.join(PAPERS, folder, pdf_name)
        if os.path.exists(c):
            return c
    for root, _, files in os.walk(PAPERS):
        if pdf_name in files:
            return os.path.join(root, pdf_name)
    return None

def read_pdf_pages(path, n=4):
    try:
        doc = fitz.open(path)
        out = []
        for i in range(min(n, doc.page_count)):
            out.append(doc[i].get_text())
        doc.close()
        return '\n--PAGE--\n'.join(out)
    except Exception as e:
        return f'[ERR pdf {e}]'

ids = ['QT012','QT013','QT018','QT020','QT023','QT029','QT030','QT031','QT034',
       'QT035','QT039','QT046','QT048','QT049','QT051','QT052','QT053','QT059',
       'QT060','QT064','QT066','QT070','QT072']

manifest = []
for qid in ids:
    info = index.get(qid, {})
    sp = find_summary(qid)
    pp = find_pdf(qid, info)
    summary = read_docx(sp) if sp else '[NO SUMMARY]'
    pdf = read_pdf_pages(pp, 4) if pp else '[NO PDF]'
    out_path = os.path.join(OUT, f'{qid}.txt')
    with open(out_path, 'w', encoding='utf-8', errors='replace') as f:
        f.write(f'===QID=== {qid}\n')
        f.write(f'===SUMMARY_PATH=== {sp}\n')
        f.write(f'===PDF_PATH=== {pp}\n')
        f.write(f'===DESCRIPTOR=== {info.get("descriptor")}\n')
        f.write('---SUMMARY TEXT---\n')
        f.write(summary)
        f.write('\n---PDF TEXT (first 4 pages)---\n')
        f.write(pdf)
    manifest.append((qid, sp is not None, pp is not None))
    print(f'{qid}: summary={"Y" if sp else "N"} pdf={"Y" if pp else "N"}')

print('DONE')
