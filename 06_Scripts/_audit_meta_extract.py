"""Extract summary text and first-pages PDF text for fact-check comparison."""
import os, sys, json, glob, io
import docx
import fitz
# Force UTF-8 stdout
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

ROOT = r'c:/Users/OS/OneDrive/read_books/Lo-au'
TOM_TAT = os.path.join(ROOT, 'Tom-tat-tung-bai')
PAPERS = os.path.join(ROOT, '02_Papers-goc')

with open(os.path.join(PAPERS, 'canonical_index.json'), 'r', encoding='utf-8') as f:
    index = json.load(f)

def read_docx(path):
    try:
        d = docx.Document(path)
        return '\n'.join(p.text for p in d.paragraphs if p.text.strip())
    except Exception as e:
        return f'[ERR {e}]'

def find_summary(qid):
    # Prefer FIXED_27052026
    pats = [
        os.path.join(TOM_TAT, f'{qid}_*_FIXED_27052026.docx'),
        os.path.join(TOM_TAT, f'{qid}_*.docx'),
    ]
    for pat in pats:
        files = [f for f in glob.glob(pat) if 'backup' not in f.lower() and 'before_fix' not in f.lower()]
        if files:
            return files[0]
    return None

def find_pdf(qid, info):
    pdf_name = info.get('pdf')
    folder = info.get('pdf_folder')
    if not pdf_name:
        return None
    candidates = []
    if folder:
        candidates.append(os.path.join(PAPERS, folder, pdf_name))
    # Also try direct
    candidates.append(os.path.join(PAPERS, pdf_name))
    # Glob fallback
    for c in candidates:
        if os.path.exists(c):
            return c
    # Recursive search
    for root, _, files in os.walk(PAPERS):
        if pdf_name in files:
            return os.path.join(root, pdf_name)
    return None

def read_pdf_first_pages(path, n=4):
    try:
        doc = fitz.open(path)
        out = []
        for i in range(min(n, doc.page_count)):
            out.append(doc[i].get_text())
        doc.close()
        return '\n--PAGE--\n'.join(out)
    except Exception as e:
        return f'[ERR {e}]'

target_ids = ['QT012','QT013','QT018','QT020','QT023','QT029','QT030','QT031','QT034','QT035','QT039','QT046','QT048','QT049','QT051','QT052','QT053','QT059','QT060','QT064','QT066','QT070','QT072']

qid = sys.argv[1]
info = index.get(qid, {})
summary_path = find_summary(qid)
pdf_path = find_pdf(qid, info)
print('===SUMMARY_PATH===', summary_path)
print('===PDF_PATH===', pdf_path)
print('===DESCRIPTOR===', info.get('descriptor'))
print('---SUMMARY TEXT---')
if summary_path:
    print(read_docx(summary_path))
print('---PDF TEXT (first 4 pages)---')
if pdf_path:
    print(read_pdf_first_pages(pdf_path, 4))
else:
    print('[NO PDF AVAILABLE]')
