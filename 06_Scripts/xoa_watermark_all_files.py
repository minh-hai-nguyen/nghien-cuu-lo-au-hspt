# -*- coding: utf-8 -*-
"""Strip watermark (w:pict elements) tu tat ca file docx trong Luan an TS/.
Theo yeu cau cua user: 'luon xoa het water mark di'."""
import os, sys, io
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
FOLDER = os.path.join(ROOT, 'Luận án TS')

def strip_watermarks(docx_path, save_path=None):
    """Loai bo w:pict elements khoi body cua docx.
    Cung loai bo cac element trong header/footer.
    Return (n_removed_body, n_removed_header_footer)."""
    if save_path is None:
        save_path = docx_path
    d = Document(docx_path)
    body = d.element.body

    # Body
    picts_body = body.findall('.//' + qn('w:pict'))
    n_body = len(picts_body)
    for p in picts_body:
        # Find the containing w:r run, remove the entire run if pict is the only content
        parent = p.getparent()
        # Remove the pict element itself
        parent.remove(p)

    # Header/Footer
    n_hf = 0
    for sec in d.sections:
        for hf in [sec.header, sec.footer]:
            for p_para in hf.paragraphs:
                picts = p_para._p.findall('.//' + qn('w:pict'))
                for pi in picts:
                    parent = pi.getparent()
                    parent.remove(pi)
                    n_hf += 1

    d.save(save_path)
    return n_body, n_hf

# Process tat ca file docx trong folder Luan an TS
files_to_process = []
for fname in os.listdir(FOLDER):
    if fname.endswith('.docx'):
        fpath = os.path.join(FOLDER, fname)
        files_to_process.append(fpath)

print(f"Tim duoc {len(files_to_process)} file docx trong {FOLDER}")
print()

# Truoc khi sua, BACKUP file goc va v2
import shutil
backup_files = [
    '01_Rối loạn lo âu_ 26-05.docx',
    '01_RoiLoanLoAu_v2_RedEdits_26052026.docx',
]
for fname in backup_files:
    fpath = os.path.join(FOLDER, fname)
    if os.path.exists(fpath):
        backup = fpath.replace('.docx', '_backup_original.docx')
        if not os.path.exists(backup):
            shutil.copy2(fpath, backup)
            print(f"Backup: {os.path.basename(backup)}")

# Strip watermark cho TAT CA file
results = []
for fpath in files_to_process:
    if '_backup_' in fpath:
        continue  # Skip backup files
    fname = os.path.basename(fpath)
    try:
        n_b, n_hf = strip_watermarks(fpath)
        results.append((fname, n_b, n_hf))
        marker = '✓ DA XOA' if (n_b + n_hf) > 0 else '✓ Da clean'
        print(f"  {marker}: {fname} (body={n_b}, header/footer={n_hf})")
    except Exception as e:
        print(f"  ! ERROR: {fname}: {e}")

print()
print("=== SUMMARY ===")
total_b = sum(r[1] for r in results)
total_hf = sum(r[2] for r in results)
print(f"Tong watermark da xoa: body={total_b}, header/footer={total_hf}")
print(f"Tong file da xu ly: {len(results)}")
