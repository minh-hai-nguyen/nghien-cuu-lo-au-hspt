"""Print compact comparison views per QID: summary text + PDF abstract+findings."""
import os, sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

OUT = r'c:/Users/OS/OneDrive/read_books/Lo-au/_audit_tmp'

qid = sys.argv[1]
with open(os.path.join(OUT, f'{qid}.txt'), 'r', encoding='utf-8') as f:
    txt = f.read()

print('='*80)
print(qid)
print('='*80)
# Show summary in full + PDF first page or so for comparison
parts = txt.split('---PDF TEXT (first 4 pages)---')
summary_part = parts[0]
pdf_part = parts[1] if len(parts) > 1 else ''
print(summary_part)
print('---PDF EXTRACT---')
# Print first 5000 chars of PDF (abstract + first results)
print(pdf_part[:6000])
