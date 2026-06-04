# -*- coding: utf-8 -*-
"""Kiem tra cac yeu to khac co the bi vo format.
28/05/2026."""
import os, sys, io, zipfile
os.environ['PYTHONIOENCODING'] = 'utf-8'
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
BACKUP = os.path.join(ROOT, 'Luận án TS', '01_LuanAn_v3_1_FixCoping_28052026_BEFORE_TOC.docx')
CURRENT = os.path.join(ROOT, 'Luận án TS', '01_LuanAn_v3_1_FixCoping_28052026.docx')

print("="*75)
print("KIEM TRA CAC YEU TO KHAC CO THE BI VO FORMAT")
print("="*75)

bk = Document(BACKUP)
nw = Document(CURRENT)

def count_xml_elements(doc, tag):
    """Count occurrences of a tag in entire document XML."""
    return sum(1 for _ in doc.element.body.iter(tag))


# Count via XML elements
checks = [
    ('Page breaks (w:br type=page)', None),  # special: count w:br with type=page
    ('Section breaks (w:sectPr in body)', qn('w:sectPr')),
    ('Hyperlinks (w:hyperlink)', qn('w:hyperlink')),
    ('Bookmarks (w:bookmarkStart)', qn('w:bookmarkStart')),
    ('Comments references (w:commentReference)', qn('w:commentReference')),
    ('Footnotes (w:footnoteReference)', qn('w:footnoteReference')),
    ('Endnotes (w:endnoteReference)', qn('w:endnoteReference')),
    ('Drawings (w:drawing)', qn('w:drawing')),
    ('Pictures (w:pict)', qn('w:pict')),
    ('Objects (w:object)', qn('w:object')),
    ('Field codes (w:fldChar begin)', None),  # custom
    ('Field instructions (w:instrText)', qn('w:instrText')),
    ('Numbering refs (w:numPr)', qn('w:numPr')),
    ('List items with numId', None),  # custom
    ('Tabs (w:tab)', qn('w:tab')),
    ('Line breaks (w:br no type)', None),  # custom
    ('Math equations (m:oMath)', '{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath'),
    ('Inline shapes (w:txbxContent textboxes)', qn('w:txbxContent')),
    ('Track changes (w:ins insertions)', qn('w:ins')),
    ('Track changes (w:del deletions)', qn('w:del')),
    ('SDT content controls (w:sdt)', qn('w:sdt')),
]

print("\n--- XML element counts ---")
print(f"{'Element':<55}{'Backup':>10}{'Current':>10}")
print("-"*75)
results = []
for name, tag in checks:
    if tag is None:
        # Custom counts
        if 'Page breaks' in name:
            bk_count = sum(1 for e in bk.element.body.iter(qn('w:br')) if e.get(qn('w:type')) == 'page')
            nw_count = sum(1 for e in nw.element.body.iter(qn('w:br')) if e.get(qn('w:type')) == 'page')
        elif 'Field codes (w:fldChar begin)' in name:
            bk_count = sum(1 for e in bk.element.body.iter(qn('w:fldChar')) if e.get(qn('w:fldCharType')) == 'begin')
            nw_count = sum(1 for e in nw.element.body.iter(qn('w:fldChar')) if e.get(qn('w:fldCharType')) == 'begin')
        elif 'List items with numId' in name:
            bk_count = sum(1 for e in bk.element.body.iter(qn('w:numId')))
            nw_count = sum(1 for e in nw.element.body.iter(qn('w:numId')))
        elif 'Line breaks' in name:
            bk_count = sum(1 for e in bk.element.body.iter(qn('w:br')) if e.get(qn('w:type')) is None)
            nw_count = sum(1 for e in nw.element.body.iter(qn('w:br')) if e.get(qn('w:type')) is None)
        else:
            bk_count = 0
            nw_count = 0
    else:
        bk_count = sum(1 for _ in bk.element.body.iter(tag))
        nw_count = sum(1 for _ in nw.element.body.iter(tag))
    diff = nw_count - bk_count
    mark = "✓" if (diff == 0 or (name.startswith('Field codes') and diff == 1) or (name.startswith('Field instructions') and diff == 1)) else "?"
    diff_str = f"({diff:+d})" if diff != 0 else ""
    print(f"  {mark} {name:<53}{bk_count:>10}{nw_count:>10} {diff_str}")
    results.append((name, bk_count, nw_count, diff))


# Check archive contents (parts inside .docx zip)
print("\n--- DOCX archive parts (zip contents) ---")
with zipfile.ZipFile(BACKUP) as zb, zipfile.ZipFile(CURRENT) as zn:
    bk_names = set(zb.namelist())
    nw_names = set(zn.namelist())
    only_bk = bk_names - nw_names
    only_nw = nw_names - bk_names
    print(f"  Backup parts: {len(bk_names)}")
    print(f"  Current parts: {len(nw_names)}")
    if only_bk:
        print(f"  REMOVED from current: {sorted(only_bk)}")
    if only_nw:
        print(f"  ADDED in current: {sorted(only_nw)}")
    if not only_bk and not only_nw:
        print("  ✓ Same set of parts (no media/files added or removed)")


# Check that document has expected parts (images, embeddings, etc.)
print("\n--- Important parts present ---")
with zipfile.ZipFile(CURRENT) as zn:
    for name in sorted(zn.namelist()):
        if 'word/media/' in name or 'word/embeddings/' in name or 'word/charts/' in name:
            print(f"  {name} ({zn.getinfo(name).file_size} bytes)")


# Check styles.xml hasn't changed
print("\n--- Style definitions ---")
with zipfile.ZipFile(BACKUP) as zb, zipfile.ZipFile(CURRENT) as zn:
    try:
        bk_styles = zb.read('word/styles.xml')
        nw_styles = zn.read('word/styles.xml')
        same = bk_styles == nw_styles
        print(f"  word/styles.xml: {'✓ identical' if same else '? different'}")
        print(f"  Backup styles size: {len(bk_styles)} bytes")
        print(f"  Current styles size: {len(nw_styles)} bytes")
    except KeyError as e:
        print(f"  Error: {e}")


# Check numbering.xml (list definitions)
print("\n--- Numbering / list definitions ---")
with zipfile.ZipFile(BACKUP) as zb, zipfile.ZipFile(CURRENT) as zn:
    bk_nums = 'word/numbering.xml' in zb.namelist()
    nw_nums = 'word/numbering.xml' in zn.namelist()
    print(f"  word/numbering.xml present: backup={bk_nums} current={nw_nums}")
    if bk_nums and nw_nums:
        bk_num_xml = zb.read('word/numbering.xml')
        nw_num_xml = zn.read('word/numbering.xml')
        print(f"  Numbering xml identical: {bk_num_xml == nw_num_xml}")


# Check footnotes/endnotes parts
print("\n--- Footnotes / endnotes ---")
with zipfile.ZipFile(BACKUP) as zb, zipfile.ZipFile(CURRENT) as zn:
    for part in ['word/footnotes.xml', 'word/endnotes.xml']:
        bk_has = part in zb.namelist()
        nw_has = part in zn.namelist()
        if bk_has or nw_has:
            print(f"  {part}: backup={bk_has} current={nw_has}")
            if bk_has and nw_has:
                same = zb.read(part) == zn.read(part)
                print(f"    Identical: {same}")


# Check media (images)
print("\n--- Media (images) ---")
with zipfile.ZipFile(BACKUP) as zb, zipfile.ZipFile(CURRENT) as zn:
    bk_media = sorted([n for n in zb.namelist() if 'word/media/' in n])
    nw_media = sorted([n for n in zn.namelist() if 'word/media/' in n])
    print(f"  backup media files: {len(bk_media)}")
    print(f"  current media files: {len(nw_media)}")
    same_count = 0
    diff_count = 0
    for m in bk_media:
        if m in nw_media:
            if zb.read(m) == zn.read(m):
                same_count += 1
            else:
                diff_count += 1
    print(f"  Identical media: {same_count}, Modified: {diff_count}")


print("\n" + "="*75)
print("TONG KET")
print("="*75)
non_zero = [(n,b,c,d) for n,b,c,d in results if d != 0]
expected_changes = ['Field codes', 'Field instructions']
unexpected = [(n,b,c,d) for n,b,c,d in non_zero if not any(e in n for e in expected_changes)]
if not unexpected:
    print("  ✓ KHONG CO YEU TO NAO BI VO FORMAT")
    print("    Chi thay doi: +1 fldChar (begin) va +1 instrText cho TOC field — DUNG NHU MONG DOI")
else:
    print(f"  ? Co {len(unexpected)} thay doi can xem xet:")
    for n,b,c,d in unexpected:
        print(f"    {n}: {b} -> {c} (diff {d:+d})")
