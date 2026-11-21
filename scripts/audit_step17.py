"""Audit the Step 17 Project Document for completeness and formatting."""
import docx
from docx import Document
from docx.shared import Pt
from lxml import etree
import re

doc = Document("docs/p107_Project_Document.docx")

WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ns = {"w": WML}

# 1. Paragraph styles
print("=== PARAGRAPH STYLES ===")
style_counts = {}
for p in doc.paragraphs:
    s = p.style.name
    style_counts[s] = style_counts.get(s, 0) + 1
for s, c in sorted(style_counts.items()):
    print(f"  {s}: {c}")

# 2. Check for Heading 1 / Heading 2 usage
print("\n=== HEADINGS (Word styles) ===")
for i, p in enumerate(doc.paragraphs):
    if "Heading" in p.style.name or "heading" in p.style.name:
        print(f"  [{i}] style='{p.style.name}' text='{p.text[:70]}'")

# 3. Heading-like (bold + large font) but NOT heading style
print("\n=== HEADING-LIKE BUT NOT HEADING STYLE ===")
for i, p in enumerate(doc.paragraphs):
    if p.text.strip() and len(p.text) < 100 and "Heading" not in p.style.name:
        for run in p.runs:
            if run.bold and run.font.size and run.font.size >= Pt(12):
                print(f"  [{i}] style='{p.style.name}' size={run.font.size} text='{p.text[:70]}'")
                break

# 4. Headers & footers
print("\n=== HEADERS & FOOTERS ===")
for i, sec in enumerate(doc.sections):
    h = sec.header
    f = sec.footer
    print(f"  Section {i}: header_linked={h.is_linked_to_previous}, footer_linked={f.is_linked_to_previous}")
    for p in h.paragraphs:
        if p.text.strip():
            print(f"    Header: '{p.text}'")
    for p in f.paragraphs:
        if p.text.strip():
            print(f"    Footer: '{p.text}'")
    # Check for PAGE field in footer
    for p in f.paragraphs:
        flds = p._element.findall(".//w:fldChar", ns)
        instrs = p._element.findall(".//w:instrText", ns)
        if flds:
            print(f"    Footer has {len(flds)} field chars")
        for inst in instrs:
            print(f"    Footer instrText: '{inst.text}'")

# 5. Tables audit
print("\n=== TABLES ===")
for i, t in enumerate(doc.tables):
    rows = len(t.rows)
    cols = len(t.columns)
    first = t.rows[0].cells[0].text[:50] if rows > 0 else ""
    # Check table width
    tblW = t._tbl.find(".//w:tblW", ns)
    w_val = tblW.get(etree.QName(WML, "w")) if tblW is not None else "?"
    # Check layout
    tblLayout = t._tbl.find(".//w:tblLayout", ns)
    layout = tblLayout.get(etree.QName(WML, "type")) if tblLayout is not None else "?"
    # Check header row repeat
    first_tr = t._tbl.findall("w:tr", ns)
    has_header = False
    if first_tr:
        trPr = first_tr[0].find("w:trPr", ns)
        if trPr is not None:
            has_header = trPr.find("w:tblHeader", ns) is not None
    print(f"  Table {i}: {rows}r x {cols}c | w={w_val} | layout={layout} | headerRepeat={has_header} | '{first}'")

# 6. Check updateFields
print("\n=== SETTINGS ===")
uf_elements = doc.settings.element.findall(".//w:updateFields", ns)
print(f"  updateFields count: {len(uf_elements)}")
for u in uf_elements:
    val = u.get(etree.QName(WML, "val"))
    print(f"  updateFields val={val}")

# 7. Bracketed placeholders
print("\n=== BRACKETED TEXT ===")
all_bracketed = []
for p in doc.paragraphs:
    matches = re.findall(r"\[.*?\]", p.text)
    for m in matches:
        all_bracketed.append(("para", m, p.text[:60]))
for t in doc.tables:
    for row in t.rows:
        for cell in row.cells:
            matches = re.findall(r"\[.*?\]", cell.text)
            for m in matches:
                all_bracketed.append(("cell", m, cell.text[:60]))
print(f"  Total: {len(all_bracketed)}")
for src, m, ctx in all_bracketed:
    print(f"  [{src}] {m} in '{ctx}'")

# 8. Section content check: look for expected section titles
print("\n=== EXPECTED SECTIONS CHECK ===")
expected = [
    "1 Description", "2 HMW", "3 Stakeholders", "4 Traceability",
    "5 User Stories", "6 BRD", "7 TRD", "8 Database Design",
    "9 Test Cases", "10 UI Framework", "Appendix A", "Appendix B"
]
all_text = "\n".join(p.text for p in doc.paragraphs)
for sec in expected:
    found = sec in all_text
    print(f"  {'✓' if found else '✗'} {sec}")

# 9. File size
import os
size = os.path.getsize("docs/p107_Project_Document.docx")
print(f"\n=== FILE SIZE: {size:,} bytes ===")

# 10. Margin check
print("\n=== PAGE MARGINS ===")
for i, sec in enumerate(doc.sections):
    print(f"  Section {i}: top={sec.top_margin} bottom={sec.bottom_margin} left={sec.left_margin} right={sec.right_margin}")
    print(f"    header_dist={sec.header_distance} footer_dist={sec.footer_distance}")
    print(f"    page_width={sec.page_width} page_height={sec.page_height}")
