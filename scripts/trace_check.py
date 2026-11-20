"""Final traceability cross-check for Step 17 document."""
from docx import Document
import re

doc = Document("docs/HKD-2026_Project_Document.docx")

# Collect all text
all_text = ""
for p in doc.paragraphs:
    all_text += p.text + "\n"
for t in doc.tables:
    for row in t.rows:
        for cell in row.cells:
            all_text += cell.text + "\n"

# Check ID families
id_families = {
    "HMW": r"HMW-\d+",
    "SH": r"SH-\d+",
    "BR": r"BR-\d+",
    "FR": r"FR-\d+",
    "NFR": r"NFR-\d+",
    "US": r"US-\d+",
    "UI": r"UI-\d+",
    "TC-F": r"TC-F-\d+",
    "TC-T": r"TC-T-\d+",
}

for name, pattern in id_families.items():
    ids = sorted(set(re.findall(pattern, all_text)))
    print(f"{name}: {len(ids)} unique IDs")
    print(f"  {ids}")

# Check Must BRs have US, TC, UI in RTM
print("\n=== Must BR trace check (from RTM table) ===")
must_brs = ["BR-01","BR-02","BR-03","BR-04","BR-05","BR-06","BR-07","BR-09","BR-13"]
# Find the traceability table (table 4 is the 9-column one)
rtm_table = doc.tables[4]
for row in rtm_table.rows[1:]:  # skip header
    cells = [c.text.strip() for c in row.cells]
    br_id = cells[1]
    priority = cells[5]
    us = cells[6]
    tc = cells[7]
    ui = cells[8]
    if priority == "Must":
        has_us = bool(re.search(r"US-\d+", us))
        has_tc = bool(re.search(r"TC-[FT]-\d+", tc))
        has_ui = bool(re.search(r"UI-\d+", ui))
        status = "OK" if (has_us and has_tc and has_ui) else "MISSING"
        print(f"  {br_id} ({priority}): US={us}, TC={tc}, UI={ui} -> {status}")

print("\nDone.")
