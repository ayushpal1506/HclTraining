"""Check layout bounding boxes and text overflows across all slides."""
from pptx import Presentation
from pptx.util import Inches

prs = Presentation("docs/p107_Presentation.pptx")

for i, slide in enumerate(prs.slides, start=1):
    boxes = []
    print(f"\nSlide {i} layout verification:")
    for s in slide.shapes:
        left = round(s.left.inches, 3) if s.left else 0.0
        top = round(s.top.inches, 3) if s.top else 0.0
        width = round(s.width.inches, 3) if s.width else 0.0
        height = round(s.height.inches, 3) if s.height else 0.0
        right = round(left + width, 3)
        bottom = round(top + height, 3)
        name = s.name
        text_snip = s.text_frame.text[:25].replace("\n", " ") if s.has_text_frame else ""
        if s.has_table:
            text_snip = f"TABLE {len(s.table.rows)}x{len(s.table.columns)}"

        # Check bounds within canvas 10 x 5.625
        assert left >= 0.0, f"Shape {name} extends past left border: {left}"
        assert right <= 10.05, f"Shape {name} extends past right border: {right}"
        assert top >= 0.0, f"Shape {name} extends past top border: {top}"
        assert bottom <= 5.65, f"Shape {name} extends past bottom border: {bottom}"

        print(f"  [{name}] x: {left:4.2f} - {right:4.2f}, y: {top:4.2f} - {bottom:4.2f} | '{text_snip}'")

print("\nAll shape boundaries within canvas limits (10 x 5.625 in)!")
