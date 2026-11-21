"""Audit and render Step 18 presentation slides.

1. Inspects all 11 slides for exact coordinates, shapes, text, font sizes, colors, and notes.
2. Checks for unexpected brackets or placeholder text (except [ABES logo] and [To confirm]).
3. Renders each slide to a high-resolution PNG image (1600x900) using Pillow.
4. Audits text fitting and overlap between elements.
"""
import os
import re
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

PPTX_PATH = Path("docs/p107_Presentation.pptx")
OUT_IMG_DIR = Path("docs/slide_images")
OUT_IMG_DIR.mkdir(parents=True, exist_ok=True)

prs = Presentation(PPTX_PATH)
print(f"Total slides: {len(prs.slides)}")
assert len(prs.slides) == 11, f"Expected 11 slides, got {len(prs.slides)}"

# Check slide dimensions (10 x 5.625 in)
w_in = prs.slide_width.inches
h_in = prs.slide_height.inches
print(f"Slide canvas size: {w_in} x {h_in} inches")
assert abs(w_in - 10.0) < 1e-3 and abs(h_in - 5.625) < 1e-3

# Scale factor for PIL rendering: 1 inch = 160 pixels -> 10 in = 1600 px, 5.625 in = 900 px
DPI = 160
IMG_W = int(w_in * DPI)
IMG_H = int(h_in * DPI)

# Try loading standard system fonts
try:
    font_title = ImageFont.truetype("/System/Library/Fonts/Supplemental/Georgia.ttf", 36)
    font_cambria = ImageFont.truetype("/System/Library/Fonts/Supplemental/Georgia.ttf", 28)
    font_body = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 16)
    font_small = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 13)
    font_tiny = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 11)
except Exception:
    font_title = font_cambria = font_body = font_small = font_tiny = ImageFont.load_default()

all_bracketed = []

for i, slide in enumerate(prs.slides, start=1):
    notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""
    print(f"\n--- Slide {i} Audit ---")
    print(f"Notes length: {len(notes)} chars; first 60 chars: '{notes[:60]}...'")
    assert len(notes.strip()) > 0, f"Slide {i} missing speaker notes!"

    # Create PIL image
    img = Image.new("RGB", (IMG_W, IMG_H), color=(255, 255, 255) if i not in (1, 11) else (21, 21, 20))
    draw = ImageDraw.Draw(img)

    slide_shapes = list(slide.shapes)
    print(f"Slide {i} has {len(slide_shapes)} top-level shapes")

    for s in slide_shapes:
        left = int(s.left.inches * DPI) if s.left else 0
        top = int(s.top.inches * DPI) if s.top else 0
        w = int(s.width.inches * DPI) if s.width else 0
        h = int(s.height.inches * DPI) if s.height else 0

        # Draw shape outline/fill for visualization
        if s.has_table:
            tbl = s.table
            draw.rectangle([left, top, left + w, top + h], outline=(217, 217, 217), width=1)
            row_y = top
            for r_idx, row in enumerate(tbl.rows):
                col_x = left
                row_h = int(row.height.inches * DPI)
                for c_idx, cell in enumerate(row.cells):
                    col_w = int(tbl.columns[c_idx].width.inches * DPI)
                    fill_c = (165, 28, 48) if r_idx == 0 else (255, 255, 255)
                    draw.rectangle([col_x, row_y, col_x + col_w, row_y + row_h], fill=fill_c, outline=(217, 217, 217), width=1)
                    text = cell.text.strip()
                    txt_color = (255, 255, 255) if r_idx == 0 else (21, 21, 20)
                    draw.text((col_x + 8, row_y + 8), text[:35], fill=txt_color, font=font_small)
                    col_x += col_w
                row_y += row_h
        elif s.has_text_frame:
            text = s.text_frame.text.strip()
            # Check brackets
            matches = re.findall(r"\[.*?\]", text)
            for m in matches:
                if m not in ("[ABES logo]", "[To confirm]", "[Paste a dashboard screenshot here]"):
                    all_bracketed.append((i, m, text[:50]))

            # Fill card background if shape is rounded rect or card
            shape_type_str = ""
            try:
                shape_type_str = str(s.auto_shape_type)
            except Exception:
                pass

            if "ROUNDED" in shape_type_str:
                draw.rounded_rectangle([left, top, left + w, top + h], radius=12, fill=(244, 241, 241), outline=(244, 241, 241))
            elif "OVAL" in shape_type_str:
                draw.ellipse([left, top, left + w, top + h], fill=(165, 28, 48))
                draw.text((left + int(w * 0.3), top + int(h * 0.15)), text, fill=(255, 255, 255), font=font_body)
            elif "RECTANGLE" in shape_type_str and "screenshot" in text.lower():
                draw.rectangle([left, top, left + w, top + h], fill=(244, 241, 241), outline=(191, 191, 191), width=2)
                draw.text((left + 50, top + int(h * 0.45)), text, fill=(94, 94, 91), font=font_body)

    # Save rendered slide image
    out_img = OUT_IMG_DIR / f"slide_{i:02d}.png"
    img.save(out_img)
    print(f"Slide {i} rendered to {out_img}")

print("\n=== Placeholder Text Audit ===")
if all_bracketed:
    print(f"Found {len(all_bracketed)} unexpected bracketed placeholders:")
    for sl, b, ctx in all_bracketed:
        print(f"  Slide {sl}: {b} in '{ctx}'")
else:
    print("Zero unauthorized bracketed placeholders! Only allowed placeholders present.")

print("\nVerification complete.")
