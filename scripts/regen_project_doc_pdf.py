"""Re-generate only the Project Document PDF with the ABES logo embedded."""
import io
import shutil
from pathlib import Path

from docx import Document
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.table import Table as DocxTable
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
)
from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_LEFT, TA_CENTER

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
SUB_DIR = ROOT / "submission_project_folder"

C_RED    = colors.HexColor("#A51C30")
C_DARK   = colors.HexColor("#151514")
C_GREY   = colors.HexColor("#7F7F7F")
C_BORDER = colors.HexColor("#BFBFBF")
C_LIGHT  = colors.HexColor("#F5F5F5")
C_WHITE  = colors.white


# --------------------------------------------------------------------------- styles
def get_styles():
    ss = getSampleStyleSheet()
    base_font = "Helvetica"

    styles = {}
    styles['DocTitle'] = ParagraphStyle('DocTitle', parent=ss['Normal'],
        fontName=base_font+'-Bold', fontSize=18, textColor=C_RED,
        spaceAfter=10, spaceBefore=6, alignment=TA_LEFT)
    styles['DocSubTitle'] = ParagraphStyle('DocSubTitle', parent=ss['Normal'],
        fontName=base_font, fontSize=11, textColor=C_GREY,
        spaceAfter=8, spaceBefore=2, alignment=TA_LEFT)
    styles['Heading1_Custom'] = ParagraphStyle('Heading1_Custom', parent=ss['Normal'],
        fontName=base_font+'-Bold', fontSize=13, textColor=C_RED,
        spaceBefore=14, spaceAfter=5)
    styles['Heading2_Custom'] = ParagraphStyle('Heading2_Custom', parent=ss['Normal'],
        fontName=base_font+'-Bold', fontSize=11, textColor=C_DARK,
        spaceBefore=8, spaceAfter=4)
    styles['Body_Custom'] = ParagraphStyle('Body_Custom', parent=ss['Normal'],
        fontName=base_font, fontSize=9.5, textColor=C_DARK,
        spaceAfter=3, leading=14)
    styles['TableHead'] = ParagraphStyle('TableHead', parent=ss['Normal'],
        fontName=base_font+'-Bold', fontSize=8.5, textColor=C_WHITE, spaceAfter=0)
    styles['TableCell'] = ParagraphStyle('TableCell', parent=ss['Normal'],
        fontName=base_font, fontSize=8.5, textColor=C_DARK, spaceAfter=0, leading=12)
    return styles


class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):
        self.setFont("Helvetica", 8)
        self.setFillColor(C_GREY)
        self.drawRightString(8.5*inch - 54, 36, f"Ayush Pal | 2500320140037     Page {self._pageNumber} of {page_count}")
        self.drawString(54, 36, "ABES Engineering College | Project Document | p107")


def _extract_inline_images(p_element, doc):
    """Return list of image bytes for all inline images in a paragraph element."""
    ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    blips = p_element.findall(f'.//{{{ns}}}blip')
    images = []
    for blip in blips:
        rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        if rId and rId in doc.part.rels:
            img_part = doc.part.rels[rId].target_part
            images.append(img_part.blob)
    return images


def convert_docx_to_pdf(docx_path: Path, pdf_path: Path):
    doc = Document(docx_path)
    styles = get_styles()
    story = []
    usable_width = 8.5 * inch - 108

    for element in doc.element.body:
        if isinstance(element, CT_P):
            p = DocxParagraph(element, doc)
            text = p.text.strip()

            # Detect inline images in this paragraph
            inline_imgs = _extract_inline_images(element, doc)
            if inline_imgs:
                for img_bytes in inline_imgs:
                    try:
                        rl_img = Image(io.BytesIO(img_bytes), width=1.1*inch, height=1.1*inch)
                        rl_img.hAlign = 'LEFT'
                        story.append(rl_img)
                        story.append(Spacer(1, 6))
                    except Exception as e:
                        print(f"  [warn] Could not embed image: {e}")
                if not text:
                    continue  # image-only paragraph, done

            if not text:
                story.append(Spacer(1, 4))
                continue

            style_name = p.style.name if p.style else ""
            clean = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

            if "Heading 1" in style_name or "Heading1" in style_name:
                story.append(Paragraph(clean, styles['Heading1_Custom']))
            elif "Heading 2" in style_name or "Heading2" in style_name:
                story.append(Paragraph(clean, styles['Heading2_Custom']))
            elif text.startswith("# "):
                story.append(Paragraph(clean[2:], styles['Heading1_Custom']))
            elif text.startswith("## "):
                story.append(Paragraph(clean[3:], styles['Heading2_Custom']))
            elif "Title" in style_name:
                story.append(Paragraph(clean, styles['DocTitle']))
            elif "Subtitle" in style_name:
                story.append(Paragraph(clean, styles['DocSubTitle']))
            else:
                story.append(Paragraph(clean, styles['Body_Custom']))

        elif isinstance(element, CT_Tbl):
            tbl = DocxTable(element, doc)
            rows_data = []
            for row in tbl.rows:
                row_cells = [cell.text.strip() for cell in row.cells]
                if any(c for c in row_cells):
                    rows_data.append(row_cells)

            if rows_data:
                num_cols = max(len(r) for r in rows_data)
                col_width = usable_width / num_cols

                table_data = []
                for r_idx, row in enumerate(rows_data):
                    while len(row) < num_cols:
                        row.append("")
                    row_p = []
                    for val in row:
                        clean = val.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                        if len(clean) > 500:
                            clean = clean[:497] + "..."
                        if r_idx == 0:
                            row_p.append(Paragraph(clean, styles['TableHead']))
                        else:
                            row_p.append(Paragraph(clean, styles['TableCell']))
                    table_data.append(row_p)

                chunk_size = 15
                for i in range(0, len(table_data), chunk_size):
                    chunk = table_data[i:i + chunk_size]
                    if i > 0:
                        chunk = [table_data[0]] + chunk
                    t = Table(chunk, colWidths=[col_width]*num_cols, repeatRows=1)
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), C_RED),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        ('INNERGRID', (0, 0), (-1, -1), 0.5, C_BORDER),
                        ('BOX', (0, 0), (-1, -1), 0.5, C_BORDER),
                        ('TOPPADDING', (0, 0), (-1, -1), 3),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                        ('LEFTPADDING', (0, 0), (-1, -1), 3),
                        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, C_LIGHT]),
                    ]))
                    story.append(t)
                    story.append(Spacer(1, 8))

    doc_pdf = SimpleDocTemplate(
        str(pdf_path),
        pagesize=letter,
        leftMargin=54, rightMargin=54,
        topMargin=54, bottomMargin=54
    )
    doc_pdf.build(story, canvasmaker=NumberedCanvas)
    print(f"Built PDF: {pdf_path.name} ({pdf_path.stat().st_size:,} bytes)")


if __name__ == "__main__":
    docx = DOCS / "p107_Project_Document.docx"
    pdf  = DOCS / "p107_Project_Document.pdf"

    print("Converting DOCX → PDF with embedded ABES logo...")
    convert_docx_to_pdf(docx, pdf)

    dest = SUB_DIR / "1_Project_Document.pdf"
    import shutil
    shutil.copy(pdf, dest)
    print(f"Copied to: {dest} ({dest.stat().st_size:,} bytes)")
    print("Done!")
