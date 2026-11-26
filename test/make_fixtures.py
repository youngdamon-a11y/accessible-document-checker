"""Generate deliberately inaccessible sample documents for tests and demos.

Run: python3 test/make_fixtures.py   (needs python-docx, python-pptx, reportlab, pillow)
"""
import io, os, re, zipfile
from PIL import Image, ImageDraw
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pptx import Presentation
from pptx.util import Inches, Pt as PPt
from pptx.dml.color import RGBColor as PRGB
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

HERE = os.path.join(os.path.dirname(__file__), 'fixtures')
os.makedirs(HERE, exist_ok=True)

def chart_png():
    img = Image.new('RGB', (480, 300), 'white')
    d = ImageDraw.Draw(img)
    for i, h in enumerate([120, 200, 90, 240]):
        d.rectangle([60 + i * 100, 280 - h, 120 + i * 100, 280], fill=(35, 8, 90))
    d.line([40, 280, 460, 280], fill='black', width=2)
    d.text((150, 20), 'Students per region 2026', fill='black')
    b = io.BytesIO(); img.save(b, 'PNG'); return b.getvalue()

def add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    h = OxmlElement('w:hyperlink'); h.set(qn('r:id'), r_id)
    r = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = text; r.append(t); h.append(r)
    paragraph._p.append(h)

def make_docx():
    d = Document()
    d.core_properties.title = ''
    d.add_heading('Annual English Programme Report', 1)
    d.add_paragraph('This report summarises the English programme across all regions for 2026. ' * 3)
    d.add_heading('Regional results', 3)  # skipped level
    d.add_paragraph('Results improved in every region compared with last year.')
    d.add_picture(io.BytesIO(chart_png()))
    p = d.add_paragraph()
    r = p.add_run('Key findings')
    r.bold = True
    d.add_paragraph('Teachers reported higher engagement and students achieved better exam results than in previous years.')
    low = d.add_paragraph().add_run('This light grey note is hard to read for many people.')
    low.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
    d.add_paragraph('• Increase teacher training')
    d.add_paragraph('• Expand online courses')
    d.add_paragraph('• Review assessment methods')
    t = d.add_table(rows=3, cols=3)
    for i, row in enumerate([('Region', 'Students', 'Pass rate'), ('Europe', '1200', '92%'), ('Asia', '3400', '88%')]):
        for j, v in enumerate(row): t.cell(i, j).text = v
    lp = d.add_paragraph('For the full data set ')
    add_hyperlink(lp, 'https://example.org/reports/english-programme-data-2026', 'click here')
    for _ in range(4): d.add_paragraph('')
    d.add_paragraph('End of report.')
    path = os.path.join(HERE, 'sample-inaccessible.docx'); d.save(path)
    # strip the default language from styles.xml so the language check fires
    strip_part(path, 'word/styles.xml', lambda s: re.sub(r'<w:lang [^>]*/>', '', s))

def strip_part(path, part, fn):
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    items[part] = fn(items[part].decode('utf-8')).encode('utf-8')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, b in items.items(): z.writestr(n, b)

def make_pptx():
    prs = Presentation()
    s1 = prs.slides.add_slide(prs.slide_layouts[0])
    s1.shapes.title.text = 'Programme Update'
    s1.placeholders[1].text = 'November 2026'
    s2 = prs.slides.add_slide(prs.slide_layouts[6])  # blank: no title
    tb = s2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
    tb.text_frame.text = 'Our results this year'
    r = tb.text_frame.paragraphs[0].runs[0]; r.font.color.rgb = PRGB(0xBB, 0xBB, 0xBB); r.font.size = PPt(16)
    s2.shapes.add_picture(io.BytesIO(chart_png()), Inches(1), Inches(2), Inches(5))
    s3 = prs.slides.add_slide(prs.slide_layouts[5])
    s3.shapes.title.text = 'Programme Update'  # duplicate title
    rows = s3.shapes.add_table(3, 3, Inches(1), Inches(2), Inches(6), Inches(2)).table
    for i, row in enumerate([('Region', 'Students', 'Pass rate'), ('Europe', '1200', '92%'), ('Asia', '3400', '88%')]):
        for j, v in enumerate(row): rows.cell(i, j).text = v
    tbl = s3.shapes[-1]._element.graphic.graphicData.tbl
    tbl.tblPr.set('firstRow', '0')
    s4 = prs.slides.add_slide(prs.slide_layouts[5])
    s4.shapes.title.text = 'Find out more'
    box = s4.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1))
    run = box.text_frame.paragraphs[0].add_run(); run.text = 'https://example.org/programme'
    run.hyperlink.address = 'https://example.org/programme'
    prs.core_properties.title = ''
    path = os.path.join(HERE, 'sample-inaccessible.pptx'); prs.save(path)
    strip_part(path, 'ppt/presentation.xml', lambda s: re.sub(r' lang="[^"]*"', '', s))
    for n in ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/slides/slide3.xml', 'ppt/slides/slide4.xml']:
        strip_part(path, n, lambda s: re.sub(r' lang="[^"]*"', '', s))

def make_pdf():
    path = os.path.join(HERE, 'sample-untagged.pdf')
    c = canvas.Canvas(path, pagesize=A4)
    c.setTitle('')
    for page in range(3):
        c.setFont('Helvetica-Bold', 18); c.drawString(72, 770, f'Section {page + 1}')
        c.setFont('Helvetica', 11); c.drawString(72, 740, 'This PDF has no tags, no language and no title.')
        c.linkURL('https://example.org', (72, 700, 300, 715)); c.drawString(72, 703, 'https://example.org')
        c.showPage()
    c.save()

if __name__ == '__main__':
    make_docx(); make_pptx(); make_pdf()
    print('fixtures written to', HERE)
