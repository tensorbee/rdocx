# area: tables
# needs: #317
from rpptx import Presentation
from rpptx.util import Inches, Pt

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
table = slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(4), Inches(1)).table
cell = table.cell(0, 0)
cell.text = 'Bold header'
para = cell.text_frame.paragraphs[0]
para.font.bold = True
para.font.size = Pt(14)
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert 'b="1"' in xml and 'sz="1400"' in xml
