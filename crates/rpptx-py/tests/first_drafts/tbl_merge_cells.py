# area: tables
# expected: pass
from rpptx import Presentation
from rpptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
table = slide.shapes.add_table(2, 3, Inches(1), Inches(1), Inches(6), Inches(1)).table
cell = table.cell(0, 0)
cell.merge(table.cell(0, 2))
cell.text = 'Merged title'
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert 'gridSpan="3"' in xml and 'Merged title' in xml
