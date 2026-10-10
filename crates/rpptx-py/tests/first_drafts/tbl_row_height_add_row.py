# area: tables
# expected: pass
from rpptx import Presentation
from rpptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
table = slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(4), Inches(1)).table
table.rows[0].height = Inches(0.6)
row = table.rows.add_row()
row.cells[0].text = 'added'
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert xml.count('<a:tr ') == 3 and 'h="548640"' in xml and 'added' in xml
