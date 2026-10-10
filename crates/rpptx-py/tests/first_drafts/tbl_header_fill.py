# area: tables
# expected: pass
from rpptx import Presentation
from rpptx.dml.color import RGBColor
from rpptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
table = slide.shapes.add_table(3, 3, Inches(1), Inches(1), Inches(6), Inches(1.5)).table
for col, heading in enumerate(['Region', 'Q1', 'Q2']):
    cell = table.cell(0, col)
    cell.text = heading
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(0x2F, 0x54, 0x96)
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert xml.count('srgbClr val="2F5496"') == 3 and 'Region' in xml
