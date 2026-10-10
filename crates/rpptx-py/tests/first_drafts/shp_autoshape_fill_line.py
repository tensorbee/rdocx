# area: shapes
# expected: pass
from rpptx import Presentation
from rpptx.dml.color import RGBColor
from rpptx.enum.shapes import MSO_SHAPE
from rpptx.util import Inches, Pt

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1.5))
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(0x00, 0x70, 0xC0)
shape.line.color.rgb = RGBColor(0x00, 0x20, 0x60)
shape.line.width = Pt(2)
shape.text = 'Box'
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert 'srgbClr val="0070C0"' in xml and '<a:ln w="25400"' in xml and 'srgbClr val="002060"' in xml
