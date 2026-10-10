# area: shapes
# needs: #326
from rpptx import Presentation
from rpptx.enum.shapes import MSO_SHAPE
from rpptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
group = slide.shapes.add_group_shape()
group.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(1), Inches(1))
group.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.5), Inches(1), Inches(1), Inches(1))
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert '<p:grpSp>' in xml and 'prst="ellipse"' in xml and 'prst="rect"' in xml
