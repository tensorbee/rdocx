# area: shapes
# expected: pass
from rpptx import Presentation
from rpptx.enum.shapes import MSO_SHAPE
from rpptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(1), Inches(1), Inches(2), Inches(1))
arrow.rotation = 45
arrow.name = 'Arrow'
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert 'rot="2700000"' in xml and 'name="Arrow"' in xml
