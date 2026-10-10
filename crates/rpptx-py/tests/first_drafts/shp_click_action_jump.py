# area: shapes
# needs: #326
from rpptx import Presentation
from rpptx.enum.shapes import MSO_SHAPE
from rpptx.util import Inches

prs = Presentation()
first = prs.slides.add_slide(prs.slide_layouts[6])
second = prs.slides.add_slide(prs.slide_layouts[6])
button = first.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1))
button.text = 'Next'
button.click_action.target_slide = second
prs.save('out.pptx')
# --- check
assert 'ppaction://hlinksldjump' in part('out.pptx')
