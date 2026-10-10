# area: shapes
# expected: pass
from rpptx import Presentation
from rpptx.enum.shapes import MSO_CONNECTOR
from rpptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(1), Inches(1), Inches(4), Inches(2))
line.line.width = Inches(0.02)
prs.save('out.pptx')
# --- check
assert '<p:cxnSp>' in part('out.pptx')
