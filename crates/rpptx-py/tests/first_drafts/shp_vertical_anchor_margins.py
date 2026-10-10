# area: shapes
# expected: pass
from rpptx import Presentation
from rpptx.enum.text import MSO_ANCHOR
from rpptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
tf = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(2)).text_frame
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
tf.margin_left = Inches(0.2)
tf.text = 'Middle'
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert 'anchor="ctr"' in xml and 'lIns="182880"' in xml
