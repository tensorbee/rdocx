# area: docs
# needs: #326
# python-pptx getting started, add_textbox() example
from rpptx import Presentation
from rpptx.util import Inches, Pt

prs = Presentation()
blank_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_slide_layout)

left = top = width = height = Inches(1)
txBox = slide.shapes.add_textbox(left, top, width, height)
tf = txBox.text_frame

tf.text = "This is text inside a textbox"

p = tf.add_paragraph()
p.text = "This is a second paragraph that's bold"
p.font.bold = True

p = tf.add_paragraph()
p.text = "This is a third paragraph that's big"
p.font.size = Pt(40)

prs.save('test.pptx')
# --- check
xml = part('test.pptx')
assert 'b="1"' in xml and 'sz="4000"' in xml and 'inside a textbox' in xml
