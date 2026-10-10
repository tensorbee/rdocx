# area: shapes
# expected: pass
from rpptx import Presentation
from rpptx.enum.text import MSO_AUTO_SIZE
from rpptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
tf = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1)).text_frame
tf.word_wrap = True
tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
tf.text = 'A long sentence that wraps inside the box.'
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert 'wrap="square"' in xml and '<a:spAutoFit/>' in xml
