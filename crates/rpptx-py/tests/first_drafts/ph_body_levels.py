# area: placeholders
# needs: #326
from rpptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[1])
tf = slide.placeholders[1].text_frame
tf.text = 'Top'
for level, text in [(1, 'Second'), (2, 'Third')]:
    p = tf.add_paragraph()
    p.text = text
    p.level = level
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert 'lvl="1"' in xml and 'lvl="2"' in xml
