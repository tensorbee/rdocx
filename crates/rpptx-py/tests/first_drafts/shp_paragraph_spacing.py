# area: shapes
# needs: #326
from rpptx import Presentation
from rpptx.util import Inches, Pt

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
tf = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(2)).text_frame
tf.text = 'First'
p = tf.add_paragraph()
p.text = 'Second'
p.space_before = Pt(12)
p.line_spacing = 1.5
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert '<a:spcPts val="1200"/>' in xml and '<a:spcPct val="150000"/>' in xml
