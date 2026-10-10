# area: shapes
# needs: #317
from rpptx import Presentation
from rpptx.dml.color import RGBColor
from rpptx.util import Inches, Pt

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
tf = txBox.text_frame
p = tf.paragraphs[0]
run = p.add_run()
run.text = 'Styled run'
font = run.font
font.name = 'Calibri'
font.size = Pt(24)
font.bold = True
font.color.rgb = RGBColor(0xFF, 0x7F, 0x50)
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert 'sz="2400"' in xml and 'srgbClr val="FF7F50"' in xml and 'typeface="Calibri"' in xml
