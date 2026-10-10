# area: shapes
# needs: #326
from rpptx import Presentation
from rpptx.enum.text import PP_ALIGN
from rpptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
box.text_frame.text = 'Centered title'
box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
prs.save('out.pptx')
# --- check
assert 'algn="ctr"' in part('out.pptx')
