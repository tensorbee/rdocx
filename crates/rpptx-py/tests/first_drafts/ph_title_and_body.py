# area: placeholders
# needs: #326
from rpptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = 'Agenda'
body = slide.placeholders[1]
body.text_frame.text = 'Welcome'
for item in ['Results', 'Questions']:
    p = body.text_frame.add_paragraph()
    p.text = item
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert 'Welcome' in xml and 'Questions' in xml and 'Agenda' in xml
