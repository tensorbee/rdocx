# area: placeholders
# needs: #326
from rpptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[8])
placeholder = slide.placeholders[1]
picture = placeholder.insert_picture('logo.png')
prs.save('out.pptx')
# --- check
assert '<p:pic>' in part('out.pptx')
