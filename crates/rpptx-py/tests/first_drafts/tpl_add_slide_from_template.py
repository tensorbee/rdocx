# area: templates
# needs: #326
from rpptx import Presentation

prs = Presentation('template.pptx')
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = 'Appended'
slide.placeholders[1].text = 'From the template layout'
prs.save('out.pptx')
# --- check
assert 'From the template layout' in part('out.pptx', 'ppt/slides/slide3.xml')
