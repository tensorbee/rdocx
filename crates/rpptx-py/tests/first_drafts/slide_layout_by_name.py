# area: slides
# expected: pass
from rpptx import Presentation

prs = Presentation()
layout = next(l for l in prs.slide_layouts if l.name == 'Title and Content')
slide = prs.slides.add_slide(layout)
slide.shapes.title.text = 'Chosen by name'
prs.save('out.pptx')
# --- check
assert 'Chosen by name' in part('out.pptx')
