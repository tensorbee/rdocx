# area: placeholders
# needs: #324, #326
from rpptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[1])
found = []
for shape in slide.placeholders:
    found.append((shape.placeholder_format.idx, shape.name))
    shape.text = f'placeholder {shape.placeholder_format.idx}'
prs.save('out.pptx')
# --- check
assert 'placeholder 0' in part('out.pptx') and 'placeholder 1' in part('out.pptx')
