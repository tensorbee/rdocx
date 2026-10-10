# area: slides
# expected: pass
from rpptx import Presentation

prs = Presentation()
for heading in ['Introduction', 'Results', 'Next steps']:
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = heading
prs.save('out.pptx')
# --- check
assert len([n for n in names('out.pptx') if re.match(r'ppt/slides/slide\d+\.xml$', n)]) == 3
assert 'Next steps' in part('out.pptx', 'ppt/slides/slide3.xml')
