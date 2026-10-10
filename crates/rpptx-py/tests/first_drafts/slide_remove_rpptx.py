# area: slides
# expected: pass
# rpptx's own slide removal, as an agent finds it in the rpptx docs
from rpptx import Presentation

prs = Presentation()
for title in ['keep', 'drop', 'keep too']:
    prs.slides.add_slide(prs.slide_layouts[5]).shapes.title.text = title
prs.slides.remove(prs.slides[1])
prs.save('out.pptx')
# --- check
assert part('out.pptx', 'ppt/presentation.xml').count('<p:sldId ') == 2
