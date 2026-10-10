# area: slides
# divergence: slides.remove
from rpptx import Presentation

prs = Presentation()
for _ in range(3):
    prs.slides.add_slide(prs.slide_layouts[6])
xml_slides = prs.slides._sldIdLst
slides = list(xml_slides)
xml_slides.remove(slides[1])
prs.save('out.pptx')
# --- check
assert part('out.pptx', 'ppt/presentation.xml').count('<p:sldId ') == 2
