# area: slides
# expected: pass
from rpptx import Presentation
from rpptx.util import Inches

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])
prs.save('out.pptx')
# --- check
assert re.search(r'<p:sldSz cx="121\d{5}" cy="6858000"', part('out.pptx', 'ppt/presentation.xml'))
