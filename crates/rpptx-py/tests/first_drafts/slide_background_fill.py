# area: slides
# expected: pass
from rpptx import Presentation
from rpptx.dml.color import RGBColor

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
fill = slide.background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(0x1F, 0x3B, 0x5C)
prs.save('out.pptx')
# --- check
assert re.search(r'<p:bg>.*srgbClr val="1F3B5C"', part('out.pptx'), re.S)
