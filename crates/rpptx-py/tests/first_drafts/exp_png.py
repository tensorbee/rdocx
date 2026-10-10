# area: export
# expected: pass
from rpptx import Presentation

prs = Presentation()
prs.slides.add_slide(prs.slide_layouts[0]).shapes.title.text = 'Thumbnail'
png = prs.render_slide_to_png(0, dpi=48)
with open('slide1.png', 'wb') as f:
    f.write(png)
# --- check
assert open('slide1.png', 'rb').read(8) == b'\x89PNG\r\n\x1a\n'
