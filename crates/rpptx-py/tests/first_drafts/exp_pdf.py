# area: export
# expected: pass
from rpptx import Presentation

prs = Presentation()
prs.slides.add_slide(prs.slide_layouts[0]).shapes.title.text = 'Exported'
with open('deck.pdf', 'wb') as f:
    f.write(prs.to_pdf())
# --- check
assert open('deck.pdf', 'rb').read(5) == b'%PDF-'
