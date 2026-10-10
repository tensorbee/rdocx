# area: notes
# expected: pass
# rpptx's own notes property, as an agent finds it in the rpptx docs
from rpptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.notes_text = 'Plain notes'
prs.save('out.pptx')
# --- check
notes = [part('out.pptx', n) for n in names('out.pptx') if n.startswith('ppt/notesSlides/notesSlide')]
assert any('Plain notes' in n for n in notes)
