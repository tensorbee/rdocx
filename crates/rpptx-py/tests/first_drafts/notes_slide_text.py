# area: notes
# needs: #326
from rpptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = 'With notes'
notes_slide = slide.notes_slide
text_frame = notes_slide.notes_text_frame
text_frame.text = 'Remember to smile.'
prs.save('out.pptx')
# --- check
notes = [part('out.pptx', n) for n in names('out.pptx') if re.match(r'ppt/notesSlides/notesSlide\d+\.xml$', n)]
assert any('Remember to smile.' in n for n in notes)
