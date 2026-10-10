# area: notes
# expected: pass
from rpptx import Presentation

prs = Presentation('template.pptx')
lines = []
for index, slide in enumerate(prs.slides):
    if slide.has_notes_slide:
        lines.append(f'{index}: {slide.notes_slide.notes_text_frame.text}')
with open('out.txt', 'w') as f:
    f.write('\n'.join(lines))
# --- check
assert 'Speaker notes here' in open('out.txt').read()
