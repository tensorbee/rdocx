# area: templates
# needs: #326
from rpptx import Presentation

prs = Presentation('template.pptx')
for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame and '{{title}}' in shape.text_frame.text:
            shape.text_frame.text = shape.text_frame.text.replace('{{title}}', 'Board update')
prs.save('out.pptx')
# --- check
assert 'Board update' in part('out.pptx')
