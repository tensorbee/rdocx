# area: docs
# expected: pass
# python-pptx getting started, extract all text from slides
from rpptx import Presentation

prs = Presentation('template.pptx')

text_runs = []

for slide in prs.slides:
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text_runs.append(run.text)

with open('out.txt', 'w') as f:
    f.write('\n'.join(text_runs))
# --- check
content = open('out.txt').read()
assert '{{title}}' in content and 'Agenda' in content
