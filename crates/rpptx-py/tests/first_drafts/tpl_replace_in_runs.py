# area: templates
# expected: pass
from rpptx import Presentation

replacements = {'{{title}}': 'Q3 review', '{{author}}': 'Finance'}
prs = Presentation('template.pptx')
for slide in prs.slides:
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                for key, value in replacements.items():
                    if key in run.text:
                        run.text = run.text.replace(key, value)
prs.save('out.pptx')
# --- check
xml = part('out.pptx')
assert 'Q3 review' in xml and 'Prepared by Finance' in xml and '{{' not in xml
