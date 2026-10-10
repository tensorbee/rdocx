# area: templates
# expected: pass
from rdocx import Document

doc = Document('template.docx')
for paragraph in doc.paragraphs:
    for run in paragraph.runs:
        if '{{name}}' in run.text:
            run.text = run.text.replace('{{name}}', 'Grace Hopper')
doc.save('out.docx')
# --- check
assert 'Grace Hopper' in part('out.docx')
