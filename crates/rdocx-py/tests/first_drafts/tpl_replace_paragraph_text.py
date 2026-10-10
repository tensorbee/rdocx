# area: templates
# expected: pass
from rdocx import Document

doc = Document('template.docx')
for paragraph in doc.paragraphs:
    if '{{name}}' in paragraph.text:
        paragraph.text = paragraph.text.replace('{{name}}', 'Ada Lovelace')
    if '{{date}}' in paragraph.text:
        paragraph.text = paragraph.text.replace('{{date}}', '2026-10-10')
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert 'Ada Lovelace' in xml and '2026-10-10' in xml and '{{name}}' not in xml
