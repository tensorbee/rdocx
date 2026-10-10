# area: document
# divergence: remove_content
from rdocx import Document

doc = Document()
doc.add_paragraph('Keep')
doc.add_paragraph('DELETE ME')
for paragraph in doc.paragraphs:
    if paragraph.text == 'DELETE ME':
        paragraph._element.getparent().remove(paragraph._element)
doc.save('out.docx')
# --- check
assert 'DELETE ME' not in part('out.docx')
