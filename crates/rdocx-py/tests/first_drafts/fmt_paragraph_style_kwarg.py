# area: formatting
# expected: pass
from rdocx import Document

doc = Document()
doc.add_paragraph('To be or not to be.', style='Quote')
doc.add_paragraph('Introduction', style='Heading 1')
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert 'w:pStyle w:val="Quote"' in xml and 'w:pStyle w:val="Heading1"' in xml
