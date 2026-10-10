# area: formatting
# expected: pass
from rdocx import Document

doc = Document()
p = doc.add_paragraph('Chapter one')
p.style = 'Heading 2'
doc.save('out.docx')
# --- check
assert 'w:pStyle w:val="Heading2"' in part('out.docx')
