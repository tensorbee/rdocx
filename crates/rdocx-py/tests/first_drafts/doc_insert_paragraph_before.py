# area: document
# expected: pass
from rdocx import Document

doc = Document()
second = doc.add_paragraph('Second')
second.insert_paragraph_before('First')
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.index('First') < xml.index('Second')
