# area: docs
# expected: pass
# python-docx documents.rst "Opening a document"
from rdocx import Document

document = Document()
document.save('test.docx')
# --- check
assert '<w:body' in part('test.docx')
