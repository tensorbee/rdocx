# area: docs
# expected: pass
# python-docx quickstart.rst "Opening a document"
from rdocx import Document

document = Document()
document.save('out.docx')
# --- check
assert '<w:body' in part('out.docx')
