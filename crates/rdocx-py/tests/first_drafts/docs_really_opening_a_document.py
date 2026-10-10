# area: docs
# expected: pass
# python-docx documents.rst "REALLY opening a document"
from rdocx import Document

document = Document('existing-document-file.docx')
document.save('new-file-name.docx')
# --- check
assert 'existing' in part('new-file-name.docx')
