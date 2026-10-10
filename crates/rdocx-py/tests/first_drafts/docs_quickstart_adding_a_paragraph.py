# area: docs
# expected: pass
# python-docx quickstart.rst "Adding a paragraph"
from rdocx import Document

document = Document()
paragraph = document.add_paragraph('Lorem ipsum dolor sit amet.')
document.save('out.docx')
# --- check
assert 'Lorem ipsum dolor sit amet.' in part('out.docx')
