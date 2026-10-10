# area: docs
# expected: pass
# python-docx quickstart.rst "Adding a table"
from rdocx import Document

document = Document()
table = document.add_table(rows=2, cols=2)
document.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.count('<w:tr') == 2 and xml.count('<w:tc>') + xml.count('<w:tc ') == 4
