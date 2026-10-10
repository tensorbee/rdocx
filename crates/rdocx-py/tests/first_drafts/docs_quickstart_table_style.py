# area: docs
# expected: pass
# python-docx quickstart.rst, table style
from rdocx import Document

document = Document()
table = document.add_table(rows=2, cols=2)
table.style = 'LightShading-Accent1'
document.save('out.docx')
# --- check
assert 'w:tblStyle w:val="LightShading-Accent1"' in part('out.docx')
