# area: docs
# expected: pass
# python-docx quickstart.rst, cell text
from rdocx import Document

document = Document()
table = document.add_table(rows=2, cols=2)
cell = table.cell(0, 1)
cell.text = 'parrot, possibly dead'
document.save('out.docx')
# --- check
assert 'parrot, possibly dead' in part('out.docx')
