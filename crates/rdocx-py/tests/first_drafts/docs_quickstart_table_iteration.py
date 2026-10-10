# area: docs
# needs: #322
# python-docx quickstart.rst, iterating a table
from rdocx import Document

document = Document()
table = document.add_table(rows=2, cols=2)
for r, values in enumerate((('alpha', 'beta'), ('gamma', 'delta'))):
    for c, value in enumerate(values):
        document.tables[0].cell(r, c).text = value
seen = []
for row in table.rows:
    for cell in row.cells:
        seen.append(cell.text)
assert seen == ['alpha', 'beta', 'gamma', 'delta'], seen
document.save('out.docx')
# --- check
assert 'delta' in part('out.docx')
