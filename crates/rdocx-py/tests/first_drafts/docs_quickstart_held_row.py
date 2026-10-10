# area: docs
# needs: #322
# python-docx quickstart.rst, two cells written through one held row
from rdocx import Document

document = Document()
table = document.add_table(rows=2, cols=2)
row = table.rows[1]
row.cells[0].text = 'Foo bar to you.'
row.cells[1].text = 'And a hearty foo bar to you too sir!'
document.save('out.docx')
# --- check
xml = part('out.docx')
assert 'Foo bar to you.' in xml and 'And a hearty foo bar' in xml
