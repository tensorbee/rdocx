# area: tables
# needs: #322
from rdocx import Document

data = [['Name', 'Score'], ['Ann', '91'], ['Bob', '78']]
doc = Document()
table = doc.add_table(rows=len(data), cols=2)
for i, row in enumerate(data):
    for j, value in enumerate(row):
        table.cell(i, j).text = value
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert all(v in xml for row in data for v in row)
