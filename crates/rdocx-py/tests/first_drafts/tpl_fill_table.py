# area: templates
# needs: #322
from rdocx import Document

doc = Document('template.docx')
table = doc.tables[0]
for row in table.rows[1:]:
    for cell in row.cells:
        if cell.text == '{{value}}':
            cell.text = '42'
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.count('>42<') == 2 and '{{value}}' not in xml
