# area: tables
# needs: #322
from rdocx import Document

records = ((3, '101', 'Spam'), (7, '422', 'Eggs'), (4, '631', 'Spam, spam, eggs, and spam'))
doc = Document()
table = doc.add_table(rows=1, cols=3)
table.style = 'Table Grid'
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Qty'
hdr_cells[1].text = 'Id'
hdr_cells[2].text = 'Desc'
for qty, id, desc in records:
    row_cells = table.add_row().cells
    row_cells[0].text = str(qty)
    row_cells[1].text = id
    row_cells[2].text = desc
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.count('<w:tr') == 4 and 'Spam, spam, eggs, and spam' in xml and 'TableGrid' in xml
