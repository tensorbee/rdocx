# area: tables
# needs: #322
from rdocx import Document
from rdocx.shared import Inches

doc = Document()
table = doc.add_table(rows=2, cols=2)
column = table.add_column(Inches(1.5))
table.cell(0, 2).text = 'New'
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.count('<w:gridCol') == 3 and 'New' in xml
