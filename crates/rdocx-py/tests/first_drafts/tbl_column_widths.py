# area: tables
# needs: #322
from rdocx import Document
from rdocx.shared import Inches

doc = Document()
table = doc.add_table(rows=2, cols=2)
table.autofit = False
for cell in table.columns[0].cells:
    cell.width = Inches(1)
for cell in table.columns[1].cells:
    cell.width = Inches(4)
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert len(re.findall(r'<w:tcW [^>]*w:w="1440"', xml)) == 2
assert len(re.findall(r'<w:tcW [^>]*w:w="5760"', xml)) == 2
