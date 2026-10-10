# area: tables
# needs: #322
from rdocx import Document
from rdocx.enum.table import WD_TABLE_ALIGNMENT

doc = Document()
table = doc.add_table(rows=2, cols=2, style='Table Grid')
table.alignment = WD_TABLE_ALIGNMENT.CENTER
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert 'w:tblStyle w:val="TableGrid"' in xml and re.search(r'<w:tblPr>.*<w:jc w:val="center"/>', xml)
