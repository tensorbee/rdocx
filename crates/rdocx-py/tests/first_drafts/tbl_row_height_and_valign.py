# area: tables
# expected: pass
from rdocx import Document
from rdocx.enum.table import WD_ALIGN_VERTICAL
from rdocx.shared import Inches

doc = Document()
table = doc.add_table(rows=1, cols=1)
table.rows[0].height = Inches(0.5)
table.cell(0, 0).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert 'w:val="720"' in xml and 'w:vAlign w:val="center"' in xml
