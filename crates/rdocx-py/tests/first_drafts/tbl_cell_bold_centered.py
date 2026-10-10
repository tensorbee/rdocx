# area: tables
# expected: pass
from rdocx import Document
from rdocx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()
table = doc.add_table(rows=1, cols=2)
cell = table.cell(0, 0)
paragraph = cell.paragraphs[0]
paragraph.add_run('Total').bold = True
paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert '<w:b/>' in xml and 'w:jc w:val="center"' in xml and 'Total' in xml
