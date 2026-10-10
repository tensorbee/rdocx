# area: tables
# expected: pass
from rdocx import Document

doc = Document()
table = doc.add_table(rows=3, cols=2)
table.rows[0].is_header = True
doc.save('out.docx')
# --- check
assert '<w:tblHeader/>' in part('out.docx')
