# area: sections
# expected: pass
from rdocx import Document

doc = Document()
doc.add_paragraph('Page one')
doc.add_page_break()
doc.add_paragraph('Page two')
doc.save('out.docx')
# --- check
assert '<w:br w:type="page"/>' in part('out.docx')
