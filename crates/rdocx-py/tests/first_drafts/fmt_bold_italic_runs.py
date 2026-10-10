# area: formatting
# expected: pass
from rdocx import Document

doc = Document()
p = doc.add_paragraph('Plain, ')
p.add_run('bold').bold = True
p.add_run(', and ')
p.add_run('italic').italic = True
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert '<w:b/>' in xml and '<w:i/>' in xml
