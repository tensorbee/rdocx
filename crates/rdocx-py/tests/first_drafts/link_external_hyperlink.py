# area: links
# expected: pass
from rdocx import Document

doc = Document()
p = doc.add_paragraph('See ')
p.add_hyperlink('our site', 'https://example.com')
doc.save('out.docx')
# --- check
assert '<w:hyperlink' in part('out.docx')
assert 'Target="https://example.com"' in part('out.docx', 'word/_rels/document.xml.rels')
