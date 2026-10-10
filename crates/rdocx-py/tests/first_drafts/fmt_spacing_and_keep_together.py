# area: formatting
# expected: pass
from rdocx import Document
from rdocx.shared import Pt

doc = Document()
p = doc.add_paragraph('Spaced and kept together.')
pf = p.paragraph_format
pf.space_after = Pt(6)
pf.line_spacing = 1.5
pf.keep_together = True
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert 'w:after="120"' in xml and 'w:line="360"' in xml and '<w:keepLines/>' in xml
