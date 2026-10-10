# area: docs
# expected: pass
# python-docx text.rst, paragraph spacing
from rdocx import Document
from rdocx.shared import Pt

document = Document()
paragraph_format = document.add_paragraph('spaced').paragraph_format
paragraph_format.space_before = Pt(18)
paragraph_format.space_after = Pt(12)
document.save('out.docx')
# --- check
xml = part('out.docx')
assert 'w:before="360"' in xml and 'w:after="240"' in xml
