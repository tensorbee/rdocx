# area: docs
# expected: pass
# python-docx text.rst, line spacing
from rdocx import Document
from rdocx.shared import Pt

document = Document()
paragraph_format = document.add_paragraph('spaced').paragraph_format
paragraph_format.line_spacing = Pt(18)
paragraph_format.line_spacing = 1.75
document.save('out.docx')
# --- check
assert re.search(r'w:line="420"', part('out.docx'))
