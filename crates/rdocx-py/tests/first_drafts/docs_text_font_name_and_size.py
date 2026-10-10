# area: docs
# expected: pass
# python-docx text.rst, font name and size
from rdocx import Document
from rdocx.shared import Pt

document = Document()
run = document.add_paragraph('').add_run('sized')
font = run.font
font.name = 'Calibri'
font.size = Pt(12)
document.save('out.docx')
# --- check
xml = part('out.docx')
assert 'w:ascii="Calibri"' in xml and 'w:sz w:val="24"' in xml
