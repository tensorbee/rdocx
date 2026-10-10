# area: formatting
# expected: pass
from rdocx import Document
from rdocx.shared import Pt

doc = Document()
p = doc.add_paragraph()
run = p.add_run('Big Arial text')
run.font.size = Pt(20)
run.font.name = 'Arial'
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert 'w:sz w:val="40"' in xml and 'w:ascii="Arial"' in xml
