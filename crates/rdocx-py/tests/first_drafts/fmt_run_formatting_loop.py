# area: formatting
# expected: pass
from rdocx import Document
from rdocx.shared import Pt

doc = Document()
p = doc.add_paragraph()
for word in ['alpha', 'beta', 'gamma']:
    run = p.add_run(word + ' ')
    run.bold = word == 'beta'
    run.font.size = Pt(12)
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.count('w:sz w:val="24"') == 3 and xml.count('<w:b/>') == 1
