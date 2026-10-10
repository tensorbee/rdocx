# area: lists
# expected: pass
from rdocx import Document

doc = Document()
doc.add_heading('Action items', level=1)
for owner, task in [('Ann', 'Draft'), ('Bob', 'Review')]:
    p = doc.add_paragraph(style='List Bullet')
    p.add_run(owner).bold = True
    p.add_run(f': {task}')
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.count('w:pStyle w:val="ListBullet"') == 2 and xml.count('<w:b/>') == 2
