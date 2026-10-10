# area: lists
# expected: pass
from rdocx import Document

doc = Document()
doc.add_paragraph('Shopping list:')
for item in ['Eggs', 'Milk', 'Bread']:
    doc.add_paragraph(item, style='List Bullet')
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.count('w:pStyle w:val="ListBullet"') == 3
numbering = part('out.docx', 'word/numbering.xml')
assert 'w:numFmt w:val="bullet"' in numbering
styles = part('out.docx', 'word/styles.xml')
assert 'w:styleId="ListBullet"' in styles and '<w:numId' in styles
