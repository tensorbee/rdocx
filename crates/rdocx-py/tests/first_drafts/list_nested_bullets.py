# area: lists
# expected: pass
from rdocx import Document

doc = Document()
doc.add_paragraph('Fruit', style='List Bullet')
doc.add_paragraph('Apple', style='List Bullet 2')
doc.add_paragraph('Pear', style='List Bullet 2')
doc.add_paragraph('Vegetables', style='List Bullet')
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.count('w:pStyle w:val="ListBullet2"') == 2 and xml.count('w:pStyle w:val="ListBullet"') == 2
styles = part('out.docx', 'word/styles.xml')
bullet2 = styles[styles.index('w:styleId="ListBullet2"'):]
assert '<w:numPr>' in bullet2[:bullet2.index('</w:style>')]
