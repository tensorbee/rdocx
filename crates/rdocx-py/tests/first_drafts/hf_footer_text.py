# area: headers-footers
# needs: #304
from rdocx import Document

doc = Document()
footer = doc.sections[0].footer
p = footer.paragraphs[0]
p.text = 'Confidential'
doc.save('out.docx')
# --- check
footers = [part('out.docx', n) for n in names('out.docx') if n.startswith('word/footer')]
assert any('Confidential' in f for f in footers)
