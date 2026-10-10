# area: headers-footers
# needs: #304
from rdocx import Document

doc = Document()
section = doc.sections[0]
header = section.header
header.paragraphs[0].text = 'Quarterly report'
doc.add_paragraph('Body')
doc.save('out.docx')
# --- check
assert any(n.startswith('word/header') for n in names('out.docx'))
headers = [part('out.docx', n) for n in names('out.docx') if n.startswith('word/header')]
assert any('Quarterly report' in h for h in headers)
