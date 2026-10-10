# area: pictures
# needs: #322
from rdocx import Document
from rdocx.shared import Inches

doc = Document()
doc.add_picture('logo.png', width=Inches(1.25))
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert 'cx="1143000"' in xml and 'cy="1143000"' in xml
assert any(n.startswith('word/media/') for n in names('out.docx'))
