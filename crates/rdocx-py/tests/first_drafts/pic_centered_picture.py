# area: pictures
# needs: #322
from rdocx import Document
from rdocx.enum.text import WD_ALIGN_PARAGRAPH
from rdocx.shared import Inches

doc = Document()
doc.add_picture('logo.png', width=Inches(2))
last_paragraph = doc.paragraphs[-1]
last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert re.search(r'<w:jc w:val="center"/>.*<wp:inline', xml, re.S)
