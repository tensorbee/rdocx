# area: sections
# needs: #304
# #304 then raises naming document.update_section, the rdocx way to change page geometry
from rdocx import Document
from rdocx.shared import Inches

doc = Document()
for section in doc.sections:
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert 'w:top="720"' in xml and 'w:left="1080"' in xml
