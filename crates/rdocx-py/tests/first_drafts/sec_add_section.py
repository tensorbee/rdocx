# area: sections
# needs: #304
from rdocx import Document
from rdocx.enum.section import WD_SECTION

doc = Document()
doc.add_paragraph('Portrait part')
section = doc.add_section(WD_SECTION.NEW_PAGE)
doc.add_paragraph('Second section')
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.count('<w:sectPr') == 2 and xml.count('<w:pgSz') == 2
