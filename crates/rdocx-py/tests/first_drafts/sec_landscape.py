# area: sections
# needs: #304
# #304 then raises naming document.update_section, the rdocx way to change page geometry
from rdocx import Document
from rdocx.enum.section import WD_ORIENT

doc = Document()
section = doc.sections[-1]
new_width, new_height = section.page_height, section.page_width
section.orientation = WD_ORIENT.LANDSCAPE
section.page_width = new_width
section.page_height = new_height
doc.save('out.docx')
# --- check
xml = part('out.docx')
m = re.search(r'<w:pgSz [^>]*>', xml).group(0)
assert 'w:orient="landscape"' in m
w = int(re.search(r'w:w="(\d+)"', m).group(1)); h = int(re.search(r'w:h="(\d+)"', m).group(1))
assert w > h
