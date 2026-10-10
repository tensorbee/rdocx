# area: lists
# needs: #320
# rdocx's own bullet helper, as an agent finds it in the rdocx docs
from rdocx import Document

doc = Document()
doc.add_bullet_list_item('Top level')
doc.add_bullet_list_item('Nested', level=1)
doc.add_numbered_list_item('Step one')
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.count('<w:numPr>') == 3 and 'w:ilvl w:val="1"' in xml
