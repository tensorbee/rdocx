# area: formatting
# needs: #320
from rdocx import Document
from rdocx.shared import Inches

doc = Document()
p = doc.add_paragraph('Name\tValue')
p.paragraph_format.tab_stops.add_tab_stop(Inches(3))
doc.save('out.docx')
# --- check
assert re.search(r'<w:tab w:val="left" w:pos="4320"/>|<w:tab w:pos="4320" w:val="left"/>', part('out.docx'))
