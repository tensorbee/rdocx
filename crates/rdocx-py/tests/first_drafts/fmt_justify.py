# area: formatting
# expected: pass
from rdocx import Document
from rdocx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()
p = doc.add_paragraph('A long justified paragraph. ' * 10)
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
doc.save('out.docx')
# --- check
assert 'w:jc w:val="both"' in part('out.docx')
