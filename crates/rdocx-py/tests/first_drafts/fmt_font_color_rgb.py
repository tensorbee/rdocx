# area: formatting
# needs: #317
from rdocx import Document
from rdocx.shared import RGBColor

doc = Document()
run = doc.add_paragraph().add_run('Warning')
run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
doc.save('out.docx')
# --- check
assert 'w:color w:val="C00000"' in part('out.docx')
