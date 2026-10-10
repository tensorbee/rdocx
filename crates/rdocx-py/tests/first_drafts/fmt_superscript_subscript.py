# area: formatting
# needs: #320
from rdocx import Document

doc = Document()
p = doc.add_paragraph('E = mc')
p.add_run('2').font.superscript = True
p.add_run(' and H')
p.add_run('2').font.subscript = True
p.add_run('O')
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert 'w:vertAlign w:val="superscript"' in xml and 'w:vertAlign w:val="subscript"' in xml
