# area: lists
# expected: pass
from rdocx import Document

doc = Document()
for step in ['Open the box', 'Remove the device', 'Plug it in']:
    doc.add_paragraph(step, style='List Number')
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.count('w:pStyle w:val="ListNumber"') == 3
assert 'w:numFmt w:val="decimal"' in part('out.docx', 'word/numbering.xml')
