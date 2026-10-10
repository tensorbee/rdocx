# area: lists
# expected: pass
from rdocx import Document

doc = Document()
for text in ['first', 'second']:
    p = doc.add_paragraph(text)
    p.style = 'List Number'
doc.save('out.docx')
# --- check
assert part('out.docx').count('w:pStyle w:val="ListNumber"') == 2
