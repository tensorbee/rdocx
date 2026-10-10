# area: formatting
# expected: pass
from rdocx import Document

doc = Document()
doc.add_heading('Annual report', 0)
doc.add_heading('Summary', level=1)
doc.add_paragraph('Body text.')
doc.add_heading('Details', level=2)
doc.save('out.docx')
# --- check
xml = part('out.docx')
for style in ('Title', 'Heading1', 'Heading2'):
    assert f'w:pStyle w:val="{style}"' in xml, style
