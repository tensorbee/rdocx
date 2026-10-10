# area: formatting
# expected: pass
from rdocx import Document

doc = Document()
p = doc.add_paragraph('')
r1 = p.add_run('underlined')
r1.font.underline = True
r2 = p.add_run(' struck')
r2.font.strike = True
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert 'w:u w:val="single"' in xml and '<w:strike/>' in xml
