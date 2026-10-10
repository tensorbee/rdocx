# area: formatting
# needs: #320
from rdocx import Document

doc = Document()
p = doc.add_paragraph('')
p.add_run('all caps').font.all_caps = True
p.add_run(' small caps').font.small_caps = True
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert '<w:caps/>' in xml and '<w:smallCaps/>' in xml
