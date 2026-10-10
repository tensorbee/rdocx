# area: templates
# expected: pass
from rdocx import Document

doc = Document('template.docx')
doc.core_properties.title = 'Offer letter'
doc.core_properties.author = 'HR'
doc.save('out.docx')
# --- check
core = part('out.docx', 'docProps/core.xml')
assert 'Offer letter' in core and '>HR<' in core
