# area: templates
# expected: pass
from rdocx import Document

doc = Document('template.docx')
doc.add_heading('Appendix', level=1)
doc.add_paragraph('Added after the template content.')
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.index('Appendix') > xml.index('{{name}}')
