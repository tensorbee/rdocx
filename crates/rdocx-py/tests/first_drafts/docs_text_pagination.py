# area: docs
# expected: pass
# python-docx text.rst, pagination properties
from rdocx import Document

document = Document()
paragraph_format = document.add_paragraph('kept').paragraph_format
paragraph_format.keep_with_next = True
paragraph_format.page_break_before = False
document.save('out.docx')
# --- check
xml = part('out.docx')
assert '<w:keepNext/>' in xml and '<w:pageBreakBefore/>' not in xml
