# area: docs
# expected: pass
# python-docx text.rst, horizontal alignment
from rdocx import Document
from rdocx.enum.text import WD_ALIGN_PARAGRAPH

document = Document()
paragraph = document.add_paragraph('centered')
paragraph_format = paragraph.paragraph_format
paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
document.save('out.docx')
# --- check
assert 'w:jc w:val="center"' in part('out.docx')
