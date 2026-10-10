# area: docs
# expected: pass
# python-docx text.rst, indentation
from rdocx import Document
from rdocx.shared import Inches, Pt

document = Document()
paragraph_format = document.add_paragraph('indented').paragraph_format
paragraph_format.left_indent = Inches(0.5)
paragraph_format.right_indent = Pt(24)
paragraph_format.first_line_indent = Inches(-0.25)
document.save('out.docx')
# --- check
xml = part('out.docx')
assert re.search(r'<w:ind [^>]*w:(left|start)="720"', xml), xml
assert 'w:hanging="360"' in xml
