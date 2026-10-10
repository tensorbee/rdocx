# area: docs
# expected: pass
# python-docx quickstart.rst, adding runs
from rdocx import Document

document = Document()
paragraph = document.add_paragraph('Lorem ipsum ')
paragraph.add_run('dolor sit amet.')
document.save('out.docx')
# --- check
xml = part('out.docx')
assert xml.count('<w:r>') + xml.count('<w:r ') >= 2 and 'dolor sit amet.' in xml
