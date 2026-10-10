# area: notes
# needs: #304
from rdocx import Document

doc = Document()
p = doc.add_paragraph('A claim that needs a source.')
doc.add_footnote(p, 'Smith, 2020.')
doc.save('out.docx')
# --- check
assert 'Smith, 2020.' in part('out.docx', 'word/footnotes.xml')
assert '<w:footnoteReference' in part('out.docx')
