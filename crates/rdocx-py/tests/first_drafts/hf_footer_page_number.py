# area: headers-footers
# needs: #304
from rdocx import Document

doc = Document()
footer = doc.sections[0].footer
footer.add_page_number()
doc.save('out.docx')
# --- check
footers = [part('out.docx', n) for n in names('out.docx') if n.startswith('word/footer')]
assert any('PAGE' in f for f in footers)
