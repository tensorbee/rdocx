# area: notes
# needs: #304
from rdocx import Document

doc = Document()
p = doc.add_paragraph('Closing remark.')
run = p.runs[0]
doc.add_endnote(run, 'See the appendix.')
doc.save('out.docx')
# --- check
assert 'See the appendix.' in part('out.docx', 'word/endnotes.xml')
