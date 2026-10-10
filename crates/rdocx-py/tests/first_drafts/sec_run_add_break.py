# area: sections
# needs: #304
from rdocx import Document
from rdocx.enum.text import WD_BREAK

doc = Document()
run = doc.add_paragraph().add_run('Before the break')
run.add_break(WD_BREAK.PAGE)
doc.save('out.docx')
# --- check
assert '<w:br w:type="page"/>' in part('out.docx')
