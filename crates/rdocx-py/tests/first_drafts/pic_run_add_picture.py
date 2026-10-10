# area: pictures
# needs: #317, #322
from rdocx import Document
from rdocx.shared import Inches

doc = Document()
paragraph = doc.add_paragraph()
run = paragraph.add_run()
run.add_picture('logo.png', width=Inches(1))
paragraph.add_run(' Company logo')
doc.save('out.docx')
# --- check
xml = part('out.docx')
assert '<wp:inline' in xml and 'Company logo' in xml
